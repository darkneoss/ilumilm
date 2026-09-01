"""Genera la memoria de calculo en HTML a partir de `resultados.json`.

Uso:
    python -m engine.report estudios/<nombre>/resultados.json

Escribe `reporte.html` junto al JSON. El archivo es autocontenido salvo por las
tipografias de Google Fonts, y esta pensado para publicarse como Artifact.
"""
from __future__ import annotations

import json
import sys
from html import escape
from pathlib import Path
from typing import Any, Dict, List, Sequence

from . import nom

# Rampa de la escala isolux: de la noche al vapor de sodio. Los luminarios de
# vialidad viven en ese gradiente, asi que la escala dice algo del tema en vez
# de ser un arcoiris generico.
RAMPA = [
    (0.00, (0x16, 0x23, 0x3A)),
    (0.25, (0x2E, 0x4B, 0x7A)),
    (0.50, (0x7A, 0x63, 0x62)),
    (0.75, (0xB5, 0x76, 0x2A)),
    (0.90, (0xE8, 0xB4, 0x5C)),
    (1.00, (0xFB, 0xEB, 0xC8)),
]


def _color(t: float) -> str:
    t = max(0.0, min(1.0, t))
    for (t0, c0), (t1, c1) in zip(RAMPA, RAMPA[1:]):
        if t <= t1:
            f = (t - t0) / (t1 - t0) if t1 > t0 else 0.0
            r, g, b = (round(a + (b_ - a) * f) for a, b_ in zip(c0, c1))
            return "#{:02x}{:02x}{:02x}".format(r, g, b)
    return "#fbebc8"


def _num(x: float, dec: int = 2) -> str:
    if x == float("inf"):
        return "&infin;"
    return "{:,.{}f}".format(x, dec).replace(",", " ")


def _chip(criterio: Dict[str, Any], unidad: str, dec: int = 2) -> str:
    ok = criterio["cumple"]
    return (
        '<div class="criterio {cls}">'
        '<span class="criterio-nombre">{nombre}</span>'
        '<span class="criterio-valor">{valor}<span class="u">{u}</span></span>'
        '<span class="criterio-limite">{comp} {limite}</span>'
        "</div>"
    ).format(
        cls="ok" if ok else "bad",
        nombre=escape(criterio["nombre"]),
        valor=_num(criterio["valor_obtenido"], dec),
        u=escape(unidad),
        comp=escape(criterio["comparacion"]),
        limite=_num(criterio["valor_limite"], dec),
    )


def _isolux(r: Dict[str, Any], vial: Dict[str, Any]) -> str:
    """Mapa de la malla: X a lo largo de la vialidad, Y a traves."""
    malla = r["malla"]
    xs: List[float] = malla["xs"]
    ys: List[float] = malla["ys"]
    e: List[List[float]] = malla["e"]
    lo, hi = malla["minimo_lx"], malla["maximo_lx"]
    rango = (hi - lo) or 1.0

    mx, my = 82, 30           # margenes para las etiquetas de los ejes
    cw, ch = 62, 46           # tamano de celda
    w = mx + cw * len(xs) + 16
    h = my + ch * len(ys) + 46

    p: List[str] = [
        '<svg viewBox="0 0 {} {}" class="isolux" role="img" '
        'aria-label="Mapa de iluminancia en lux sobre la calzada">'.format(w, h)
    ]
    for i, x in enumerate(xs):
        for j, y in enumerate(ys):
            v = e[i][j]
            t = (v - lo) / rango
            cx, cy = mx + i * cw, my + j * ch
            p.append(
                '<rect x="{}" y="{}" width="{}" height="{}" fill="{}"/>'.format(
                    cx, cy, cw, ch, _color(t))
            )
            # texto claro sobre lo oscuro y al reves, para que siempre se lea
            p.append(
                '<text x="{}" y="{}" class="celda" fill="{}">{}</text>'.format(
                    cx + cw / 2, cy + ch / 2 + 4,
                    "#12161d" if t > 0.62 else "#eef2f8", _num(v, 1))
            )
    # eje transversal: la posicion real de cada linea de calculo
    for j, y in enumerate(ys):
        p.append(
            '<text x="{}" y="{}" class="eje ejey">{} m</text>'.format(
                mx - 10, my + j * ch + ch / 2 + 4, _num(y, 2))
        )
    # eje longitudinal: solo los extremos, para no saturar
    for i in (0, len(xs) - 1):
        p.append(
            '<text x="{}" y="{}" class="eje">{} m</text>'.format(
                mx + i * cw + cw / 2, my - 10, _num(xs[i], 2))
        )
    p.append(
        '<text x="{}" y="{}" class="eje ejetit">a lo largo de la vialidad &rarr;</text>'.format(
            mx, my + ch * len(ys) + 24)
    )
    p.append(
        '<text x="12" y="{}" class="eje ejetit" transform="rotate(-90 12 {})">'
        "&larr; ancho de calzada</text>".format(
            my + ch * len(ys) / 2, my + ch * len(ys) / 2)
    )
    p.append("</svg>")
    return "".join(p)


def _leyenda(lo: float, hi: float) -> str:
    paradas = "".join(
        '<stop offset="{:.0%}" stop-color="{}"/>'.format(t, _color(t))
        for t, _ in RAMPA
    )
    return (
        '<div class="leyenda">'
        '<svg viewBox="0 0 220 10" preserveAspectRatio="none" aria-hidden="true">'
        '<defs><linearGradient id="g">{}</linearGradient></defs>'
        '<rect width="220" height="10" fill="url(#g)"/></svg>'
        '<div class="leyenda-t"><span>{} lx</span><span>{} lx</span></div>'
        "</div>"
    ).format(paradas, _num(lo, 1), _num(hi, 1))


def _fila(r: Dict[str, Any], recomendado: bool) -> str:
    m, n = r["malla"], r["nom"]
    return (
        '<tr class="{cls}">'
        '<td class="modelo"><span class="mnombre">{modelo}</span>'
        '<span class="mfab">{fab}</span>{marca}</td>'
        '<td class="n">{w}</td><td class="n">{eprom}</td><td class="n">{emin}</td>'
        '<td class="n">{unif}</td><td class="n">{dpea}</td>'
        '<td><span class="veredicto {vcls}">{vtxt}</span></td>'
        "</tr>"
    ).format(
        cls="recomendado" if recomendado else "",
        modelo=escape(r["catalogo"]),
        fab=escape(r["fabricante"] or ""),
        marca='<span class="rec">recomendado</span>' if recomendado else "",
        w=_num(r["watts"], 1),
        eprom=_num(m["promedio_lx"], 2),
        emin=_num(m["minimo_lx"], 2),
        unif=_num(m["uniformidad"], 2),
        dpea=_num(r["dpea_w_m2"], 3),
        vcls="ok" if n["cumple"] else "bad",
        vtxt="CUMPLE" if n["cumple"] else "NO CUMPLE",
    )


def _detalle(r: Dict[str, Any], vial: Dict[str, Any]) -> str:
    m, n = r["malla"], r["nom"]
    crit = n["criterios"]
    return (
        '<section class="detalle">'
        '<h3>{modelo} <span class="w">{w} W</span></h3>'
        '<div class="criterios">{c1}{c2}{c3}</div>'
        '<div class="mapa">{leyenda}{svg}</div>'
        "</section>"
    ).format(
        modelo=escape(r["catalogo"]),
        w=_num(r["watts"], 1),
        c1=_chip(crit["nivel_iluminacion"], " lx"),
        c2=_chip(crit["uniformidad"], ""),
        c3=_chip(crit["dpea"], " W/m²", 3),
        svg=_isolux(r, vial),
        leyenda=_leyenda(m["minimo_lx"], m["maximo_lx"]),
    )


CSS = """
:root{
  --ground:#f7f8fa; --surface:#ffffff; --surface-2:#eef1f6;
  --ink:#14181f; --ink-2:#4a5361; --ink-3:#79828f;
  --line:#dde1e8; --slate:#1f3a5f; --amber:#d08a1e;
  --ok:#1f7a4d; --ok-bg:#e7f3ec; --bad:#b3261e; --bad-bg:#fbeae8;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --ground:#0f1319; --surface:#171c24; --surface-2:#1f2630;
    --ink:#e8ebf0; --ink-2:#99a3b2; --ink-3:#6d7889;
    --line:#2a323e; --slate:#7fa8dc; --amber:#e8a93f;
    --ok:#4fbf87; --ok-bg:#14301f; --bad:#e2685c; --bad-bg:#361a18;
  }
}
:root[data-theme="dark"]{
  --ground:#0f1319; --surface:#171c24; --surface-2:#1f2630;
  --ink:#e8ebf0; --ink-2:#99a3b2; --ink-3:#6d7889;
  --line:#2a323e; --slate:#7fa8dc; --amber:#e8a93f;
  --ok:#4fbf87; --ok-bg:#14301f; --bad:#e2685c; --bad-bg:#361a18;
}
*{box-sizing:border-box}
body{
  margin:0; background:var(--ground); color:var(--ink);
  font-family:"Source Sans 3",system-ui,-apple-system,"Segoe UI",sans-serif;
  font-size:16px; line-height:1.6;
}
.wrap{max-width:1000px; margin:0 auto; padding:40px 24px 72px; display:flex; flex-direction:column; gap:36px}
h1,h2,h3{font-family:Archivo,system-ui,sans-serif; text-wrap:balance; margin:0; line-height:1.2}
h1{font-size:clamp(28px,4vw,40px); font-weight:700; letter-spacing:-.02em}
h2{font-size:22px; font-weight:650; letter-spacing:-.01em}
h3{font-size:18px; font-weight:650}
p{margin:0}
.eyebrow{
  font-family:Archivo,sans-serif; font-size:12px; font-weight:600;
  letter-spacing:.14em; text-transform:uppercase; color:var(--slate);
}
header .sub{color:var(--ink-2); max-width:62ch; margin-top:10px}

/* Bloque de identificacion: rejilla de dato/valor, como un dictamen */
.ident{
  background:var(--surface); border:1px solid var(--line); border-radius:6px;
  padding:22px 24px; display:grid; gap:18px 32px;
  grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
}
.ident div{display:flex; flex-direction:column; gap:3px}
.ident dt,.ident .k{
  font-size:11px; font-weight:600; letter-spacing:.08em; text-transform:uppercase;
  color:var(--ink-3); font-family:Archivo,sans-serif;
}
.ident .v{font-family:"IBM Plex Mono",ui-monospace,monospace; font-size:15px; font-variant-numeric:tabular-nums}

/* Tabla comparativa: la pieza central */
.tablabox{overflow-x:auto; border:1px solid var(--line); border-radius:6px; background:var(--surface)}
table{border-collapse:collapse; width:100%; min-width:720px}
th,td{padding:13px 14px; text-align:left; border-bottom:1px solid var(--line)}
thead th{
  font-family:Archivo,sans-serif; font-size:11px; font-weight:600;
  letter-spacing:.07em; text-transform:uppercase; color:var(--ink-3);
  background:var(--surface-2); white-space:nowrap;
}
tbody tr:last-child td{border-bottom:none}
td.n{font-family:"IBM Plex Mono",monospace; font-variant-numeric:tabular-nums; text-align:right; white-space:nowrap}
th.n{text-align:right}
.modelo{min-width:190px}
.mnombre{display:block; font-weight:600}
.mfab{display:block; font-size:12px; color:var(--ink-3)}
tr.recomendado{background:color-mix(in srgb,var(--amber) 9%,transparent)}
tr.recomendado td:first-child{box-shadow:inset 3px 0 0 var(--amber)}
.rec{
  display:inline-block; margin-top:5px; font-family:Archivo,sans-serif;
  font-size:10px; font-weight:600; letter-spacing:.08em; text-transform:uppercase;
  color:var(--amber); border:1px solid var(--amber); border-radius:3px; padding:1px 6px;
}
.veredicto{
  font-family:Archivo,sans-serif; font-size:11px; font-weight:650;
  letter-spacing:.07em; padding:4px 9px; border-radius:3px; white-space:nowrap;
}
.veredicto.ok{color:var(--ok); background:var(--ok-bg)}
.veredicto.bad{color:var(--bad); background:var(--bad-bg)}

/* Detalle por luminario */
.detalle{display:flex; flex-direction:column; gap:16px; padding-top:8px; border-top:1px solid var(--line)}
.detalle h3 .w{font-family:"IBM Plex Mono",monospace; font-size:14px; color:var(--ink-3); font-weight:400}
.criterios{display:grid; gap:12px; grid-template-columns:repeat(auto-fit,minmax(190px,1fr))}
.criterio{
  border:1px solid var(--line); border-left-width:3px; border-radius:5px;
  padding:11px 14px; display:flex; flex-direction:column; gap:2px; background:var(--surface);
}
.criterio.ok{border-left-color:var(--ok)}
.criterio.bad{border-left-color:var(--bad)}
.criterio-nombre{font-size:12px; color:var(--ink-2)}
.criterio-valor{font-family:"IBM Plex Mono",monospace; font-size:22px; font-variant-numeric:tabular-nums}
.criterio.ok .criterio-valor{color:var(--ok)}
.criterio.bad .criterio-valor{color:var(--bad)}
.criterio-valor .u{font-size:13px; color:var(--ink-3)}
.criterio-limite{font-size:11px; color:var(--ink-3); font-family:"IBM Plex Mono",monospace}

.mapa{display:flex; flex-direction:column; gap:8px; overflow-x:auto}
.isolux{min-width:620px; width:100%; height:auto; display:block}
.isolux .celda{font-family:"IBM Plex Mono",monospace; font-size:12px; text-anchor:middle}
.isolux .eje{font-family:"IBM Plex Mono",monospace; font-size:11px; fill:var(--ink-3); text-anchor:middle}
.isolux .ejey{text-anchor:end}
.isolux .ejetit{font-family:Archivo,sans-serif; font-size:11px; text-anchor:start; letter-spacing:.04em}
.leyenda{max-width:220px}
.leyenda svg{width:220px; height:10px; border-radius:2px; display:block}
.leyenda-t{display:flex; justify-content:space-between; font-family:"IBM Plex Mono",monospace; font-size:11px; color:var(--ink-3); margin-top:3px}
.aviso{
  font-size:13.5px; color:var(--ink-2); background:var(--surface-2);
  border-left:3px solid var(--amber); border-radius:0 5px 5px 0; padding:11px 14px; max-width:78ch;
}

.notas{display:flex; flex-direction:column; gap:14px; font-size:14px; color:var(--ink-2); max-width:78ch}
.notas h2{color:var(--ink); margin-bottom:2px}
.notas strong{color:var(--ink)}
footer{font-size:12.5px; color:var(--ink-3); border-top:1px solid var(--line); padding-top:18px}
@media (prefers-reduced-motion:reduce){*{animation:none!important; transition:none!important}}
"""


def html(datos: Dict[str, Any]) -> str:
    v = datos["vialidad"]
    n = datos["nom"]
    p = datos["perdidas"]
    res: List[Dict[str, Any]] = datos["resultados"]
    rec = datos.get("recomendacion")

    etiqueta_vial = nom.ETIQUETAS_VIALIDAD.get(
        n["clasificacion_vialidad"], n["clasificacion_vialidad"])
    disp = {
        "Single-side": "unilateral", "Staggered": "tresbolillo",
        "Median mounted": "central doble", "Opposite": "bilateral opuesta",
    }.get(v["disposicion"], v["disposicion"])
    cumplen = sum(1 for r in res if r["nom"]["cumple"])

    ident = [
        ("Clasificación", etiqueta_vial),
        ("Pavimento", n["pavimento"]),
        ("Carriles", "{} &times; {} m".format(v["num_carriles"], _num(v["ancho_carril"], 2))),
        ("Camellón", "{} m".format(_num(v["camellon"], 2))),
        ("Disposición", disp),
        ("Altura de montaje", "{} m".format(_num(v["altura_montaje"], 2))),
        ("Interpostal", "{} m".format(_num(v["interpostal"], 2))),
        ("Retranqueo", "{} m".format(_num(v["retranqueo"], 2))),
        ("Largo de brazo", "{} m".format(_num(v["largo_brazo"], 2))),
        ("Ancho para DPEA", "{} m".format(_num(v["ancho_sin_camellon"], 2))),
        ("LLD &times; LDD &times; BF", "{} &times; {} &times; {} = {}".format(
            _num(p["lld"], 2), _num(p["ldd"], 2), _num(p["bf"], 2), _num(p["llf_total"], 3))),
        ("Criterio de azimut", "IESNA RP-8"),
    ]
    ident_html = "".join(
        '<div><span class="k">{}</span><span class="v">{}</span></div>'.format(k, val)
        for k, val in ident
    )

    return """<meta charset="utf-8">
<title>Estudio de alumbrado {vial}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@400;600;700&family=IBM+Plex+Mono:wght@400;500&family=Source+Sans+3:wght@400;600&display=swap">
<style>{css}</style>
<div class="wrap">
  <header>
    <p class="eyebrow">Memoria de cálculo &middot; NOM-013-ENER-2013</p>
    <h1>{vial}</h1>
    <p class="sub">Evaluación de {total} luminario{s} sobre la geometría descrita abajo.
    {cumplen} de {total} cumplen los tres criterios de la norma.{recfrase}</p>
  </header>

  <section class="ident">{ident}</section>

  <section>
    <h2>Comparativa</h2>
    <div class="tablabox"><table>
      <thead><tr>
        <th>Luminario</th><th class="n">W</th><th class="n">E<sub>prom</sub> lx</th>
        <th class="n">E<sub>min</sub> lx</th><th class="n">E<sub>prom</sub>/E<sub>min</sub></th>
        <th class="n">DPEA W/m&sup2;</th><th>Veredicto</th>
      </tr></thead>
      <tbody>{filas}</tbody>
    </table></div>
  </section>

  <section>
    <h2>Detalle por luminario</h2>
    {detalles}
  </section>

  <section class="notas">
    <h2>Método y alcances</h2>
    <p>La iluminancia en cada punto es
    <strong>E = I(&gamma;,C) &middot; cos<sup>3</sup>&gamma; &middot; LLF / H<sup>2</sup></strong>,
    sumada sobre todos los luminarios del tramo, con I interpolada bilinealmente
    de la matriz fotométrica del archivo IES. La malla sigue el método IES:
    paso de interpostal/10 topado a 5 m a lo largo de la vialidad, dos puntos
    por carril a un cuarto y tres cuartos de su ancho, sobre un tramo
    interpostal completo. E<sub>prom</sub> es la media aritmética de esos puntos
    y la uniformidad es E<sub>prom</sub>/E<sub>min</sub> (NOM-013 8.3).</p>
    <p><strong>Esto es un cálculo de diseño, no un dictamen.</strong> La
    verificación en campo que hace una Unidad de Verificación usa la fórmula de
    nueve puntos del Apéndice C, que es una malla distinta y más gruesa. Este
    estudio sirve para elegir la sustitución; no sustituye la verificación.</p>
    <p>El DPEA se calcula como la carga conectada por tramo interpostal entre el
    área de ese tramo, con el ancho de calzada <strong>sin aceras ni
    camellones</strong>, como pide la sección 6.1 de la norma.</p>
  </section>

  <footer>Generado por el motor de cálculo de ilumilm. Los watts de cada
  luminario provienen del archivo IES o fueron capturados a mano; conviene
  confirmarlos contra la ficha del fabricante, porque de ellos depende
  directamente el DPEA.</footer>
</div>
""".format(
        css=CSS,
        vial=escape(etiqueta_vial),
        total=len(res),
        s="" if len(res) == 1 else "s",
        cumplen=cumplen,
        recfrase=" Se recomienda el <strong>{}</strong> por ser el de menor potencia entre los que cumplen.".format(
            escape(rec)) if rec else "",
        ident=ident_html,
        filas="".join(_fila(r, r["catalogo"] == rec) for r in res),
        detalles="".join(_detalle(r, v) for r in res),
    )


def main(argv: Sequence[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else list(argv)
    if len(argv) != 1:
        print("Uso: python -m engine.report estudios/<nombre>/resultados.json",
              file=sys.stderr)
        return 2
    ruta = Path(argv[0])
    if not ruta.exists():
        print("No existe el archivo: {}".format(ruta), file=sys.stderr)
        return 1
    datos = json.loads(ruta.read_text(encoding="utf-8"))
    destino = ruta.parent / "reporte.html"
    destino.write_text(html(datos), encoding="utf-8")
    print("Reporte escrito en {}".format(destino))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
