"""Genera la memoria de calculo en HTML a partir de `resultados.json`.

Uso:
    python -m engine.report estudios/<nombre>/resultados.json

Escribe `reporte.html` junto al JSON. El archivo es autocontenido salvo por las
tipografias de Google Fonts, y esta pensado para publicarse como Artifact.

El reporte deja mover tres parametros sin volver a correr nada:

* Clasificacion de vialidad. No cambia la fisica, solo los umbrales de la
  norma, asi que se resuelve en el navegador con los umbrales ya embebidos.
* Interpostal y altura de montaje. Estas SI cambian la geometria, los angulos
  y el area del tramo, o sea el calculo completo. Se resuelven con el barrido
  que precalcula `engine.cli`: una malla por cada combinacion.

De ahi que la pagina no recalcule fotometria en JavaScript. Reimplementar el
motor ahi seria tener dos motores que mantener de acuerdo.
"""
from __future__ import annotations

import json
import sys
from html import escape
from pathlib import Path
from typing import Any, Dict, List, Sequence

from . import nom

TITULO = "Estudio de alumbrado público"

# Rampa de la escala isolux: de la noche al vapor de sodio. Los luminarios de
# vialidad viven en ese gradiente, asi que la escala dice algo del tema en vez
# de ser un arcoiris generico.
RAMPA = [
    [0.00, [0x16, 0x23, 0x3A]],
    [0.25, [0x2E, 0x4B, 0x7A]],
    [0.50, [0x7A, 0x63, 0x62]],
    [0.75, [0xB5, 0x76, 0x2A]],
    [0.90, [0xE8, 0xB4, 0x5C]],
    [1.00, [0xFB, 0xEB, 0xC8]],
]

# Geometria del mapa. La malla siempre es de 10 puntos a lo largo (la ventana
# evaluada del metodo IES) por dos por carril, asi que el esqueleto del SVG se
# emite una vez y el script solo cambia colores y numeros.
MX, MY, CW, CH = 82, 30, 62, 46


def _num(x: float, dec: int = 2) -> str:
    if x == float("inf"):
        return "∞"
    return "{:,.{}f}".format(x, dec).replace(",", " ")


def _isolux(idx: int, n_x: int, n_y: int) -> str:
    """Esqueleto del mapa. Los colores y los valores los pone el script."""
    w = MX + CW * n_x + 16
    h = MY + CH * n_y + 46
    p: List[str] = [
        '<svg viewBox="0 0 {} {}" class="isolux" data-i="{}" role="img" '
        'aria-label="Mapa de iluminancia en lux sobre la calzada">'.format(w, h, idx)
    ]
    for i in range(n_x):
        for j in range(n_y):
            cx, cy = MX + i * CW, MY + j * CH
            p.append('<rect x="{}" y="{}" width="{}" height="{}"/>'.format(cx, cy, CW, CH))
            p.append('<text x="{}" y="{}" class="celda"></text>'.format(
                cx + CW / 2, cy + CH / 2 + 4))
    for j in range(n_y):
        p.append('<text x="{}" y="{}" class="eje ejey"></text>'.format(
            MX - 10, MY + j * CH + CH / 2 + 4))
    for i in (0, n_x - 1):
        p.append('<text x="{}" y="{}" class="eje ejex"></text>'.format(
            MX + i * CW + CW / 2, MY - 10))
    p.append('<text x="{}" y="{}" class="eje ejetit">a lo largo de la vialidad →</text>'.format(
        MX, MY + CH * n_y + 24))
    p.append('<text x="12" y="{}" class="eje ejetit" transform="rotate(-90 12 {})">'
             "← ancho de calzada</text>".format(MY + CH * n_y / 2, MY + CH * n_y / 2))
    p.append("</svg>")
    return "".join(p)


def _leyenda() -> str:
    paradas = "".join(
        '<stop offset="{:.0%}" stop-color="rgb({},{},{})"/>'.format(t, *c)
        for t, c in RAMPA
    )
    return (
        '<div class="leyenda">'
        '<svg viewBox="0 0 220 10" preserveAspectRatio="none" aria-hidden="true">'
        '<defs><linearGradient id="g">{}</linearGradient></defs>'
        '<rect width="220" height="10" fill="url(#g)"/></svg>'
        '<div class="leyenda-t"><span class="lo"></span><span class="hi"></span></div>'
        "</div>"
    ).format(paradas)


def _fila(i: int, r: Dict[str, Any]) -> str:
    """Fila de la tabla. Todos los numeros los rellena el script."""
    return (
        '<tr data-i="{i}">'
        '<td class="modelo"><span class="mnombre">{modelo}</span>'
        '<span class="mfab">{fab}</span><span class="rec" hidden>recomendado</span></td>'
        '<td class="n" data-label="Potencia">{w}</td>'
        '<td class="n c-eprom" data-label="Eprom"></td>'
        '<td class="n c-emin" data-label="Emin"></td>'
        '<td class="n c-unif" data-label="Uniformidad"></td>'
        '<td class="n c-dpea" data-label="DPEA"></td>'
        '<td class="vcell" data-label="Veredicto"></td>'
        "</tr>"
    ).format(i=i, modelo=escape(r["catalogo"]),
             fab=escape(r["fabricante"] or ""), w=_num(r["watts"], 1))


def _detalle(i: int, r: Dict[str, Any], n_x: int, n_y: int) -> str:
    return (
        '<section class="detalle" data-i="{i}">'
        '<h3>{modelo} <span class="w">{w} W</span></h3>'
        '<div class="criterios"></div>'
        '<div class="mapa">{leyenda}{svg}</div>'
        "</section>"
    ).format(i=i, modelo=escape(r["catalogo"]), w=_num(r["watts"], 1),
             svg=_isolux(i, n_x, n_y), leyenda=_leyenda())


CSS = """
:root{
  --ground:#f7f8fa; --surface:#ffffff; --surface-2:#eef1f6;
  --ink:#14181f; --ink-2:#4a5361; --ink-3:#565e6c;
  --line:#dde1e8; --slate:#1f3a5f; --amber:#d08a1e; --amber-txt:#875608;
  --ok:#1f7a4d; --ok-bg:#e7f3ec; --bad:#b3261e; --bad-bg:#fbeae8;
  --medida:62ch;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --ground:#0f1319; --surface:#171c24; --surface-2:#1f2630;
    --ink:#e8ebf0; --ink-2:#99a3b2; --ink-3:#98a2b2;
    --line:#2a323e; --slate:#7fa8dc; --amber:#e8a93f; --amber-txt:#e8a93f;
    --ok:#4fbf87; --ok-bg:#14301f; --bad:#e2685c; --bad-bg:#361a18;
  }
}
:root[data-theme="dark"]{
  --ground:#0f1319; --surface:#171c24; --surface-2:#1f2630;
  --ink:#e8ebf0; --ink-2:#99a3b2; --ink-3:#98a2b2;
  --line:#2a323e; --slate:#7fa8dc; --amber:#e8a93f; --amber-txt:#e8a93f;
  --ok:#4fbf87; --ok-bg:#14301f; --bad:#e2685c; --bad-bg:#361a18;
}
*{box-sizing:border-box}
[hidden]{display:none!important}
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
header .sub,.notas{max-width:var(--medida)}
header .sub{color:var(--ink-2); margin-top:10px}
.barra{display:flex; align-items:flex-start; justify-content:space-between; gap:20px}
.tema{
  flex:none; display:inline-flex; align-items:center; gap:7px; cursor:pointer;
  background:var(--surface); color:var(--ink-2); border:1px solid var(--line);
  border-radius:999px; padding:7px 14px; font-family:Archivo,sans-serif;
  font-size:12px; font-weight:600; letter-spacing:.04em;
}
.tema:hover{color:var(--ink); border-color:var(--ink-3)}
.tema:focus-visible{outline:2px solid var(--slate); outline-offset:2px}
.tema svg{width:14px; height:14px; fill:currentColor}
.tema .i-sol,.tema .i-luna{display:none}
/* Se muestra el icono de la accion, no el del estado: en oscuro, un sol. */
:root[data-theme="dark"] .tema .i-sol{display:block}
:root[data-theme="light"] .tema .i-luna{display:block}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .tema .i-sol{display:block}}
@media (prefers-color-scheme: light){:root:not([data-theme="dark"]) .tema .i-luna{display:block}}

/* Controles: lo que se puede mover, separado de lo que es dato fijo */
.panel{
  background:var(--surface); border:1px solid var(--line); border-radius:6px;
  padding:20px 24px; display:flex; flex-direction:column; gap:18px;
}
.panel-tit{display:flex; align-items:baseline; justify-content:space-between; gap:16px; flex-wrap:wrap}
.panel-tit h2{
  font-size:12px; letter-spacing:.1em; text-transform:uppercase;
  color:var(--ink-3); font-weight:600;
}
.reset{
  background:none; border:none; cursor:pointer; padding:0; color:var(--slate);
  font-family:Archivo,sans-serif; font-size:12px; font-weight:600; text-decoration:underline;
}
.reset:focus-visible{outline:2px solid var(--slate); outline-offset:2px}
.controles{display:grid; gap:20px 32px; grid-template-columns:repeat(auto-fit,minmax(230px,1fr))}
.ctrl{display:flex; flex-direction:column; gap:6px; min-width:0}
.ctrl-cab{display:flex; align-items:baseline; justify-content:space-between; gap:12px; min-height:20px}
.ctrl label{
  font-size:11px; font-weight:600; letter-spacing:.08em; text-transform:uppercase;
  color:var(--ink-3); font-family:Archivo,sans-serif;
}
.ctrl output{
  font-family:"IBM Plex Mono",monospace; font-size:15px; font-variant-numeric:tabular-nums;
  color:var(--ink);
}
.ctrl output.movido{color:var(--amber-txt)}
.ctrl select{
  font-family:"IBM Plex Mono",ui-monospace,monospace; font-size:15px;
  background:var(--surface-2); color:var(--ink); border:1px solid var(--line);
  border-radius:4px; padding:6px 8px; width:100%; cursor:pointer;
}
.ctrl select:focus-visible,.ctrl input:focus-visible{outline:2px solid var(--slate); outline-offset:1px}
.ctrl input[type=range]{width:100%; accent-color:var(--slate); cursor:pointer; margin:6px 0 0}
.marcas{display:flex; justify-content:space-between; font-family:"IBM Plex Mono",monospace; font-size:10.5px; color:var(--ink-3)}
.avisomov{
  font-size:13px; color:var(--ink-2); background:var(--surface-2);
  border-left:3px solid var(--amber); border-radius:0 4px 4px 0; padding:9px 13px;
}

.ident{
  background:var(--surface); border:1px solid var(--line); border-radius:6px;
  padding:22px 24px; display:grid; gap:18px 32px;
  grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
}
.ident div{display:flex; flex-direction:column; gap:3px; min-width:0}
.ident .k{
  font-size:11px; font-weight:600; letter-spacing:.08em; text-transform:uppercase;
  color:var(--ink-3); font-family:Archivo,sans-serif;
}
.ident .v{font-family:"IBM Plex Mono",ui-monospace,monospace; font-size:15px; font-variant-numeric:tabular-nums}

.tablabox{overflow-x:auto; border:1px solid var(--line); border-radius:6px; background:var(--surface)}
table{border-collapse:collapse; width:100%; min-width:720px; color:var(--ink)}
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
  color:var(--amber-txt); border:1px solid var(--amber-txt); border-radius:3px; padding:1px 6px;
}
.veredicto{
  font-family:Archivo,sans-serif; font-size:11px; font-weight:650;
  letter-spacing:.07em; padding:4px 9px; border-radius:3px; white-space:nowrap;
  display:inline-block;
}
.veredicto.ok{color:var(--ok); background:var(--ok-bg)}
.veredicto.bad{color:var(--bad); background:var(--bad-bg)}
.falla{display:block; margin-top:5px; font-size:11.5px; color:var(--bad); line-height:1.45}

@media (max-width:760px){
  .tablabox{overflow-x:visible; border:none; background:none; border-radius:0}
  table{min-width:0; display:block}
  thead{display:none}
  tbody,tr,td{display:block}
  tbody tr{
    border:1px solid var(--line); border-radius:6px; background:var(--surface);
    margin-bottom:12px; padding:4px 0;
  }
  tbody tr:last-child{margin-bottom:0}
  td{border-bottom:none; padding:7px 14px; display:flex; justify-content:space-between;
     align-items:baseline; gap:16px; text-align:right}
  td::before{
    content:attr(data-label); flex:none; text-align:left;
    font-family:Archivo,sans-serif; font-size:11px; font-weight:600;
    letter-spacing:.07em; text-transform:uppercase; color:var(--ink-3);
  }
  td.modelo{display:block; text-align:left; padding-top:12px}
  td.modelo::before{display:none}
  tr.recomendado td:first-child{box-shadow:none}
  tr.recomendado{box-shadow:inset 3px 0 0 var(--amber)}
  .falla{text-align:right}
}

.detalle{display:flex; flex-direction:column; gap:16px; padding-top:8px; border-top:1px solid var(--line)}
.detalle h3{display:flex; align-items:baseline; gap:10px; flex-wrap:wrap}
.detalle h3 .w{
  font-family:"IBM Plex Mono",monospace; font-size:14px; font-weight:500;
  color:var(--ink-2); background:var(--surface-2); border:1px solid var(--line);
  border-radius:3px; padding:2px 8px; font-variant-numeric:tabular-nums;
}
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

.notas{display:flex; flex-direction:column; gap:14px; font-size:14px; color:var(--ink-2)}
.notas h2{color:var(--ink); margin-bottom:2px}
.notas strong{color:var(--ink)}
footer{font-size:12.5px; color:var(--ink-3); border-top:1px solid var(--line); padding-top:18px; max-width:var(--medida)}
@media (prefers-reduced-motion:reduce){*{animation:none!important; transition:none!important}}
"""

JS = r"""
(function(){
  var raiz = document.documentElement;

  // --- tema -------------------------------------------------------------
  var btn = document.getElementById('tema');
  try{
    var g = localStorage.getItem('tema');
    if(g === 'dark' || g === 'light') raiz.setAttribute('data-theme', g);
  }catch(e){}
  function oscuro(){
    var t = raiz.getAttribute('data-theme');
    return t ? t === 'dark' : window.matchMedia('(prefers-color-scheme: dark)').matches;
  }
  function rotulo(){
    btn.setAttribute('aria-pressed', oscuro() ? 'true' : 'false');
    btn.querySelector('.txt').textContent = oscuro() ? 'Modo claro' : 'Modo oscuro';
  }
  btn.addEventListener('click', function(){
    var n = oscuro() ? 'light' : 'dark';
    raiz.setAttribute('data-theme', n);
    try{ localStorage.setItem('tema', n); }catch(e){}
    rotulo();
  });
  rotulo();

  // --- utilidades -------------------------------------------------------
  function color(t){
    t = Math.max(0, Math.min(1, t));
    for(var k = 0; k < RAMPA.length - 1; k++){
      var a = RAMPA[k], b = RAMPA[k + 1];
      if(t <= b[0]){
        var f = b[0] > a[0] ? (t - a[0]) / (b[0] - a[0]) : 0;
        return 'rgb(' + a[1].map(function(c, m){
          return Math.round(c + (b[1][m] - c) * f);
        }).join(',') + ')';
      }
    }
    return 'rgb(251,235,200)';
  }
  function fmt(x, d){
    if(!isFinite(x)) return '∞';
    return x.toFixed(d).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  }

  // --- controles --------------------------------------------------------
  var selC = document.getElementById('clasif');
  var rgS  = document.getElementById('interpostal');
  var rgH  = document.getElementById('altura');
  var outS = document.getElementById('vInterpostal');
  var outH = document.getElementById('vAltura');
  var aviso = document.getElementById('avisomov');
  var reset = document.getElementById('reset');

  var filas = [].slice.call(document.querySelectorAll('tbody tr'));
  var detalles = [].slice.call(document.querySelectorAll('.detalle'));
  var svgs = [].slice.call(document.querySelectorAll('svg.isolux'));
  var NY = B.ys.length;

  function chip(nombre, valor, unidad, comp, limite, ok, dec){
    return '<div class="criterio ' + (ok ? 'ok' : 'bad') + '">' +
      '<span class="criterio-nombre">' + nombre + '</span>' +
      '<span class="criterio-valor">' + fmt(valor, dec) +
      '<span class="u">' + unidad + '</span></span>' +
      '<span class="criterio-limite">' + comp + ' ' + fmt(limite, dec) + '</span></div>';
  }

  function pintaMapa(idx, malla, lo, hi){
    var svg = svgs[idx];
    var rango = (hi - lo) || 1;
    var rects = svg.querySelectorAll('rect');
    var celdas = svg.querySelectorAll('text.celda');
    for(var n = 0; n < malla.length; n++){
      var t = (malla[n] - lo) / rango;
      rects[n].setAttribute('fill', color(t));
      celdas[n].textContent = malla[n].toFixed(1);
      celdas[n].setAttribute('fill', t > 0.62 ? '#12161d' : '#eef2f8');
    }
    var ejesY = svg.querySelectorAll('text.ejey');
    for(var j = 0; j < NY; j++) ejesY[j].textContent = B.ys[j].toFixed(2) + ' m';
    var xs = B.xs_por_interpostal[iS];
    var ejesX = svg.querySelectorAll('text.ejex');
    ejesX[0].textContent = xs[0].toFixed(2) + ' m';
    ejesX[1].textContent = xs[xs.length - 1].toFixed(2) + ' m';
    var ley = detalles[idx].querySelector('.leyenda-t');
    ley.querySelector('.lo').textContent = fmt(lo, 1) + ' lx';
    ley.querySelector('.hi').textContent = fmt(hi, 1) + ' lx';
  }

  var iS = B.i_interpostal_estudio, iH = B.i_altura_estudio;

  function pinta(){
    var S = B.interpostales[iS], H = B.alturas[iH];
    var req = REQ[selC.value];
    var area = B.ancho_sin_camellon * S;
    var mejor = -1, mejorW = Infinity, cumplen = 0;

    outS.textContent = fmt(S, 2) + ' m';
    outH.textContent = fmt(H, 2) + ' m';
    outS.classList.toggle('movido', iS !== B.i_interpostal_estudio);
    outH.classList.toggle('movido', iH !== B.i_altura_estudio);
    var movido = (iS !== B.i_interpostal_estudio) || (iH !== B.i_altura_estudio);
    aviso.hidden = !movido;
    reset.hidden = !movido;

    LUM.forEach(function(l, i){
      var st = B.stats[iS][iH][i];           // [eprom, emin, emax, uniformidad]
      var dpea = B.watts_por_tramo[i] / area;
      var cIlum = st[0] >= req.eprom_min;
      var cUnif = st[3] <= req.unif_max;
      var cDpea = dpea <= req.dpea_max;
      var todo = cIlum && cUnif && cDpea;
      if(todo){ cumplen++; if(l.w < mejorW){ mejorW = l.w; mejor = i; } }

      var f = filas[i];
      f.querySelector('.c-eprom').textContent = fmt(st[0], 2);
      f.querySelector('.c-emin').textContent  = fmt(st[1], 2);
      f.querySelector('.c-unif').textContent  = fmt(st[3], 2);
      f.querySelector('.c-dpea').textContent  = fmt(dpea, 3);

      var fallos = [];
      if(!cIlum) fallos.push('nivel de iluminación');
      if(!cUnif) fallos.push('uniformidad');
      if(!cDpea) fallos.push('DPEA');
      f.querySelector('.vcell').innerHTML =
        '<span class="veredicto ' + (todo ? 'ok' : 'bad') + '">' +
        (todo ? 'CUMPLE' : 'NO CUMPLE') + '</span>' +
        (fallos.length ? '<span class="falla">falla ' + fallos.join(', ') + '</span>' : '');

      detalles[i].querySelector('.criterios').innerHTML =
        chip('Nivel de iluminación (Eprom)', st[0], ' lx', '&ge; mínimo', req.eprom_min, cIlum, 2) +
        chip('Uniformidad (Eprom/Emin)', st[3], '', '&le; máximo', req.unif_max, cUnif, 2) +
        chip('DPEA', dpea, ' W/m²', '&le; máximo', req.dpea_max, cDpea, 3);

      pintaMapa(i, B.mallas[iS][iH][i], st[1], st[2]);
    });

    filas.forEach(function(f, i){
      f.classList.toggle('recomendado', i === mejor);
      f.querySelector('.rec').hidden = (i !== mejor);
    });

    document.getElementById('resumen').innerHTML =
      'Evaluación de ' + LUM.length + ' luminario' + (LUM.length === 1 ? '' : 's') +
      ' con interpostal de ' + fmt(S, 2) + ' m y altura de montaje de ' + fmt(H, 2) +
      ' m. ' + cumplen + ' de ' + LUM.length + ' cumplen los tres criterios de la norma.' +
      (mejor >= 0 ? ' Se recomienda el <strong>' + LUM[mejor].cat +
        '</strong> por ser el de menor potencia entre los que cumplen.' : '');
  }

  rgS.addEventListener('input', function(){ iS = +rgS.value; pinta(); });
  rgH.addEventListener('input', function(){ iH = +rgH.value; pinta(); });
  selC.addEventListener('change', pinta);
  reset.addEventListener('click', function(){
    iS = B.i_interpostal_estudio; iH = B.i_altura_estudio;
    rgS.value = iS; rgH.value = iH; pinta();
  });
  pinta();
})();
"""


def html(datos: Dict[str, Any]) -> str:
    v = datos["vialidad"]
    n = datos["nom"]
    p = datos["perdidas"]
    res: List[Dict[str, Any]] = datos["resultados"]
    b = datos["barrido"]

    pavimento = n["pavimento"]
    ancho = v["ancho_sin_camellon"]
    tabla = nom.TABLAS_POR_PAVIMENTO[pavimento]

    # Umbrales de las 7 clasificaciones, ya resueltos para este ancho de
    # calzada. El ancho no cambia con ningun control, asi que basta un numero.
    req = {
        clave: {
            "eprom_min": r.iluminancia_minima_promedio_lx,
            "unif_max": r.uniformidad_maxima,
            "dpea_max": nom.dpea_maximo(r, ancho),
        }
        for clave, r in tabla.items()
    }
    lum = [{"cat": r["catalogo"], "w": r["watts"]} for r in res]

    sel_actual = n["clasificacion_vialidad"]
    opciones = "".join(
        '<option value="{}"{}>{}</option>'.format(
            escape(clave), " selected" if clave == sel_actual else "",
            escape(nom.ETIQUETAS_VIALIDAD.get(clave, clave)))
        for clave in tabla
    )

    disp = {
        "Single-side": "unilateral", "Staggered": "tresbolillo",
        "Median mounted": "central doble", "Opposite": "bilateral opuesta",
    }.get(v["disposicion"], v["disposicion"])

    ident = [
        ("Pavimento", escape(pavimento)),
        ("Carriles", "{} × {} m".format(v["num_carriles"], _num(v["ancho_carril"], 2))),
        ("Camellón", "{} m".format(_num(v["camellon"], 2))),
        ("Disposición", disp),
        ("Retranqueo", "{} m".format(_num(v["retranqueo"], 2))),
        ("Largo de brazo", "{} m".format(_num(v["largo_brazo"], 2))),
        ("Ancho para DPEA", "{} m".format(_num(ancho, 2))),
        ("LLD × LDD × BF", "{} × {} × {} = {}".format(
            _num(p["lld"], 2), _num(p["ldd"], 2), _num(p["bf"], 2), _num(p["llf_total"], 3))),
        ("Criterio de azimut", "IESNA RP-8"),
    ]
    ident_html = "".join(
        '<div><span class="k">{}</span><span class="v">{}</span></div>'.format(k, val)
        for k, val in ident
    )

    ints, alts = b["interpostales"], b["alturas"]
    n_x = len(b["xs_por_interpostal"][0])
    n_y = len(b["ys"])

    return """<meta charset="utf-8">
<title>{titulo}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@400;600;700&family=IBM+Plex+Mono:wght@400;500&family=Source+Sans+3:wght@400;600&display=swap">
<style>{css}</style>
<div class="wrap">
  <header>
    <div class="barra">
      <div>
        <p class="eyebrow">Memoria de cálculo &middot; NOM-013-ENER-2013</p>
        <h1>{titulo}</h1>
      </div>
      <button class="tema" id="tema" type="button" aria-pressed="false">
        <svg class="i-sol" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 17a5 5 0 100-10 5 5 0 000 10zm0 2.5a1 1 0 011 1V22a1 1 0 11-2 0v-1.5a1 1 0 011-1zm0-19a1 1 0 011 1V3a1 1 0 11-2 0V1.5a1 1 0 011-1zM21 11h1.5a1 1 0 110 2H21a1 1 0 110-2zM1.5 11H3a1 1 0 110 2H1.5a1 1 0 110-2zm16.9 6l1 1a1 1 0 01-1.4 1.4l-1-1a1 1 0 011.4-1.4zM5 3.6l1 1A1 1 0 014.6 6l-1-1A1 1 0 015 3.6zm14.4 0A1 1 0 0120.4 5l-1 1A1 1 0 0118 4.6zM6 17a1 1 0 011.4 1.4l-1 1A1 1 0 015 18z"/></svg>
        <svg class="i-luna" viewBox="0 0 24 24" aria-hidden="true"><path d="M21 13.2A9 9 0 1110.8 3a7 7 0 1010.2 10.2z"/></svg>
        <span class="txt">Modo claro</span>
      </button>
    </div>
    <p class="sub" id="resumen"></p>
  </header>

  <section class="panel">
    <div class="panel-tit">
      <h2>Parámetros que puedes mover</h2>
      <button class="reset" id="reset" type="button" hidden>Volver a los valores del estudio</button>
    </div>
    <div class="controles">
      <div class="ctrl">
        <div class="ctrl-cab"><label for="clasif">Clasificación de vialidad</label></div>
        <select id="clasif">{opciones}</select>
      </div>
      <div class="ctrl">
        <div class="ctrl-cab"><label for="interpostal">Distancia interpostal</label>
          <output id="vInterpostal"></output></div>
        <input type="range" id="interpostal" min="0" max="{maxS}" step="1" value="{iS}">
        <div class="marcas"><span>{s0} m</span><span>{s1} m</span></div>
      </div>
      <div class="ctrl">
        <div class="ctrl-cab"><label for="altura">Altura de montaje</label>
          <output id="vAltura"></output></div>
        <input type="range" id="altura" min="0" max="{maxH}" step="1" value="{iH}">
        <div class="marcas"><span>{h0} m</span><span>{h1} m</span></div>
      </div>
    </div>
    <p class="avisomov" id="avisomov" hidden>Estás viendo una configuración distinta a la
    del estudio, que es de {sEst} m de interpostal y {hEst} m de altura de montaje.</p>
  </section>

  <section class="ident">{ident}</section>

  <section>
    <h2>Comparativa</h2>
    <div class="tablabox"><table>
      <thead><tr>
        <th>Luminario</th><th class="n">W</th><th class="n">E<sub>prom</sub> lx</th>
        <th class="n">E<sub>min</sub> lx</th><th class="n">Uniformidad</th>
        <th class="n">DPEA W/m²</th><th>Veredicto</th>
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
    camellones</strong>, como pide la sección 6.1 de la norma. Por eso separar
    los postes baja el DPEA y al mismo tiempo baja la iluminancia: los dos
    criterios se mueven en sentidos opuestos.</p>
    <p>Cada combinación de interpostal y altura está calculada de antemano con
    el mismo motor que produjo el estudio; no son valores interpolados. Cambiar
    la clasificación, en cambio, solo mueve los umbrales, porque la iluminancia
    depende de la geometría y de la fotometría, no de cómo se clasifique la
    calle.</p>
  </section>

  <footer>Generado por el motor de cálculo de ilumilm. Los watts de cada
  luminario provienen del archivo IES o fueron capturados a mano; conviene
  confirmarlos contra la ficha del fabricante, porque de ellos depende
  directamente el DPEA.</footer>
</div>
<script>
const RAMPA = {rampa};
const REQ = {req};
const LUM = {lum};
const B = {barrido};
{js}
</script>
""".format(
        css=CSS,
        js=JS,
        titulo=TITULO,
        ident=ident_html,
        opciones=opciones,
        maxS=len(ints) - 1,
        maxH=len(alts) - 1,
        iS=b["i_interpostal_estudio"],
        iH=b["i_altura_estudio"],
        s0=_num(ints[0], 1), s1=_num(ints[-1], 1),
        h0=_num(alts[0], 1), h1=_num(alts[-1], 1),
        sEst=_num(v["interpostal"], 2), hEst=_num(v["altura_montaje"], 2),
        filas="".join(_fila(i, r) for i, r in enumerate(res)),
        detalles="".join(_detalle(i, r, n_x, n_y) for i, r in enumerate(res)),
        rampa=json.dumps(RAMPA),
        req=json.dumps(req, ensure_ascii=False),
        lum=json.dumps(lum, ensure_ascii=False),
        barrido=json.dumps(b, ensure_ascii=False, separators=(",", ":")),
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
    if "barrido" not in datos:
        print("El resultados.json no trae 'barrido'; vuelve a correr "
              "'python -m engine.cli' para regenerarlo.", file=sys.stderr)
        return 1
    destino = ruta.parent / "reporte.html"
    destino.write_text(html(datos), encoding="utf-8")
    print("Reporte escrito en {}".format(destino))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
