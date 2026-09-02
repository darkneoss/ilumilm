"""Esquemas acotados de la vialidad: seccion transversal y planta.

Dos dibujos, cada uno para una pregunta:

* `seccion` es la vista de frente. Responde a que altura va el luminario, cuanto
  sale el brazo sobre la calzada y como queda cabeceado. Va **a escala** en los
  dos ejes, con la misma cantidad de pixeles por metro en horizontal y en
  vertical, asi que las proporciones que se ven son las reales.
* `planta` es la vista de arriba de un tramo interpostal. Responde donde caen
  los postes a lo largo de la vialidad, que es lo que distingue una disposicion
  de otra. El eje transversal va a escala; el longitudinal esta comprimido a un
  tramo, porque de otro modo un interpostal de 40 m contra una calzada de 10
  daria una tira ilegible. Va dicho en el pie del dibujo.

Ninguno de los dos dibuja iluminancia: para eso esta el mapa isolux del detalle
por luminario. Estos son de cotas.

Los dos se emiten con las medidas del estudio y llevan `data-` con lo que el
script del reporte necesita para actualizarlos cuando alguien mueve la
interpostal o la altura de montaje. Sin eso el dibujo contradiria a los
controles, que es peor que no tener dibujo.
"""
from __future__ import annotations

import math
from typing import List

from .geometry import Vialidad

# Lienzo. El ancho es fijo; el alto se calcula segun lo que haya que meter.
ANCHO = 720.0
MARGEN_IZQ = 66.0        # sitio para la cota de altura
MARGEN_DER = 30.0
UTIL = ANCHO - MARGEN_IZQ - MARGEN_DER

# Tope de pixeles por metro. Sin el, una vialidad angosta se dibuja enorme y
# el poste se sale del lienzo cuando el usuario sube la altura al maximo.
K_MAX = 26.0
ALTURA_MAX_M = 12.0     # el tope del control de altura del reporte


def _k(ancho_m: float) -> float:
    return min(UTIL / ancho_m, K_MAX) if ancho_m > 0 else K_MAX


def _cota_h(x1: float, x2: float, y: float, texto: str, clase: str = "",
            ancla: str = "middle") -> str:
    """Cota horizontal con sus dos remates y la etiqueta encima.

    `ancla` a "start" para las cotas cortas, donde la etiqueta es mas ancha que
    el tramo acotado y centrada se sale por los dos lados.
    """
    xm = (x1 + x2) / 2.0 if ancla == "middle" else x2 + 8
    return (
        '<g class="cota {cl}">'
        '<line x1="{x1:.1f}" y1="{y:.1f}" x2="{x2:.1f}" y2="{y:.1f}"/>'
        '<line class="rem" x1="{x1:.1f}" y1="{ya:.1f}" x2="{x1:.1f}" y2="{yb:.1f}"/>'
        '<line class="rem" x1="{x2:.1f}" y1="{ya:.1f}" x2="{x2:.1f}" y2="{yb:.1f}"/>'
        '<text x="{xm:.1f}" y="{yt:.1f}" style="text-anchor:{an}">{t}</text>'
        "</g>"
    ).format(x1=x1, x2=x2, y=y, ya=y - 4, yb=y + 4, xm=xm,
             yt=y - 6 if ancla == "middle" else y + 4, t=texto, cl=clase, an=ancla)


def _cota_v(x: float, y1: float, y2: float, texto: str, clase: str = "") -> str:
    """Cota vertical, con la etiqueta girada a la izquierda de la linea."""
    ym = (y1 + y2) / 2.0
    return (
        '<g class="cota {cl}">'
        '<line x1="{x:.1f}" y1="{y1:.1f}" x2="{x:.1f}" y2="{y2:.1f}"/>'
        '<line class="rem" x1="{xa:.1f}" y1="{y1:.1f}" x2="{xb:.1f}" y2="{y1:.1f}"/>'
        '<line class="rem" x1="{xa:.1f}" y1="{y2:.1f}" x2="{xb:.1f}" y2="{y2:.1f}"/>'
        '<text x="{xt:.1f}" y="{ym:.1f}" transform="rotate(-90 {xt:.1f} {ym:.1f})">{t}</text>'
        "</g>"
    ).format(x=x, y1=y1, y2=y2, xa=x - 4, xb=x + 4, xt=x - 7, ym=ym, t=texto,
             cl=clase)


def _m(valor: float, dec: int = 2) -> str:
    return "{:.{}f} m".format(valor, dec)


# ---------------------------------------------------------------------------
# Seccion transversal
# ---------------------------------------------------------------------------

def seccion(v: Vialidad, inclinacion: float = 0.0) -> str:
    """Vista de frente, a escala. Y=0 es la orilla del lado de los postes.

    El origen vertical del dibujo es el SUELO, y se crece hacia arriba en
    negativo. Asi el lienzo se ajusta a la altura de montaje --sin reservar
    hueco para el tope del control, que dejaba media figura vacia-- y para
    seguir al control basta con mover cuatro elementos y estirar el `viewBox`,
    en vez de rehacer el dibujo.
    """
    central = v.disposicion == "Median mounted"
    ancho_total = v.ancho_calzada + 2 * v.banqueta
    k = _k(ancho_total)

    # centrado: con calzadas angostas el tope de escala deja el dibujo pegado
    # a la izquierda y media figura vacia a la derecha
    izq = MARGEN_IZQ + max(0.0, (UTIL - ancho_total * k) / 2.0)

    def ex(y_m: float) -> float:
        """Y de la vialidad (m) -> X del dibujo (px)."""
        return izq + (y_m + v.banqueta) * k

    def ey(z_m: float) -> float:
        """Altura sobre el suelo (m) -> Y del dibujo (px), negativa hacia arriba."""
        return -z_m * k

    arriba = -(v.altura_montaje * k + 26.0)
    alto = -arriba + 78.0

    p: List[str] = [
        '<svg viewBox="0 {:.1f} {:.0f} {:.1f}" class="esquema" data-seccion="1" '
        'data-k="{:.4f}" role="img" aria-label="Seccion transversal de la '
        'vialidad con el poste, el brazo y el luminario acotados">'.format(
            arriba, ANCHO, alto, k)
    ]

    x_calz0, x_calz1 = ex(0.0), ex(v.ancho_calzada)
    if v.banqueta:
        for x0, x1 in ((ex(-v.banqueta), x_calz0),
                       (x_calz1, ex(v.ancho_calzada + v.banqueta))):
            p.append('<rect class="acera" x="{:.1f}" y="-7" width="{:.1f}" '
                     'height="7"/>'.format(x0, x1 - x0))
    p.append('<rect class="asfalto" x="{:.1f}" y="-5" width="{:.1f}" '
             'height="5"/>'.format(x_calz0, x_calz1 - x_calz0))
    if v.camellon:
        y0 = v.num_carriles * v.ancho_carril / 2.0
        p.append('<rect class="camellon" x="{:.1f}" y="-11" width="{:.1f}" '
                 'height="11"/>'.format(ex(y0), v.camellon * k))
    for i in range(1, v.num_carriles):
        y_m = i * v.ancho_carril
        if v.camellon and i == v.num_carriles / 2.0:
            continue
        if v.camellon and y_m > v.num_carriles * v.ancho_carril / 2.0:
            y_m += v.camellon
        p.append('<line class="raya" x1="{x:.1f}" y1="-5" x2="{x:.1f}" '
                 'y2="-16"/>'.format(x=ex(y_m)))

    brazos = ([(v.ancho_calzada / 2.0, +1), (v.ancho_calzada / 2.0, -1)] if central
              else [(-v.retranqueo, +1)])
    for y_base, sentido in brazos:
        x_poste = ex(y_base)
        y_lum = (y_base + sentido * v.largo_brazo) if central else v.y_poste
        giro = -sentido * inclinacion
        p.append('<line class="poste" data-poste="1" x1="{x:.1f}" y1="0" '
                 'x2="{x:.1f}" y2="{y:.1f}"/>'.format(x=x_poste, y=ey(v.altura_montaje)))
        p.append('<line class="brazo" data-brazo="1" x1="{x0:.1f}" y1="{y:.1f}" '
                 'x2="{x1:.1f}" y2="{y:.1f}"/>'.format(
                     x0=x_poste, y=ey(v.altura_montaje), x1=ex(y_lum)))
        p.append('<g class="lum" data-lum="1" data-x="{x:.1f}" data-giro="{g:.1f}" '
                 'transform="translate({x:.1f} {y:.1f}) rotate({g:.1f})">'
                 '<rect x="-11" y="-3.5" width="22" height="7" rx="2"/>'
                 '<line class="haz" x1="0" y1="4" x2="0" y2="16"/></g>'.format(
                     x=ex(y_lum), y=ey(v.altura_montaje), g=giro))

    p.append(_cota_v(MARGEN_IZQ - 22, ey(v.altura_montaje), 0.0,
                     _m(v.altura_montaje), "cota-altura"))
    # La cota abarca la seccion completa, camellon incluido, asi que la
    # ecuacion tiene que nombrarlo: "2 carriles x 3.50 m = 8.00 m" no cuadra.
    etiqueta = "{} carriles &#215; {}".format(v.num_carriles, _m(v.ancho_carril))
    if v.camellon:
        etiqueta += " + {} de camellón".format(_m(v.camellon))
    p.append(_cota_h(x_calz0, x_calz1, 26.0,
                     "{} = {}".format(etiqueta, _m(v.ancho_calzada))))
    if v.banqueta:
        p.append(_cota_h(ex(-v.banqueta), x_calz0, 26.0, _m(v.banqueta)))
    if not central and abs(v.y_poste) > 1e-9:
        p.append(_cota_h(min(x_calz0, ex(v.y_poste)), max(x_calz0, ex(v.y_poste)),
                         52.0, _m(v.y_poste, 2) + " de saliente sobre la calzada",
                         ancla="start"))
    p.append('<text class="pie" x="{:.1f}" y="{:.1f}">Sección transversal, a escala. '
             'La altura de montaje sigue al control.</text>'.format(
                 MARGEN_IZQ, alto + arriba - 8))
    p.append("</svg>")
    return "".join(p)


# ---------------------------------------------------------------------------
# Planta
# ---------------------------------------------------------------------------

def planta(v: Vialidad) -> str:
    """Vista de arriba de un tramo interpostal completo."""
    from .geometry import posiciones_luminarios

    ancho_total = v.ancho_calzada + 2 * v.banqueta
    k = min(_k(ancho_total), 22.0)
    largo_px = UTIL
    y0 = 34.0                                   # arriba del dibujo
    alto = y0 + ancho_total * k + 76.0

    def ex(x_m: float) -> float:
        """X a lo largo de la vialidad (m) -> px, comprimido a un tramo."""
        return MARGEN_IZQ + (x_m / v.interpostal) * largo_px

    def ey(y_m: float) -> float:
        return y0 + (y_m + v.banqueta) * k

    p: List[str] = [
        '<svg viewBox="0 0 {:.0f} {:.0f}" class="esquema" role="img" '
        'aria-label="Planta de un tramo interpostal con la posicion de los '
        'postes y sus cotas">'.format(ANCHO, alto)
    ]
    x0, x1 = ex(0.0), ex(v.interpostal)
    if v.banqueta:
        for a, b in ((-v.banqueta, 0.0), (v.ancho_calzada, v.ancho_calzada + v.banqueta)):
            p.append('<rect class="acera" x="{:.1f}" y="{:.1f}" width="{:.1f}" '
                     'height="{:.1f}"/>'.format(x0, ey(a), x1 - x0, (b - a) * k))
    p.append('<rect class="asfalto" x="{:.1f}" y="{:.1f}" width="{:.1f}" '
             'height="{:.1f}"/>'.format(x0, ey(0.0), x1 - x0, v.ancho_calzada * k))
    if v.camellon:
        yc = v.num_carriles * v.ancho_carril / 2.0
        p.append('<rect class="camellon" x="{:.1f}" y="{:.1f}" width="{:.1f}" '
                 'height="{:.1f}"/>'.format(x0, ey(yc), x1 - x0, v.camellon * k))
    for i in range(1, v.num_carriles):
        y_m = i * v.ancho_carril
        if v.camellon and i == v.num_carriles / 2.0:
            continue
        if v.camellon and y_m > v.num_carriles * v.ancho_carril / 2.0:
            y_m += v.camellon
        p.append('<line class="raya" x1="{x0:.1f}" y1="{y:.1f}" x2="{x1:.1f}" '
                 'y2="{y:.1f}"/>'.format(x0=x0, x1=x1, y=ey(y_m)))

    # Postes del tramo dibujado. Se toma el arreglo real y se traslada al
    # origen: las posiciones dentro de un tramo se repiten cada interpostal,
    # asi que basta el modulo. Por eso el dibujo no cambia al mover la
    # interpostal, igual que el mapa isolux, y solo la cota se actualiza.
    vistos = set()
    for lum in posiciones_luminarios(v):
        x_m = round(lum.x % v.interpostal, 6)
        clave = (x_m, round(lum.y, 6))
        if clave in vistos:
            continue
        vistos.add(clave)
        for xx in ({x_m, 0.0} if x_m == 0.0 else {x_m}):
            x = ex(xx)
            y = ey(lum.y)
            central = v.disposicion == "Median mounted"
            y_base = (v.ancho_calzada / 2.0 if central
                      else (-v.retranqueo if lum.orientacion == 1
                            else v.ancho_calzada + v.retranqueo))
            p.append('<line class="brazo" x1="{x:.1f}" y1="{yb:.1f}" x2="{x:.1f}" '
                     'y2="{y:.1f}"/>'.format(x=x, yb=ey(y_base), y=y))
            p.append('<rect class="base" x="{:.1f}" y="{:.1f}" width="7" '
                     'height="7"/>'.format(x - 3.5, ey(y_base) - 3.5))
            p.append('<circle class="lumpt" cx="{:.1f}" cy="{:.1f}" r="4.5"/>'.format(x, y))
    # el poste del final del tramo cierra el dibujo
    for lum in posiciones_luminarios(v):
        if round(lum.x % v.interpostal, 6) == 0.0:
            y = ey(lum.y)
            central = v.disposicion == "Median mounted"
            y_base = (v.ancho_calzada / 2.0 if central
                      else (-v.retranqueo if lum.orientacion == 1
                            else v.ancho_calzada + v.retranqueo))
            p.append('<line class="brazo" x1="{x:.1f}" y1="{yb:.1f}" x2="{x:.1f}" '
                     'y2="{y:.1f}"/>'.format(x=x1, yb=ey(y_base), y=y))
            p.append('<rect class="base" x="{:.1f}" y="{:.1f}" width="7" '
                     'height="7"/>'.format(x1 - 3.5, ey(y_base) - 3.5))
            p.append('<circle class="lumpt" cx="{:.1f}" cy="{:.1f}" r="4.5"/>'.format(x1, y))

    p.append(_cota_h(x0, x1, alto - 44, _m(v.interpostal) + " entre postes",
                     "cota-s"))
    p.append(_cota_v(ANCHO - 12, ey(0.0), ey(v.ancho_calzada), _m(v.ancho_calzada)))
    p.append('<text class="pie" x="{:.1f}" y="{:.1f}">Planta de un tramo. El ancho '
             'va a escala; el largo está comprimido al interpostal, que sigue al '
             'control.</text>'.format(MARGEN_IZQ, alto - 8))
    p.append("</svg>")
    return "".join(p)


CSS = """
.esquema{width:100%; height:auto; display:block; overflow:visible}
.esquema .acera{fill:var(--surface-2); stroke:var(--line)}
.esquema .asfalto{fill:var(--ink-3); opacity:.35}
.esquema .camellon{fill:var(--ok); opacity:.30}
.esquema .raya{stroke:var(--surface); stroke-width:1.4; stroke-dasharray:7 7}
.esquema .poste{stroke:var(--ink-2); stroke-width:3; stroke-linecap:round}
.esquema .brazo{stroke:var(--ink-2); stroke-width:2}
.esquema .base{fill:var(--ink-2)}
.esquema .lum rect{fill:var(--amber); stroke:var(--amber-txt); stroke-width:.8}
.esquema .lum .haz{stroke:var(--amber); stroke-width:2; stroke-linecap:round; opacity:.75}
.esquema .lumpt{fill:var(--amber); stroke:var(--amber-txt); stroke-width:.8}
.esquema .cota line{stroke:var(--slate); stroke-width:1}
.esquema .cota .rem{stroke-width:1.4}
.esquema .cota text{
  fill:var(--slate); font-family:"IBM Plex Mono",ui-monospace,monospace;
  font-size:12px; text-anchor:middle;
}
.esquema .pie{
  fill:var(--ink-3); font-family:Archivo,sans-serif; font-size:11px;
}
"""
