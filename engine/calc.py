"""Calculo punto por punto de iluminancia horizontal sobre la vialidad.

Puerto de `AngleCalculations.bas` + `IlluminanceAndLuminance.bas` (metodo IES).

La formula, identica a la del SEAD (`Illum`, IlluminanceAndLuminance.bas:442):

    E = I(gamma, C) * cos(gamma)^3 * LLF / H^2

que es la ley del inverso del cuadrado con la ley del coseno, escrita en
funcion de la altura de montaje H en vez de la distancia oblicua D, ya que
D = H / cos(gamma).

Los angulos, con el luminario en (xl, yl, H) y el punto de calculo en (xp, yp, 0):

    gamma = atan( dist_horizontal / H )        desde el nadir
    C     = azimut en el plano horizontal, 0 = perpendicular a la vialidad,
            90 = a lo largo de la vialidad

Con el luminario inclinado sobre el brazo (ver `_angulos`) hay dos gammas: el
del eje del luminario, que es con el que se consulta la fotometria, y el
geometrico, que es el que entra al cos^3. Con inclinacion 0 coinciden.

AZIMUT
------
El azimut se mide con

    C = atan2( |dx|, dy )           en [0, 180], dy medido hacia la calzada

de modo que C=0 apunta perpendicular a la vialidad hacia la calzada, C=90 corre
a lo largo de ella y C=180 apunta hacia la acera. Respeta el eje de simetria
0-180 de la fotometria vial y distingue el lado de la acera del de la calle, que
es lo que importa cuando el brazo es mayor que el retranqueo y el luminario
queda montado sobre la calzada. Es el criterio de la IESNA RP-8.

Validado contra 7 estudios reales de la herramienta de referencia
(reference/VALIDACION.md): con este criterio el motor reproduce sus resultados
dentro del 0.1 %.

`MODO_CUADRANTE` existe solo para diagnostico. Calcula phi = atan(|dx|/|dy|),
siempre en [0,90], plegando la acera sobre la calle. Corresponde a la funcion
`anglePhi` del VBA original, que resulta NO ser la que alimenta la tabla de
intensidades (esa es `anglePhiWithTilt`, IlluminanceAndLuminance.bas:147).
Desvia hasta un 10 % respecto de los resultados reales; no lo uses para calcular.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import List, Sequence, Tuple

from .geometry import (
    Vialidad,
    posiciones_luminarios,
    valores_x,
    valores_y,
    ventana_evaluacion,
)
from .ies import Fotometria

MODO_CUADRANTE = "cuadrante"   # solo diagnostico, ver docstring
MODO_CORRECTO = "correcto"


@dataclass
class Malla:
    """Resultado del calculo sobre la ventana evaluada."""

    xs: List[float]                  # coordenadas longitudinales evaluadas
    ys: List[float]                  # coordenadas transversales
    e: List[List[float]]             # e[i_x][j_y] en lux
    modo: str

    @property
    def valores(self) -> List[float]:
        return [v for fila in self.e for v in fila]

    @property
    def promedio(self) -> float:
        """Eprom: media aritmetica simple de todos los puntos de la malla.

        Confirmado contra el estado guardado del libro SEAD, hoja
        "Illuminance Calcs" C4 ("Average Illuminance").
        """
        vals = self.valores
        return sum(vals) / len(vals)

    @property
    def minimo(self) -> float:
        return min(self.valores)

    @property
    def maximo(self) -> float:
        return max(self.valores)

    @property
    def uniformidad(self) -> float:
        """Eprom/Emin, la relacion que pide la NOM-013 8.3.

        El SEAD la rotula "Uniformity Ratio (Iavg/Imin)": misma definicion.
        """
        m = self.minimo
        return float("inf") if m <= 0 else self.promedio / m

    def a_csv(self) -> str:
        cab = ["x_m"] + ["y={:.3f}".format(y) for y in self.ys]
        filas = [",".join(cab)]
        for i, x in enumerate(self.xs):
            filas.append(",".join(["{:.3f}".format(x)] +
                                  ["{:.4f}".format(v) for v in self.e[i]]))
        return "\n".join(filas) + "\n"


def _angulos(xl: float, yl: float, orient: int, xp: float, yp: float,
             h: float, modo: str, inclinacion: float = 0.0) -> Tuple[float, float, float]:
    """Devuelve (gamma_geometrico, gamma_tabla, C_tabla) en grados.

    Son DOS gammas y no uno porque el luminario puede estar inclinado, y
    entonces cada uno responde una pregunta distinta:

    * `gamma_tabla` y `C_tabla` se miden respecto al eje del luminario, que es
      como esta medida la fotometria. Sirven para preguntarle al .ies cuanta
      luz sale hacia ese punto.
    * `gamma_geometrico` es el angulo real del rayo contra la vertical. Es el
      que entra al cos^3 de la ley del cosseno, porque el pavimento no sabe
      como quedo orientada la lampara.

    Con inclinacion 0 los dos coinciden y esto se reduce al caso de siempre.
    Es la estructura del VBA 1.8.1 (`IlluminanceAndLuminance.bas:218-222` y
    `:583`), que `ESPEC_SEAD_ILUMINANCIA.md` habia registrado como posible bug
    ("Ambiguedad A1") cuando el tilt todavia estaba muerto. No es un bug.
    """
    dx = xp - xl
    # dy positivo = hacia el lado de la calle DE ESTE luminario
    dy = (yp - yl) * orient
    dist = math.hypot(dx, dy)
    gamma_geo = math.degrees(math.atan(dist / h)) if dist else 0.0

    # Rotacion al sistema del luminario (`angleGammaWithTilt` /
    # `anglePhiWithTilt`, AngleCalculations.bas). La 1.8.1 solo gira sobre X
    # -- el cabeceo del luminario sobre el brazo -- asi que los otros dos ejes
    # de la matriz general quedan en cero y se simplifica a esto.
    if inclinacion:
        t = math.radians(inclinacion)
        cos_t, sin_t = math.cos(t), math.sin(t)
        x_p = dx
        y_p = dy * cos_t - h * sin_t
        h_p = dy * sin_t + h * cos_t
    else:
        x_p, y_p, h_p = dx, dy, h

    r = math.hypot(x_p, y_p)
    if h_p == 0:
        gamma = 90.0
    else:
        gamma = math.degrees(math.atan(r / h_p))
        if h_p < 0:
            # el punto quedo por detras del plano del luminario inclinado
            gamma += 180.0

    if modo == MODO_CUADRANTE:
        # plegado al cuadrante; ver la nota del docstring, no es el criterio de calculo
        c = 90.0 if y_p == 0 else math.degrees(math.atan(abs(x_p) / abs(y_p)))
    else:
        # 0 = hacia la calzada, 180 = hacia la acera
        c = math.degrees(math.atan2(abs(x_p), y_p))
    return gamma_geo, gamma, c


def calcula(v: Vialidad, foto: Fotometria, llf: float,
            modo: str = MODO_CORRECTO, inclinacion: float = 0.0) -> Malla:
    """Iluminancia en cada punto de la ventana evaluada, en lux.

    `llf` es el factor total de perdidas (LLD * LDD * factor de balastro).

    `inclinacion` es el cabeceo del luminario sobre el brazo, en grados;
    positiva apunta hacia la calzada. Es propiedad del luminario y no de la
    vialidad, igual que en el Excel (`FixtureData!selectedTilt`), asi que dos
    luminarios de una misma corrida pueden traer inclinaciones distintas.
    """
    xs_todos = valores_x(v.interpostal)
    ys = valores_y(v.num_carriles, v.ancho_carril, v.camellon)
    i0, i1 = ventana_evaluacion(xs_todos, v.interpostal)
    xs = xs_todos[i0:i1 + 1]
    luminarios = posiciones_luminarios(v)
    h = v.altura_montaje

    e: List[List[float]] = []
    for xp in xs:
        fila: List[float] = []
        for yp in ys:
            total = 0.0
            # Todos los luminarios del tramo contribuyen; el metodo IES del
            # SEAD no aplica ningun corte por distancia.
            for xl, yl, orient in luminarios:
                gamma_geo, gamma, c = _angulos(
                    xl, yl, orient, xp, yp, h, modo, inclinacion)
                intensidad = foto.intensidad(gamma, c)
                cos_g = math.cos(math.radians(gamma_geo))
                total += intensidad * cos_g ** 3 * llf / (h * h)
            fila.append(total)
        e.append(fila)
    return Malla(xs=xs, ys=ys, e=e, modo=modo)


def comparar_modos(v: Vialidad, foto: Fotometria, llf: float,
                   inclinacion: float = 0.0) -> dict:
    """Compara el azimut de calculo contra el plegado al cuadrante.

    Herramienta de diagnostico. El reporte ya no la usa.
    """
    a = calcula(v, foto, llf, MODO_CUADRANTE, inclinacion)
    b = calcula(v, foto, llf, MODO_CORRECTO, inclinacion)
    def dif(x, y):
        return 100.0 * (y - x) / x if x else float("nan")
    return {
        "cuadrante": {"promedio": a.promedio, "minimo": a.minimo,
                 "maximo": a.maximo, "uniformidad": a.uniformidad},
        "correcto": {"promedio": b.promedio, "minimo": b.minimo,
                     "maximo": b.maximo, "uniformidad": b.uniformidad},
        "dif_pct": {"promedio": dif(a.promedio, b.promedio),
                    "minimo": dif(a.minimo, b.minimo),
                    "uniformidad": dif(a.uniformidad, b.uniformidad)},
    }


def llf(lld: float = 0.85, ldd: float = 0.90, bf: float = 1.0) -> float:
    """Factor total de perdidas de luz.

    Los valores por omision son los del tutorial de uso para tecnologia LED:
    depreciacion de lumenes 0.85, depreciacion por suciedad 0.90, y factor de
    balastro 1.0 (el driver ya viene incluido en la potencia nominal del LED).
    """
    return lld * ldd * bf
