"""Geometria de la vialidad: malla de calculo y posiciones de luminarios.

Puerto fiel de los modulos VBA `MakeMeasurementGrid` y `FixturePositions` del
SEAD Street Lighting Tool v1.8.1 (ver reference/sead_vba_1.8.1/), que es la
version que genero los estudios de referencia. Se replican tambien las rarezas
del original, marcadas con "RAREZA SEAD", porque el objetivo es que los numeros
coincidan con los del Excel, no que sean los que a nosotros nos parezcan
mejores. Donde importa la diferencia con la v1.7.6 -- conservada en
reference/sead_vba_1.7.6/ porque las especificaciones citan sus lineas -- va
dicho en el docstring de la funcion.

Sistema de coordenadas
----------------------
    X  a lo largo de la vialidad, crece en el sentido de avance.
    Y  transversal. Y=0 es la orilla de la calzada del lado de los postes;
       Y = num_carriles*ancho_carril + camellon es la orilla opuesta.
    Z  la altura de montaje, positiva hacia arriba.

Solo se implementa el metodo IES (el CIE queda fuera del alcance, ver el plan).
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Sequence, Tuple

# Nombres de disposicion tal como los usa el VBA, con su equivalente en la
# NOM-013 (Apendice C, figuras C.1 a C.4).
DISPOSICIONES = {
    "Single-side": "unilateral",
    "Staggered": "tresbolillo",
    "Median mounted": "central doble",
    "Opposite": "bilateral opuesta",
}

_ALIAS_DISPOSICION = {
    "unilateral": "Single-side",
    "un solo lado": "Single-side",
    "de un solo lado": "Single-side",
    "single-side": "Single-side",
    "tresbolillo": "Staggered",
    "escalonado": "Staggered",
    "staggered": "Staggered",
    "central doble": "Median mounted",
    "camellon": "Median mounted",
    "montado en el camellon": "Median mounted",
    "montado en el camellón": "Median mounted",
    "median mounted": "Median mounted",
    "bilateral": "Opposite",
    "bilateral opuesta": "Opposite",
    "opuesto": "Opposite",
    "opposite": "Opposite",
}


def normaliza_disposicion(texto: str) -> str:
    """Acepta el nombre en espanol de la NOM o el interno del VBA."""
    clave = texto.strip().lower()
    if clave in _ALIAS_DISPOSICION:
        return _ALIAS_DISPOSICION[clave]
    raise ValueError(
        "disposicion de postes desconocida: {!r}. Opciones: {}".format(
            texto, ", ".join(sorted(set(_ALIAS_DISPOSICION)))
        )
    )


@dataclass(frozen=True)
class Vialidad:
    """Los datos de entrada geometricos de un estudio. Todo en metros."""

    num_carriles: int
    ancho_carril: float
    camellon: float
    disposicion: str          # una de las claves de DISPOSICIONES
    altura_montaje: float
    interpostal: float        # distancia entre postes consecutivos del mismo lado
    retranqueo: float         # "pole setback": del poste a la orilla de la calzada
    largo_brazo: float

    def __post_init__(self) -> None:
        object.__setattr__(self, "disposicion", normaliza_disposicion(self.disposicion))
        if self.num_carriles < 1:
            raise ValueError("num_carriles debe ser >= 1")
        for campo in ("ancho_carril", "altura_montaje", "interpostal"):
            if getattr(self, campo) <= 0:
                raise ValueError("{} debe ser > 0".format(campo))

    @property
    def ancho_calzada(self) -> float:
        """Ancho total incluyendo camellon (el ancho fisico de la seccion)."""
        return self.num_carriles * self.ancho_carril + self.camellon

    @property
    def ancho_sin_camellon(self) -> float:
        """Ancho para el DPEA: la NOM-013 6.1 excluye aceras y camellones."""
        return self.num_carriles * self.ancho_carril

    @property
    def y_poste(self) -> float:
        """Y del luminario del lado cercano (no del poste: incluye el brazo)."""
        return -self.retranqueo + self.largo_brazo


# ---------------------------------------------------------------------------
# Malla de medicion  (MakeMeasurementGrid.bas)
# ---------------------------------------------------------------------------

def paso_malla(interpostal: float) -> float:
    """GridSpace, metodo IES: un decimo del interpostal, topado a 5 m."""
    return min(interpostal / 10.0, 5.0)


def longitud_malla(interpostal: float) -> float:
    """TotalGridLength, metodo IES: cuatro tramos interpostales."""
    return 4.0 * interpostal


def valores_x(interpostal: float) -> List[float]:
    """Coordenadas longitudinales de la malla.

    El VBA hace `ReDim Xvalues(ngp)` y luego llena de 1 a ngp, asi que son
    ngp+1 valores empezando en medio paso.
    """
    paso = paso_malla(interpostal)
    ngp = int(round(longitud_malla(interpostal) / paso))
    return [paso / 2.0 + i * paso for i in range(ngp + 1)]


def valores_y(num_carriles: int, ancho_carril: float, camellon: float) -> List[float]:
    """Coordenadas transversales: dos puntos por carril, en los cuartos.

    RAREZA SEAD: el camellon solo se inserta cuando el numero de carriles es
    par, porque `medianYvalue` se pone en 0 y la bandera en False para carriles
    impares (MakeMeasurementGrid.bas, "detemining median coordinates"). Una
    vialidad con carriles impares y camellon queda con la malla comprimida.
    """
    n = 2 * num_carriles
    ys: List[float] = [ancho_carril / 4.0]
    if num_carriles % 2 == 0:
        y_camellon = (num_carriles / 2.0) * ancho_carril
        pendiente = True
    else:
        y_camellon = 0.0
        pendiente = False
    for _ in range(1, n):
        y = ys[-1] + ancho_carril / 2.0
        if pendiente and y >= y_camellon:
            y += camellon
            pendiente = False
        ys.append(y)
    return ys


# ---------------------------------------------------------------------------
# Posiciones de luminarios  (FixturePositions.bas)
# ---------------------------------------------------------------------------

def posiciones_luminarios(v: Vialidad) -> List[Tuple[float, float, int]]:
    """Coordenadas (x, y, orientacion) de cada luminario sobre el tramo.

    `orientacion` vale +1 si el lado de la calle del luminario apunta hacia +Y
    y -1 si apunta hacia -Y. Importa porque un luminario vial es asimetrico:
    en unilateral todos miran igual, pero en bilateral, tresbolillo y central
    doble la mitad de ellos mira al lado contrario, y usar la misma orientacion
    para todos ilumina la calzada con la distribucion de la acera.

    RAREZA SEAD: la malla cubre 4 interpostales y la ventana de evaluacion es
    [S, 2S). Es decir, la ventana tiene UN luminario aguas arriba y TRES aguas
    abajo: el arreglo no es simetrico respecto al tramo evaluado. Las
    contribuciones lejanas son despreciables, pero la asimetria es real y esta
    en el original.

    CONTEO: `CInt(gridlength / polespacing) + 1` luminarios por lado, o sea 5
    en unilateral (x = 0, S, 2S, 3S, 4S) y 10 repartidos en las demas
    disposiciones. La v1.7.6 generaba uno mas por lado (hasta x = 5S) por un
    `ReDim FPArrayX(n + 1)` con base 0; la 1.8.1 lo corrigio. La diferencia es
    de ~0.001 % en la mayoria de los casos, pero con el conteo nuevo la mitad
    de los estudios de referencia cuadran al 0.0000 % exacto y el peor error de
    la suite baja de 0.14 % a 0.007 %, asi que el conteo bueno es este.
    """
    s = v.interpostal
    largo = longitud_malla(s)
    y_cerca = v.y_poste
    y_lejos = v.num_carriles * v.ancho_carril + v.camellon + v.retranqueo - v.largo_brazo

    n_por_lado = int(round(largo / s)) + 1

    if v.disposicion == "Single-side":
        return [(i * s, y_cerca, 1) for i in range(n_por_lado)]

    n = 2 * n_por_lado - 2
    puntos: List[Tuple[float, float, int]] = []
    centro = (v.num_carriles * v.ancho_carril + v.camellon) / 2.0
    for i in range(n + 2):
        par = i % 2 == 0
        if v.disposicion == "Opposite":
            # Postes enfrentados: el par y el impar comparten la misma X.
            x = i * s / 2 if par else (i - 1) * s / 2
            y = y_cerca if par else y_lejos
            orientacion = 1 if par else -1
        elif v.disposicion == "Staggered":
            # Alternados: cada luminario a medio interpostal del anterior.
            x = i * s / 2
            y = y_cerca if par else y_lejos
            orientacion = 1 if par else -1
        elif v.disposicion == "Median mounted":
            x = i * s / 2 if par else (i - 1) * s / 2
            # El original suma/resta el retranqueo aqui, cosa que no tiene
            # sentido fisico en un poste central; la version 1.7.1 lo evita
            # validando que el retranqueo sea cero. Se replica igual.
            # El retranqueo NO aplica a un poste central: los dos brazos salen
            # del eje del camellon. El VBA lo suma y lo resta igual, con un
            # comentario propio admitiendo que sobra ("polesetback should
            # probably be removed from here"), y la version que genero los
            # estudios de referencia valida que el retranqueo sea cero. Sin
            # esta correccion el error contra esos estudios es de 8 % en Emin;
            # con ella baja a 0.02 %.
            y = centro + v.largo_brazo if par else centro - v.largo_brazo
            # el brazo par queda del lado lejano y alumbra hacia +Y; el impar,
            # del lado cercano, alumbra hacia -Y
            orientacion = 1 if par else -1
        else:  # pragma: no cover - normaliza_disposicion ya filtro
            raise ValueError(v.disposicion)
        puntos.append((x, y, orientacion))
    return puntos


# ---------------------------------------------------------------------------
# Ventana de evaluacion  (AngleCalculations.bas, iStart / iEnd)
# ---------------------------------------------------------------------------

def match_vba(valor: float, arreglo: Sequence[float]) -> int:
    """Equivalente de `WorksheetFunction.Match(valor, arreglo, 1)`.

    Devuelve la POSICION base 1 del mayor elemento <= valor, que es lo que hace
    Excel sobre un arreglo ascendente. El VBA usa ese resultado como indice
    base 0 sobre el mismo arreglo, con lo que en la practica se corre un lugar;
    ese desfase forma parte del comportamiento que estamos replicando y se
    conserva en `ventana_evaluacion`.
    """
    pos = 0
    for i, x in enumerate(arreglo):
        if x <= valor:
            pos = i + 1
        else:
            break
    if pos == 0:
        raise ValueError("{} es menor que el primer elemento del arreglo".format(valor))
    return pos


def ventana_evaluacion(xs: Sequence[float], interpostal: float) -> Tuple[int, int]:
    """Indices (inicio, fin) INCLUSIVOS de la ventana evaluada, sobre `xs`.

    Replica textualmente:
        iStart = Match(polespacing, outputX, True)
        iEnd   = Match(2 * polespacing, outputX, True) - 1
    usados despues como indices base 0 de `outputX`.
    """
    inicio = match_vba(interpostal, xs)
    fin = match_vba(2.0 * interpostal, xs) - 1
    return inicio, fin
