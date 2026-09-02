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
from typing import List, NamedTuple, Sequence, Tuple

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
    """Los datos de entrada geometricos de un estudio. Todo en metros.

    `banqueta` es la excepcion: NO entra en ningun calculo. Existe solo para
    que el perfil de la via del reporte este completo. La NOM-013 6.1 excluye
    aceras y camellones del DPEA, y la malla del metodo IES cubre unicamente la
    calzada, asi que ensanchar la acera no puede mover un resultado. Va
    documentado aqui porque un campo de geometria que no afecta la geometria
    del calculo es justo el tipo de cosa que alguien "arregla" por error.
    """

    num_carriles: int
    ancho_carril: float
    camellon: float
    disposicion: str          # una de las claves de DISPOSICIONES
    altura_montaje: float
    interpostal: float        # distancia entre postes consecutivos del mismo lado
    retranqueo: float         # "pole setback": del poste a la orilla de la calzada
    largo_brazo: float
    banqueta: float = 0.0     # ancho de acera a cada lado; SOLO para el perfil dibujado

    def __post_init__(self) -> None:
        object.__setattr__(self, "disposicion", normaliza_disposicion(self.disposicion))
        if self.num_carriles < 1:
            raise ValueError("num_carriles debe ser >= 1")
        for campo in ("ancho_carril", "altura_montaje", "interpostal"):
            if getattr(self, campo) <= 0:
                raise ValueError("{} debe ser > 0".format(campo))
        if self.banqueta < 0:
            raise ValueError("banqueta no puede ser negativa")

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


class Luminario(NamedTuple):
    """Un luminario ya colocado sobre el tramo."""

    x: float
    y: float
    orientacion: int      # +1 mira hacia +Y, -1 hacia -Y


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

def posiciones_luminarios(v: Vialidad) -> List[Luminario]:
    """Posicion y orientacion de cada luminario sobre el tramo.

    Puerto del doble bucle de `FixturePosition` (FixturePositions.bas de la
    1.8.1): fase de poste x lado de la calle. La 1.7.6 tenia una rama por
    disposicion; la 1.8.1 las unifico en tres parametros --cuantos lados,
    cuanto se desplaza en X el poste del lado B, y si el poste es central-- y
    aqui se sigue esa estructura, que dice mas de las disposiciones que cuatro
    ramas separadas: lo unico que las distingue son esos tres numeros.

    `orientacion` vale +1 si el lado de la calle del luminario apunta hacia +Y
    y -1 si apunta hacia -Y (el `facesBackwards` del VBA). Importa porque un
    luminario vial es asimetrico: en unilateral todos miran igual, pero en
    bilateral, tresbolillo y central doble la mitad mira al lado contrario, y
    usar la misma orientacion para todos ilumina la calzada con la distribucion
    de la acera.

    DOS LUMINARIOS POR POSTE: los hay, pero solo en central doble, y los pone
    esta funcion sola --son los dos brazos que salen del eje del camellon, uno
    por sentido--. El `selectedFixturesPerPole` del original, que multiplicaria
    los brazos en cualquier disposicion, NO se implementa: la interfaz del Excel
    no lo ofrece de forma usable, ninguna corrida de referencia lo usa y en
    central doble se confundiria con el doble brazo, cuadruplicando el DPEA.
    Ver reference/sead_vba_1.8.1/README.md.

    RAREZA SEAD: la malla cubre 4 interpostales y la ventana de evaluacion es
    [S, 2S). Es decir, la ventana tiene UN luminario aguas arriba y TRES aguas
    abajo: el arreglo no es simetrico respecto al tramo evaluado. Las
    contribuciones lejanas son despreciables, pero la asimetria es real y esta
    en el original.

    CONTEO: `CInt(gridlength / polespacing) + 1` postes por lado, o sea 5 en
    unilateral (x = 0, S, 2S, 3S, 4S) y 10 repartidos en las demas
    disposiciones. La v1.7.6 generaba uno mas por lado (hasta x = 5S) por un
    `ReDim FPArrayX(n + 1)` con base 0; la 1.8.1 lo corrigio. La diferencia es
    de ~0.001 % en la mayoria de los casos, pero con el conteo nuevo la mitad
    de los estudios de referencia cuadran al 0.0000 % exacto y el peor error de
    la suite baja de 0.14 % a 0.007 %, asi que el conteo bueno es este.
    """
    s = v.interpostal
    n_por_lado = int(round(longitud_malla(s) / s)) + 1

    # Los tres parametros con que la 1.8.1 describe cualquier disposicion.
    if v.disposicion == "Single-side":
        lados, dx_lado_b, central = 1, 0.0, 0
    elif v.disposicion == "Opposite":
        # Postes enfrentados: los dos lados comparten la misma X.
        lados, dx_lado_b, central = 2, 0.0, 0
    elif v.disposicion == "Staggered":
        # Alternados: el lado B va corrido medio interpostal.
        lados, dx_lado_b, central = 2, s / 2.0, 0
    elif v.disposicion == "Median mounted":
        # Un solo poste al centro con dos brazos, uno hacia cada sentido: los
        # "dos lados" son aqui los dos brazos del mismo poste.
        lados, dx_lado_b, central = 2, 0.0, 1
    else:  # pragma: no cover - normaliza_disposicion ya filtro
        raise ValueError(v.disposicion)

    ancho = v.num_carriles * v.ancho_carril + v.camellon
    puntos: List[Luminario] = []
    for fase in range(n_por_lado):
        for lado in range(1, lados + 1):
            signo = (-1) ** lado          # -1 en el lado A, +1 en el lado B
            x = fase * s + (lado - 1) * dx_lado_b
            # El retranqueo NO aplica a un poste central: los dos brazos salen
            # del eje del camellon, de ahi el factor (1 - central). El VBA lo
            # suma y lo resta igual, con un comentario propio admitiendo que
            # sobra ("polesetback should probably be removed from here"). Sin
            # esta correccion el error contra los estudios de referencia es de
            # 8 % en Emin; con ella baja a 0.02 %.
            y = (-signo * (1 - 2 * central) * v.largo_brazo
                 + signo * (1 - central) * v.retranqueo
                 + (lado - 1) * ancho * (1 - central)
                 + central * ancho / 2.0)
            # En el lado A solo mira "hacia atras" si el poste es central; en
            # el lado B, siempre que no lo sea.
            atras = (lado == 1) if central else (lado == 2)
            puntos.append(Luminario(x, y, -1 if atras else 1))
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
