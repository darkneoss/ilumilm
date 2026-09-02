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

import math
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


@dataclass(frozen=True)
class Montaje:
    """Como va montado UN luminario en su poste.

    Es propiedad del luminario y no de la vialidad, igual que en el Excel
    (hoja `FixtureData`), asi que dos luminarios comparados en un mismo estudio
    pueden traer montajes distintos. Los tres valores por omision reproducen el
    montaje corriente --un luminario por poste, perpendicular y sin cabeceo--
    que es el de las 24 corridas de referencia.
    """

    inclinacion: float = 0.0            # cabeceo sobre el brazo, grados, + hacia la calzada
    luminarios_por_poste: int = 1       # `selectedFixturesPerPole`
    angulo_separacion: float = 0.0      # entre los brazos de un mismo poste, grados

    # OJO: esto no es el doble brazo del camellon. En "central doble" la
    # disposicion YA cuelga dos luminarios de cada poste, uno por sentido, y el
    # DPEA cuenta dos por tramo sin declarar nada. `luminarios_por_poste` es un
    # multiplicador aparte que aplica a cualquier disposicion, asi que en
    # central doble daria cuatro brazos por poste. El original lo permite igual;
    # ver reference/sead_vba_1.8.1/README.md.

    def __post_init__(self) -> None:
        if not -90.0 < self.inclinacion < 90.0:
            raise ValueError("inclinacion debe estar entre -90 y 90 grados")
        if self.luminarios_por_poste < 1:
            raise ValueError("luminarios_por_poste debe ser >= 1")
        if self.luminarios_por_poste > 2:
            # El VBA reparte los brazos con `ArmLength * sin(sep/2)`, que
            # supone dos brazos simetricos respecto al poste. Su propio
            # comentario lo admite: "Would also work for 3 fixtures if other
            # logic used adjustment correctly". Con tres o mas los pondria
            # todos en dos posiciones, asi que no se acepta en vez de dar un
            # numero que parece bueno.
            raise ValueError(
                "luminarios_por_poste > 2 no esta soportado: el reparto de "
                "brazos del original supone dos brazos simetricos"
            )
        if not -180.0 < self.angulo_separacion < 180.0:
            raise ValueError("angulo_separacion debe estar entre -180 y 180 grados")
        if self.angulo_separacion and self.luminarios_por_poste == 1:
            raise ValueError(
                "el angulo_separacion solo tiene sentido con mas de un "
                "luminario por poste"
            )

    @property
    def neutro(self) -> bool:
        """True si este montaje es el corriente, el de los casos validados."""
        return (not self.inclinacion and self.luminarios_por_poste == 1
                and not self.angulo_separacion)


class Luminario(NamedTuple):
    """Un luminario ya colocado sobre el tramo."""

    x: float
    y: float
    orientacion: int      # +1 mira hacia +Y, -1 hacia -Y
    giro_z: float = 0.0   # giro del brazo sobre la vertical, grados
    x_poste: float = 0.0  # X del poste del que cuelga (== x con un solo brazo)


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

def posiciones_luminarios(v: Vialidad,
                          montaje: "Montaje" = None) -> List[Luminario]:
    """Posicion, orientacion y giro de cada luminario sobre el tramo.

    Puerto del triple bucle de `FixturePosition` (FixturePositions.bas de la
    1.8.1): fase de poste x lado de la calle x luminario del poste. La 1.7.6
    tenia una rama por disposicion; la 1.8.1 las unifico en tres parametros
    --cuantos lados, cuanto se desplaza en X el poste del lado B, y si el poste
    es central-- y aqui se sigue esa estructura, que es la que admite mas de un
    luminario por poste sin duplicar nada.

    `orientacion` vale +1 si el lado de la calle del luminario apunta hacia +Y
    y -1 si apunta hacia -Y (el `facesBackwards` del VBA). Importa porque un
    luminario vial es asimetrico: en unilateral todos miran igual, pero en
    bilateral, tresbolillo y central doble la mitad mira al lado contrario, y
    usar la misma orientacion para todos ilumina la calzada con la distribucion
    de la acera.

    `giro_z` es el giro del brazo sobre el eje vertical, en grados. Solo es
    distinto de cero cuando hay dos luminarios en un poste, y entonces cada uno
    gira hacia un lado.

    RAREZA SEAD: la malla cubre 4 interpostales y la ventana de evaluacion es
    [S, 2S). Es decir, la ventana tiene UN luminario aguas arriba y TRES aguas
    abajo: el arreglo no es simetrico respecto al tramo evaluado. Las
    contribuciones lejanas son despreciables, pero la asimetria es real y esta
    en el original.

    CONTEO: `CInt(gridlength / polespacing) + 1` postes por lado, o sea 5 en
    unilateral (x = 0, S, 2S, 3S, 4S) y 10 repartidos en las demas
    disposiciones, todo por los luminarios que lleve cada poste. La v1.7.6
    generaba uno mas por lado (hasta x = 5S) por un `ReDim FPArrayX(n + 1)` con
    base 0; la 1.8.1 lo corrigio. La diferencia es de ~0.001 % en la mayoria de
    los casos, pero con el conteo nuevo la mitad de los estudios de referencia
    cuadran al 0.0000 % exacto y el peor error de la suite baja de 0.14 % a
    0.007 %, asi que el conteo bueno es este.
    """
    if montaje is None:
        montaje = Montaje()

    s = v.interpostal
    n_por_lado = int(round(longitud_malla(s) / s)) + 1
    por_poste = montaje.luminarios_por_poste

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
        lados, dx_lado_b, central = 2, 0.0, 1
    else:  # pragma: no cover - normaliza_disposicion ya filtro
        raise ValueError(v.disposicion)

    n_total = n_por_lado * por_poste * (1 if lados == 1 else 2)

    # Con dos brazos en un poste, el angulo de separacion los abre en X y les
    # acorta el alcance en Y: el brazo deja de ser perpendicular a la vialidad.
    # Con un solo luminario el angulo es 0, dx_brazo = 0 y dy_brazo = brazo, o
    # sea el caso de siempre.
    sep = math.radians(montaje.angulo_separacion)
    dx_brazo = v.largo_brazo * math.sin(sep / 2.0)
    dy_brazo = v.largo_brazo * math.cos(sep / 2.0)

    ancho = v.num_carriles * v.ancho_carril + v.camellon

    puntos: List[Luminario] = []
    fase = 0
    while len(puntos) < n_total:
        for lado in range(1, lados + 1):
            signo = (-1) ** lado          # -1 en el lado A, +1 en el lado B
            x_poste = fase * s + (lado - 1) * dx_lado_b
            for k in range(1, por_poste + 1):
                x = x_poste + ((-1) ** k) * dx_brazo
                # El retranqueo NO aplica a un poste central: los dos brazos
                # salen del eje del camellon, de ahi el factor (1 - central).
                # El VBA lo suma y lo resta igual, con un comentario propio
                # admitiendo que sobra ("polesetback should probably be removed
                # from here"). Sin esta correccion el error contra los estudios
                # de referencia es de 8 % en Emin; con ella baja a 0.02 %.
                y = (-signo * (1 - 2 * central) * dy_brazo
                     + signo * (1 - central) * v.retranqueo
                     + (lado - 1) * ancho * (1 - central)
                     + central * ancho / 2.0)
                # En el lado A solo mira "hacia atras" si el poste es central;
                # en el lado B, siempre que no lo sea.
                atras = (lado == 1) if central else (lado == 2)
                giro = -montaje.angulo_separacion if k == 1 else montaje.angulo_separacion
                puntos.append(Luminario(x, y, -1 if atras else 1, giro, x_poste))
        fase += 1
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
