"""Lector de archivos fotometricos IESNA LM-63 (1986 / 1991 / 1995 / 2002).

A diferencia de `ReadISO.bas` del SEAD (ver reference/ESPEC_SEAD_IES.md), aqui:

* se conservan `units_type`, `photometric_type`, lumenes y ballast factor como
  metadatos consultables, en vez de leerlos y tirarlos;
* se resuelve la simetria horizontal al consultar la intensidad, para que un
  archivo 0-90 o 0-180 se pueda evaluar en cualquier angulo C;
* se valida que el numero de candelas leidas cuadre con n_vert * n_horiz, que es
  justo la comprobacion que le falta al original y que haria evidente un bloque
  TILT mal saltado.

Sobre el escalado de candelas: en LM-63 los valores del archivo YA son candelas
absolutas. Lo unico que las escala es el multiplicador. Los lumenes por lampara
son informativos (y valen -1 en fotometria absoluta), no un factor; no existe la
division entre 1000 del formato EULUMDAT. El ballast factor del archivo se lee
pero NO se aplica por omision, igual que hace el SEAD, para no contarlo dos
veces con el factor de balastro que el usuario captura en el LLF.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence

# Rangos horizontales canonicos de LM-63 y la simetria que implican.
SIMETRIA_ROTACIONAL = "rotacional"      # un solo plano C: identico en todas direcciones
SIMETRIA_CUADRANTE = "cuadrante"        # 0-90,  espejo en 0 y en 90
SIMETRIA_BILATERAL = "bilateral"        # 0-180, espejo en el eje 0-180
SIMETRIA_NINGUNA = "ninguna"            # 0-360, sin simetria


@dataclass
class Fotometria:
    """Contenido util de un archivo LM-63, ya normalizado."""

    ruta: Path
    keywords: Dict[str, str]
    angulos_v: List[float]                  # grados, 0 = nadir en fotometria tipo C
    angulos_h: List[float]                  # grados
    candelas: List[List[float]]             # [i_horizontal][j_vertical], en cd
    num_lamparas: int
    lumenes_lampara: float                  # -1 => fotometria absoluta
    multiplicador: float
    tipo_fotometrico: int                   # 1=C, 2=B, 3=A
    tipo_unidades: int                      # 1=pies, 2=metros
    dimensiones: tuple                      # (ancho, largo, alto) en las unidades del archivo
    ballast_factor: float
    watts_declarados: float
    simetria: str = field(init=False)

    def __post_init__(self) -> None:
        self.simetria = _clasifica_simetria(self.angulos_h)

    # -- metadatos de conveniencia -----------------------------------------

    @property
    def fabricante(self) -> str:
        return self.keywords.get("MANUFAC", "").strip()

    @property
    def catalogo(self) -> str:
        return (self.keywords.get("LUMCAT") or self.ruta.stem).strip()

    @property
    def absoluta(self) -> bool:
        """True si el archivo declara fotometria absoluta (lumenes = -1)."""
        return self.lumenes_lampara < 0

    @property
    def watts(self) -> Optional[float]:
        """Watts de entrada declarados, o None si el archivo trae 0/negativo.

        Muchos archivos dejan este campo en 0 y ponen la potencia real en texto
        libre. `watts_del_texto` intenta rescatarla, pero el valor SIEMPRE debe
        confirmarlo el usuario: de el depende directamente el DPEA.
        """
        return self.watts_declarados if self.watts_declarados > 0 else None

    def watts_del_texto(self) -> Optional[float]:
        """Busca una potencia en los keywords de texto libre. Heuristica."""
        for clave in ("LUMINAIRE", "LAMP", "BALLAST", "TEST"):
            texto = self.keywords.get(clave, "")
            m = re.search(r"(\d+(?:[.,]\d+)?)\s*[wW](?:atts?)?\b", texto)
            if m:
                return float(m.group(1).replace(",", "."))
        return None

    # -- consulta de intensidad --------------------------------------------

    def intensidad(self, gamma: float, c: float) -> float:
        """Intensidad luminosa en candelas hacia (gamma, C), interpolada.

        `gamma` es el angulo vertical desde el nadir y `c` el horizontal. Se
        resuelve la simetria del archivo antes de interpolar, e interpola
        bilinealmente sobre la retícula, con extrapolacion plana (se satura al
        angulo mas cercano) fuera del rango vertical medido.
        """
        cc = _pliega_c(c, self.simetria, self.angulos_h)
        return _interp_bilineal(self.angulos_h, self.angulos_v, self.candelas, cc, gamma)

    def flujo_luminario(self) -> tuple:
        """Flujo luminoso del luminario en lumenes, y de donde salio.

        Devuelve (lumenes, origen). En fotometria relativa manda lo declarado
        en la cabecera --es el dato del fabricante-- y en fotometria absoluta
        (lumenes = -1, como la mitad del catalogo) no hay declaracion y hay que
        integrar la distribucion.

        Se separa de `flujo_total` porque esa integra SIEMPRE, y para eso
        existe: contrastar lo declarado contra lo medido y cachar un archivo
        mal parseado. Aqui lo que se quiere es el mejor numero disponible para
        ponerlo en el reporte.
        """
        if self.absoluta:
            return self.flujo_total(), "integrado de la distribución"
        declarado = self.lumenes_lampara * self.num_lamparas * self.multiplicador
        return declarado, "declarado en el .ies"

    def flujo_total(self) -> float:
        """Integra la distribucion sobre la esfera, en lumenes.

        Sirve para contrastar contra los lumenes declarados y detectar un
        archivo mal parseado (deberian coincidir dentro de unos pocos por
        ciento cuando el archivo no es de fotometria absoluta).
        """
        v = self.angulos_v
        total = 0.0
        for j, gamma in enumerate(v):
            # ancho del anillo en gamma, por diferencias centradas
            g0 = v[j - 1] if j > 0 else v[0]
            g1 = v[j + 1] if j < len(v) - 1 else v[-1]
            dgamma = math.radians((g1 - g0) / 2.0) if len(v) > 1 else 0.0
            if dgamma <= 0:
                continue
            # intensidad media del anillo sobre todos los planos C
            media = _media_anillo(self, gamma, j)
            total += media * 2.0 * math.pi * math.sin(math.radians(gamma)) * dgamma
        return total


def _media_anillo(f: "Fotometria", gamma: float, j: int) -> float:
    """Promedio de la intensidad sobre 360 grados de C, a gamma fijo."""
    if f.simetria == SIMETRIA_ROTACIONAL:
        return f.candelas[0][j]
    muestras = [f.intensidad(gamma, c) for c in range(0, 360, 5)]
    return sum(muestras) / len(muestras)


# ---------------------------------------------------------------------------
# Simetria
# ---------------------------------------------------------------------------

def _clasifica_simetria(angulos_h: Sequence[float]) -> str:
    if len(angulos_h) <= 1:
        return SIMETRIA_ROTACIONAL
    ultimo = angulos_h[-1]
    if abs(ultimo) < 1e-6:
        return SIMETRIA_ROTACIONAL
    if abs(ultimo - 90.0) < 1e-6:
        return SIMETRIA_CUADRANTE
    if abs(ultimo - 180.0) < 1e-6:
        return SIMETRIA_BILATERAL
    return SIMETRIA_NINGUNA


def _pliega_c(c: float, simetria: str, angulos_h: Sequence[float]) -> float:
    """Lleva un angulo C arbitrario al rango que el archivo si contiene."""
    c = c % 360.0
    if simetria == SIMETRIA_ROTACIONAL:
        return angulos_h[0]
    if simetria == SIMETRIA_BILATERAL:
        # espejo respecto al eje 0-180
        return 360.0 - c if c > 180.0 else c
    if simetria == SIMETRIA_CUADRANTE:
        if c > 180.0:
            c = 360.0 - c
        return 180.0 - c if c > 90.0 else c
    return c


# ---------------------------------------------------------------------------
# Interpolacion bilineal sobre la reticula (C, gamma)
# ---------------------------------------------------------------------------

def _tramo(valores: Sequence[float], x: float) -> tuple:
    """Indice inferior y peso para interpolar x en `valores` (ascendente)."""
    if x <= valores[0]:
        return 0, 0.0
    if x >= valores[-1]:
        return len(valores) - 2 if len(valores) > 1 else 0, 1.0
    lo, hi = 0, len(valores) - 1
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if valores[mid] <= x:
            lo = mid
        else:
            hi = mid
    tramo = valores[hi] - valores[lo]
    return lo, (x - valores[lo]) / tramo if tramo else 0.0


def _interp_bilineal(xs: Sequence[float], ys: Sequence[float],
                     z: Sequence[Sequence[float]], x: float, y: float) -> float:
    if len(xs) == 1:
        i, tx = 0, 0.0
    else:
        i, tx = _tramo(xs, x)
    j, ty = _tramo(ys, y)
    i2 = min(i + 1, len(xs) - 1)
    j2 = min(j + 1, len(ys) - 1)
    z00, z01 = z[i][j], z[i][j2]
    z10, z11 = z[i2][j], z[i2][j2]
    a = z00 + (z01 - z00) * ty
    b = z10 + (z11 - z10) * ty
    return a + (b - a) * tx


# ---------------------------------------------------------------------------
# Lectura del archivo
# ---------------------------------------------------------------------------

_RE_KEYWORD = re.compile(r"^\s*\[([^\]]+)\]\s*(.*)$")


def lee(ruta) -> Fotometria:
    """Parsea un archivo .ies y devuelve la `Fotometria` normalizada."""
    ruta = Path(ruta)
    # latin-1 nunca falla y los archivos suelen traer acentos en cp1252
    lineas = ruta.read_text(encoding="latin-1").splitlines()

    keywords: Dict[str, str] = {}
    i = 0
    # 1) cabecera: primera linea de formato opcional, luego keywords
    while i < len(lineas):
        linea = lineas[i].strip()
        if linea.upper().startswith("TILT"):
            break
        m = _RE_KEYWORD.match(linea)
        if m:
            clave, valor = m.group(1).upper(), m.group(2).strip()
            # [OTHER] aparece muchas veces: se concatena en vez de pisarse
            keywords[clave] = (keywords[clave] + "\n" + valor) if clave in keywords else valor
        i += 1
    else:
        raise ValueError("{}: no se encontro la linea TILT".format(ruta.name))

    tilt = lineas[i].split("=", 1)[1].strip().upper() if "=" in lineas[i] else "NONE"
    i += 1
    if tilt == "INCLUDE":
        # A diferencia del SEAD, que salta 4 lineas a ciegas, se lee de verdad:
        # <orientacion-lampara> <n pares> <n angulos> <n multiplicadores>
        datos = _tokens(lineas, i)
        i = datos.avanza(1)          # lamp-to-luminaire geometry
        n = int(float(datos.siguiente()))
        for _ in range(2 * n):
            datos.siguiente()
        i = datos.linea_actual

    # 2) linea de 10 valores y linea de 3, como flujo de tokens
    flujo = _Flujo(lineas, i)
    num_lamparas = int(float(flujo.siguiente()))
    lumenes = float(flujo.siguiente())
    multiplicador = float(flujo.siguiente())
    n_vert = int(float(flujo.siguiente()))
    n_horiz = int(float(flujo.siguiente()))
    tipo_foto = int(float(flujo.siguiente()))
    tipo_unid = int(float(flujo.siguiente()))
    dims = (float(flujo.siguiente()), float(flujo.siguiente()), float(flujo.siguiente()))
    ballast = float(flujo.siguiente())
    flujo.siguiente()                       # reservado a futuro
    watts = float(flujo.siguiente())

    # 3) angulos y candelas, tambien como flujo (los archivos parten las filas
    #    en cualquier lado, asi que leer por linea no sirve)
    angulos_v = [float(flujo.siguiente()) for _ in range(n_vert)]
    angulos_h = [float(flujo.siguiente()) for _ in range(n_horiz)]

    candelas: List[List[float]] = []
    for _ in range(n_horiz):
        fila = [float(flujo.siguiente()) * multiplicador for _ in range(n_vert)]
        candelas.append(fila)

    if flujo.agotado_prematuramente:
        raise ValueError(
            "{}: el archivo termino antes de completar {}x{} valores de candela. "
            "Cabecera inconsistente o bloque TILT mal formado.".format(
                ruta.name, n_horiz, n_vert)
        )
    if not _ascendente(angulos_v) or not _ascendente(angulos_h):
        raise ValueError("{}: los angulos no vienen en orden ascendente".format(ruta.name))

    return Fotometria(
        ruta=ruta,
        keywords=keywords,
        angulos_v=angulos_v,
        angulos_h=angulos_h,
        candelas=candelas,
        num_lamparas=num_lamparas,
        lumenes_lampara=lumenes,
        multiplicador=multiplicador,
        tipo_fotometrico=tipo_foto,
        tipo_unidades=tipo_unid,
        dimensiones=dims,
        ballast_factor=ballast,
        watts_declarados=watts,
    )


def _ascendente(xs: Sequence[float]) -> bool:
    return all(b >= a for a, b in zip(xs, xs[1:]))


class _Flujo:
    """Tokens numericos consecutivos a partir de una linea dada."""

    def __init__(self, lineas: Sequence[str], inicio: int) -> None:
        self._lineas = lineas
        self._i = inicio
        self._buf: List[str] = []
        self._pos = 0
        self.agotado_prematuramente = False

    def siguiente(self) -> str:
        while self._pos >= len(self._buf):
            if self._i >= len(self._lineas):
                self.agotado_prematuramente = True
                return "0"
            self._buf = self._lineas[self._i].replace(",", " ").split()
            self._pos = 0
            self._i += 1
        tok = self._buf[self._pos]
        self._pos += 1
        return tok


def _tokens(lineas: Sequence[str], inicio: int) -> "_FlujoTilt":
    return _FlujoTilt(lineas, inicio)


class _FlujoTilt(_Flujo):
    """Variante que expone en que linea se quedo, para el bloque TILT."""

    def avanza(self, n: int) -> int:
        for _ in range(n):
            self.siguiente()
        return self.linea_actual

    @property
    def linea_actual(self) -> int:
        return self._i
