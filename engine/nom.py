"""Datos y veredicto de cumplimiento de la NOM-013-ENER-2013.

Norma Oficial Mexicana "Eficiencia Energética para Sistemas de Alumbrado en
Vialidades". Este módulo es autocontenido: no depende de ningún otro módulo
de `engine`. Fuente de verdad: assets/NOM-013-ENER-2013.md (Tablas 1, 2, 3 y
la que en el documento aparece rotulada "Tabla 20"; sección 6.1; 8.1-8.3;
Apéndice D).

Nota sobre una errata del documento fuente: la tabla de valores para
pavimento tipo R3 aparece titulada como "Tabla 20" en la transcripción a
Markdown del documento oficial (probable error de OCR/transcripción de
"Tabla 2 bis" o similar, ya que la numeración salta de la Tabla 2 a la
Tabla 20 y luego regresa a la Tabla 3). Aquí las tablas se identifican por
tipo de pavimento (R1, R2, R3, R4), nunca por ese número de tabla erróneo.

Otra particularidad (no un error a "corregir"): los valores de la tabla R2
y los de la tabla R3 son IDÉNTICOS entre sí en el documento fuente. Se
transcriben tal cual, sin alterarlos.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import NamedTuple

# ---------------------------------------------------------------------------
# 1. Tablas de vialidades (Tablas 1, 2, "20"/R3 y 3 de la norma)
# ---------------------------------------------------------------------------


class RequisitoVialidad(NamedTuple):
    """Requisitos de la norma para una clasificación de vialidad y pavimento.

    Los cuatro valores de DPEA corresponden, en orden, a los rangos de ancho
    de calle: < 9,0 m / > 9,0 y < 10,5 m / > 10,5 y < 12,0 m / > 12,0 m.
    """

    iluminancia_minima_promedio_lx: float
    uniformidad_maxima: float  # Eprom / Emin, máximo permitido
    dpea_ancho_menor_9: float  # W/m²
    dpea_ancho_9_a_10_5: float  # W/m²
    dpea_ancho_10_5_a_12: float  # W/m²
    dpea_ancho_mayor_12: float  # W/m²


# Nombres canónicos (en español) de las 7 clasificaciones de vialidad de la
# sección 5.1 de la norma.
VIALIDAD_AUTOPISTAS_CARRETERAS = "autopistas_y_carreteras"
VIALIDAD_ACCESO_CONTROLADO = "vias_de_acceso_controlado_y_vias_rapidas"
VIALIDAD_PRINCIPALES_EJES = "vias_principales_y_ejes_viales"
VIALIDAD_PRIMARIAS_COLECTORAS = "vias_primarias_y_colectoras"
VIALIDAD_SECUNDARIA_RESIDENCIAL_A = "vias_secundarias_residencial_tipo_a"
VIALIDAD_SECUNDARIA_RESIDENCIAL_B = "vias_secundarias_residencial_tipo_b"
VIALIDAD_SECUNDARIA_INDUSTRIAL_C = "vias_secundarias_industrial_tipo_c"

# Etiquetas para mostrar al usuario, en español, tal como las nombra la norma.
ETIQUETAS_VIALIDAD: dict[str, str] = {
    VIALIDAD_AUTOPISTAS_CARRETERAS: "Autopistas y carreteras",
    VIALIDAD_ACCESO_CONTROLADO: "Vías de acceso controlado y vías rápidas",
    VIALIDAD_PRINCIPALES_EJES: "Vías principales y ejes viales",
    VIALIDAD_PRIMARIAS_COLECTORAS: "Vías primarias y colectoras",
    VIALIDAD_SECUNDARIA_RESIDENCIAL_A: "Vías secundarias residencial Tipo A",
    VIALIDAD_SECUNDARIA_RESIDENCIAL_B: "Vías secundarias residencial Tipo B",
    VIALIDAD_SECUNDARIA_INDUSTRIAL_C: "Vías secundarias industrial Tipo C",
}

# Tipos de pavimento del Apéndice D.
PAVIMENTO_R1 = "R1"
PAVIMENTO_R2 = "R2"
PAVIMENTO_R3 = "R3"
PAVIMENTO_R4 = "R4"

# Tabla 1 de la norma: vialidades con pavimento tipo R1.
TABLA_R1: dict[str, RequisitoVialidad] = {
    VIALIDAD_AUTOPISTAS_CARRETERAS: RequisitoVialidad(4, 3, 0.32, 0.28, 0.26, 0.23),
    VIALIDAD_ACCESO_CONTROLADO: RequisitoVialidad(10, 3, 0.71, 0.66, 0.61, 0.56),
    VIALIDAD_PRINCIPALES_EJES: RequisitoVialidad(12, 3, 0.86, 0.81, 0.74, 0.69),
    VIALIDAD_PRIMARIAS_COLECTORAS: RequisitoVialidad(8, 4, 0.56, 0.52, 0.48, 0.44),
    VIALIDAD_SECUNDARIA_RESIDENCIAL_A: RequisitoVialidad(6, 6, 0.41, 0.38, 0.35, 0.31),
    VIALIDAD_SECUNDARIA_RESIDENCIAL_B: RequisitoVialidad(5, 6, 0.35, 0.33, 0.30, 0.28),
    VIALIDAD_SECUNDARIA_INDUSTRIAL_C: RequisitoVialidad(3, 6, 0.26, 0.23, 0.19, 0.17),
}

# Tabla 2 de la norma: vialidades con pavimento tipo R2.
TABLA_R2: dict[str, RequisitoVialidad] = {
    VIALIDAD_AUTOPISTAS_CARRETERAS: RequisitoVialidad(6, 3, 0.41, 0.38, 0.35, 0.31),
    VIALIDAD_ACCESO_CONTROLADO: RequisitoVialidad(14, 3, 1.01, 0.95, 0.86, 0.81),
    VIALIDAD_PRINCIPALES_EJES: RequisitoVialidad(17, 3, 1.17, 1.12, 1.03, 0.97),
    VIALIDAD_PRIMARIAS_COLECTORAS: RequisitoVialidad(12, 4, 0.86, 0.81, 0.74, 0.69),
    VIALIDAD_SECUNDARIA_RESIDENCIAL_A: RequisitoVialidad(9, 6, 0.64, 0.59, 0.54, 0.50),
    VIALIDAD_SECUNDARIA_RESIDENCIAL_B: RequisitoVialidad(7, 6, 0.49, 0.45, 0.42, 0.37),
    VIALIDAD_SECUNDARIA_INDUSTRIAL_C: RequisitoVialidad(4, 6, 0.32, 0.28, 0.26, 0.23),
}

# Tabla rotulada "Tabla 20" en el markdown fuente (errata: es la tabla de
# vialidades con pavimento tipo R3). Valores idénticos a los de R2 según el
# propio documento; no es un error de transcripción de este módulo.
TABLA_R3: dict[str, RequisitoVialidad] = {
    VIALIDAD_AUTOPISTAS_CARRETERAS: RequisitoVialidad(6, 3, 0.41, 0.38, 0.35, 0.31),
    VIALIDAD_ACCESO_CONTROLADO: RequisitoVialidad(14, 3, 1.01, 0.95, 0.86, 0.81),
    VIALIDAD_PRINCIPALES_EJES: RequisitoVialidad(17, 3, 1.17, 1.12, 1.03, 0.97),
    VIALIDAD_PRIMARIAS_COLECTORAS: RequisitoVialidad(12, 4, 0.86, 0.81, 0.74, 0.69),
    VIALIDAD_SECUNDARIA_RESIDENCIAL_A: RequisitoVialidad(9, 6, 0.64, 0.59, 0.54, 0.50),
    VIALIDAD_SECUNDARIA_RESIDENCIAL_B: RequisitoVialidad(7, 6, 0.49, 0.45, 0.42, 0.37),
    VIALIDAD_SECUNDARIA_INDUSTRIAL_C: RequisitoVialidad(4, 6, 0.32, 0.28, 0.26, 0.23),
}

# Tabla 3 de la norma: vialidades con pavimento tipo R4.
TABLA_R4: dict[str, RequisitoVialidad] = {
    VIALIDAD_AUTOPISTAS_CARRETERAS: RequisitoVialidad(5, 3, 0.35, 0.33, 0.30, 0.28),
    VIALIDAD_ACCESO_CONTROLADO: RequisitoVialidad(13, 3, 0.94, 0.87, 0.80, 0.75),
    VIALIDAD_PRINCIPALES_EJES: RequisitoVialidad(15, 3, 1.06, 1.00, 0.93, 0.87),
    VIALIDAD_PRIMARIAS_COLECTORAS: RequisitoVialidad(10, 4, 0.71, 0.66, 0.61, 0.56),
    VIALIDAD_SECUNDARIA_RESIDENCIAL_A: RequisitoVialidad(8, 6, 0.56, 0.52, 0.48, 0.44),
    VIALIDAD_SECUNDARIA_RESIDENCIAL_B: RequisitoVialidad(6, 6, 0.41, 0.38, 0.35, 0.31),
    VIALIDAD_SECUNDARIA_INDUSTRIAL_C: RequisitoVialidad(4, 6, 0.32, 0.28, 0.26, 0.23),
}

TABLAS_POR_PAVIMENTO: dict[str, dict[str, RequisitoVialidad]] = {
    PAVIMENTO_R1: TABLA_R1,
    PAVIMENTO_R2: TABLA_R2,
    PAVIMENTO_R3: TABLA_R3,
    PAVIMENTO_R4: TABLA_R4,
}

# Etiquetas de pavimento para mostrar al usuario.
ETIQUETAS_PAVIMENTO: dict[str, str] = {
    PAVIMENTO_R1: "R1",
    PAVIMENTO_R2: "R2",
    PAVIMENTO_R3: "R3",
    PAVIMENTO_R4: "R4",
}

# Apéndice D: coeficiente de luminancia media por tipo de pavimento.
# Solo se conserva como dato de referencia para el reporte; este módulo NO
# implementa el cálculo de luminancia (Tabla 4 usa luminancia, no iluminancia).
COEFICIENTE_LUMINANCIA_MEDIA: dict[str, float] = {
    PAVIMENTO_R1: 0.10,
    PAVIMENTO_R2: 0.07,
    PAVIMENTO_R3: 0.07,
    PAVIMENTO_R4: 0.08,
}


# ---------------------------------------------------------------------------
# 2. Selección de la columna de DPEA según el ancho de calle
# ---------------------------------------------------------------------------
#
# La norma define los rangos como:
#   < 9,0 / > 9,0 y < 10,5 / > 10,5 y < 12,0 / > 12,0
#
# Esto deja sin definir explícitamente qué pasa exactamente en los valores
# frontera (9,0, 10,5 y 12,0 m): la norma usa "<" y ">" estrictos, no "<=" ni
# ">=". Criterio adoptado aquí (documentado y aplicado de forma consistente):
# se asigna cada valor frontera a la columna del rango MÁS ANCHO (es decir,
# se trata "< 9,0" como "<= 9,0" para el límite superior siguiente, y así
# sucesivamente). Esto es el criterio conservador desde el punto de vista de
# la exigencia de la norma: ante la ambigüedad, un ancho exactamente en la
# frontera debe cumplir el DPEA más estricto (menor) del rango superior,
# nunca el más laxo del rango inferior.
def _indice_columna_dpea(ancho_m: float) -> int:
    """Determina el índice (0-3) de la columna de DPEA según el ancho de calle."""
    if ancho_m < 0:
        raise ValueError("El ancho de calle no puede ser negativo")
    if ancho_m < 9.0:
        return 0
    if ancho_m <= 10.5:
        return 1
    if ancho_m <= 12.0:
        return 2
    return 3


def dpea_maximo(requisito: RequisitoVialidad, ancho_m: float) -> float:
    """Devuelve el DPEA máximo [W/m²] aplicable según el ancho de calle."""
    columna = _indice_columna_dpea(ancho_m)
    valores = (
        requisito.dpea_ancho_menor_9,
        requisito.dpea_ancho_9_a_10_5,
        requisito.dpea_ancho_10_5_a_12,
        requisito.dpea_ancho_mayor_12,
    )
    return valores[columna]


# ---------------------------------------------------------------------------
# 3. Cálculo de DPEA (sección 8.1)
# ---------------------------------------------------------------------------


def dpea(watts_conectados: float, ancho_m: float, largo_m: float) -> float:
    """Calcula la Densidad de Potencia Eléctrica para Alumbrado [W/m²].

    DPEA = carga total conectada [W] / área total iluminada [m²]

    Según la sección 6.1/8.1 de la norma, el ancho de calle usado para el
    área NO debe incluir aceras ni camellones: se asume que `ancho_m` ya
    viene depurado de esas áreas (es responsabilidad de quien llama).
    """
    if ancho_m <= 0 or largo_m <= 0:
        raise ValueError("El ancho y el largo deben ser mayores que cero")
    area = ancho_m * largo_m
    return watts_conectados / area


# ---------------------------------------------------------------------------
# 4. Resolución tolerante de nombres
# ---------------------------------------------------------------------------


def _normalizar(texto: str) -> str:
    """Normaliza texto para comparaciones tolerantes: minúsculas, sin acentos."""
    txt = texto.strip().lower()
    reemplazos = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u", "ñ": "n",
    }
    for original, plano in reemplazos.items():
        txt = txt.replace(original, plano)
    return txt


_ALIAS_VIALIDAD: dict[str, str] = {
    # autopistas y carreteras
    "autopistas y carreteras": VIALIDAD_AUTOPISTAS_CARRETERAS,
    "autopistas": VIALIDAD_AUTOPISTAS_CARRETERAS,
    "carreteras": VIALIDAD_AUTOPISTAS_CARRETERAS,
    # acceso controlado
    "vias de acceso controlado y vias rapidas": VIALIDAD_ACCESO_CONTROLADO,
    "acceso controlado": VIALIDAD_ACCESO_CONTROLADO,
    "vias rapidas": VIALIDAD_ACCESO_CONTROLADO,
    # principales y ejes
    "vias principales y ejes viales": VIALIDAD_PRINCIPALES_EJES,
    "principales": VIALIDAD_PRINCIPALES_EJES,
    "ejes viales": VIALIDAD_PRINCIPALES_EJES,
    "vias principales": VIALIDAD_PRINCIPALES_EJES,
    # primarias y colectoras
    "vias primarias y colectoras": VIALIDAD_PRIMARIAS_COLECTORAS,
    "primarias": VIALIDAD_PRIMARIAS_COLECTORAS,
    "colectoras": VIALIDAD_PRIMARIAS_COLECTORAS,
    # secundaria residencial A
    "vias secundarias residencial tipo a": VIALIDAD_SECUNDARIA_RESIDENCIAL_A,
    "residencial tipo a": VIALIDAD_SECUNDARIA_RESIDENCIAL_A,
    "secundaria residencial a": VIALIDAD_SECUNDARIA_RESIDENCIAL_A,
    "tipo a": VIALIDAD_SECUNDARIA_RESIDENCIAL_A,
    "residencial a": VIALIDAD_SECUNDARIA_RESIDENCIAL_A,
    # secundaria residencial B
    "vias secundarias residencial tipo b": VIALIDAD_SECUNDARIA_RESIDENCIAL_B,
    "residencial tipo b": VIALIDAD_SECUNDARIA_RESIDENCIAL_B,
    "secundaria residencial b": VIALIDAD_SECUNDARIA_RESIDENCIAL_B,
    "tipo b": VIALIDAD_SECUNDARIA_RESIDENCIAL_B,
    "residencial b": VIALIDAD_SECUNDARIA_RESIDENCIAL_B,
    # secundaria industrial C
    "vias secundarias industrial tipo c": VIALIDAD_SECUNDARIA_INDUSTRIAL_C,
    "industrial tipo c": VIALIDAD_SECUNDARIA_INDUSTRIAL_C,
    "secundaria industrial c": VIALIDAD_SECUNDARIA_INDUSTRIAL_C,
    "tipo c": VIALIDAD_SECUNDARIA_INDUSTRIAL_C,
    "industrial c": VIALIDAD_SECUNDARIA_INDUSTRIAL_C,
}


def resolver_vialidad(nombre: str) -> str:
    """Resuelve, de forma tolerante, el nombre canónico de una clasificación
    de vialidad a partir de texto escrito por el usuario.

    Acepta el nombre canónico (con o sin guiones bajos), la etiqueta exacta
    de la norma, o alias comunes y abreviados (p. ej. "residencial tipo a").
    Lanza ValueError si no se reconoce.
    """
    clave = _normalizar(nombre).replace("_", " ")
    # ¿Coincide ya con un nombre canónico?
    for canonico in ETIQUETAS_VIALIDAD:
        if clave == canonico.replace("_", " "):
            return canonico
    # ¿Coincide con una etiqueta oficial?
    for canonico, etiqueta in ETIQUETAS_VIALIDAD.items():
        if clave == _normalizar(etiqueta):
            return canonico
    # ¿Coincide con un alias?
    if clave in _ALIAS_VIALIDAD:
        return _ALIAS_VIALIDAD[clave]
    raise ValueError(f"Clasificación de vialidad no reconocida: {nombre!r}")


_ALIAS_PAVIMENTO: dict[str, str] = {
    "r1": PAVIMENTO_R1,
    "r2": PAVIMENTO_R2,
    "r3": PAVIMENTO_R3,
    "r4": PAVIMENTO_R4,
}


def resolver_pavimento(nombre: str) -> str:
    """Resuelve, de forma tolerante, el tipo de pavimento (R1-R4) a partir de
    texto escrito por el usuario (p. ej. "r2", "R 2", "pavimento R3").
    """
    clave = _normalizar(nombre).replace(" ", "").replace("-", "").replace("_", "")
    clave = clave.replace("pavimento", "")
    if clave in _ALIAS_PAVIMENTO:
        return _ALIAS_PAVIMENTO[clave]
    raise ValueError(f"Tipo de pavimento no reconocido: {nombre!r}")


# ---------------------------------------------------------------------------
# 5. Veredicto de cumplimiento
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CriterioEvaluado:
    """Resultado de evaluar un único criterio de la norma."""

    nombre: str
    valor_obtenido: float
    valor_limite: float
    cumple: bool
    comparacion: str  # descripción de cómo se comparó, p. ej. ">= mínimo"


@dataclass(frozen=True)
class ResultadoEvaluacion:
    """Veredicto completo de cumplimiento de la NOM-013-ENER-2013 para
    vialidades, con un criterio independiente por cada exigencia de la norma.
    """

    vialidad: str
    pavimento: str
    ancho_m: float
    nivel_iluminacion: CriterioEvaluado
    uniformidad: CriterioEvaluado
    dpea: CriterioEvaluado

    @property
    def cumple(self) -> bool:
        """Veredicto global: cumple solo si los tres criterios cumplen."""
        return (
            self.nivel_iluminacion.cumple
            and self.uniformidad.cumple
            and self.dpea.cumple
        )


def evaluar(
    eprom: float,
    emin: float,
    dpea_calculado: float,
    vialidad: str,
    pavimento: str,
    ancho_m: float,
) -> ResultadoEvaluacion:
    """Evalúa el cumplimiento de la NOM-013-ENER-2013 para una vialidad.

    Parámetros
    ----------
    eprom: iluminancia promedio medida/calculada [lx].
    emin: iluminancia mínima medida/calculada [lx] (de los 9 puntos, Apéndice C).
    dpea_calculado: DPEA del sistema [W/m²] (ver `dpea()`).
    vialidad: clasificación de vialidad (nombre tolerante, ver `resolver_vialidad`).
    pavimento: tipo de pavimento R1-R4 (nombre tolerante, ver `resolver_pavimento`).
    ancho_m: ancho de calle [m] (sin aceras ni camellones), para elegir la
        columna de DPEA aplicable.

    Devuelve un `ResultadoEvaluacion` con los tres criterios evaluados de
    forma independiente (nivel de iluminación, uniformidad, DPEA) y el
    veredicto global. Ningún criterio se colapsa en un solo booleano: cada
    uno reporta el valor obtenido, el límite normativo y si cumple.

    Sentido de cada comparación:
    - Eprom es un valor MÍNIMO: cumple si eprom >= iluminancia_minima_promedio.
    - La uniformidad Eprom/Emin es un valor MÁXIMO: cumple si la razón
      calculada es <= la razón máxima permitida.
    - El DPEA es un valor MÁXIMO: cumple si dpea_calculado <= dpea_maximo.
    """
    vialidad_canonica = resolver_vialidad(vialidad)
    pavimento_canonico = resolver_pavimento(pavimento)
    tabla = TABLAS_POR_PAVIMENTO[pavimento_canonico]
    requisito = tabla[vialidad_canonica]

    # Criterio 1: nivel de iluminación (mínimo).
    criterio_iluminacion = CriterioEvaluado(
        nombre="Nivel de iluminación (Eprom)",
        valor_obtenido=eprom,
        valor_limite=requisito.iluminancia_minima_promedio_lx,
        cumple=eprom >= requisito.iluminancia_minima_promedio_lx,
        comparacion=">= mínimo",
    )

    # Criterio 2: uniformidad (máximo). Emin == 0 se maneja sin dividir entre
    # cero: una iluminancia mínima de 0 lx es, en la práctica, el peor caso
    # posible de uniformidad (relación infinita) y por tanto no cumple.
    if emin == 0:
        uniformidad_obtenida = float("inf")
    else:
        uniformidad_obtenida = eprom / emin
    criterio_uniformidad = CriterioEvaluado(
        nombre="Uniformidad (Eprom/Emin)",
        valor_obtenido=uniformidad_obtenida,
        valor_limite=requisito.uniformidad_maxima,
        cumple=uniformidad_obtenida <= requisito.uniformidad_maxima,
        comparacion="<= máximo",
    )

    # Criterio 3: DPEA (máximo).
    limite_dpea = dpea_maximo(requisito, ancho_m)
    criterio_dpea = CriterioEvaluado(
        nombre="DPEA",
        valor_obtenido=dpea_calculado,
        valor_limite=limite_dpea,
        cumple=dpea_calculado <= limite_dpea,
        comparacion="<= máximo",
    )

    return ResultadoEvaluacion(
        vialidad=vialidad_canonica,
        pavimento=pavimento_canonico,
        ancho_m=ancho_m,
        nivel_iluminacion=criterio_iluminacion,
        uniformidad=criterio_uniformidad,
        dpea=criterio_dpea,
    )
