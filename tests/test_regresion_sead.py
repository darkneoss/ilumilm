"""
Regresion contra los 7 estudios reales del SEAD Street Lighting Tool
(assets/estudios/*.xlsx), extraidos previamente a reference/casos_sead.json
por tools/extraer_casos_sead.py.

Un caso = un estudio x un luminario evaluado en el (14 en total). Se corre el
motor en MODO_CORRECTO (el que coincide con el Excel, ver engine/calc.py) y
se compara contra el promedio/minimo/maximo/uniformidad que dejo cacheado el
Excel.

Tolerancias (relativas, en %):
    promedio    0.5 %
    maximo      0.5 %
    minimo      1.5 %   (mas sensible a la interpolacion de la fotometria)
    uniformidad se deriva de promedio/minimo: no se le pone tolerancia propia
    explicita en el enunciado, pero se reporta para diagnostico.

Si un caso no pasa, NO se relaja la tolerancia: se deja fallando.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from engine import calc
from engine import ies
from engine.geometry import Vialidad

RAIZ = Path(__file__).resolve().parent.parent
JSON_CASOS = RAIZ / "reference" / "casos_sead.json"
DIR_ASSETS = RAIZ / "assets"
DIR_CATALOGO = RAIZ / "catalogo"


def _busca_ies(nombre):
    """catalogo/ manda: es la copia versionada. assets/ es solo respaldo local."""
    for d in (DIR_CATALOGO, DIR_ASSETS):
        ruta = d / nombre
        if ruta.exists():
            return ruta
    return None

TOL_PROMEDIO_PCT = 0.5
TOL_MAXIMO_PCT = 0.5
TOL_MINIMO_PCT = 1.5

# BUG REPORTADO (no se toca engine/geometry.py, ver reference/VALIDACION.md):
# `normaliza_disposicion` NO reconoce los textos exactos que trae la columna
# "Posicion del poste" de estos .xlsx reales:
#   - "De un solo lado"        (con el prefijo "De ", los alias solo tienen
#                                 "un solo lado")
#   - "Montado en el camellón" (con acento y espacio final; los alias solo
#                                 tienen "camellon" sin acento y "median
#                                 mounted", no la frase completa con tilde)
# Se traducen aqui, en el test, a un alias que `normaliza_disposicion` SI
# acepta, para poder ejecutar la regresion sin modificar el motor.
_MAPEO_DISPOSICION_EXCEL = {
    "de un solo lado": "unilateral",
    "montado en el camellón": "central doble",
    "montado en el camellon": "central doble",
}


def _disposicion_aceptada_por_el_motor(texto_excel: str) -> str:
    clave = texto_excel.strip().lower()
    return _MAPEO_DISPOSICION_EXCEL.get(clave, texto_excel)


def _cargar_casos():
    with JSON_CASOS.open(encoding="utf-8") as f:
        datos = json.load(f)

    casos = []
    ids = []
    for estudio in datos["estudios"]:
        nombre_estudio = estudio.get("estudio", estudio.get("archivo"))
        for caso in estudio.get("casos", []):
            casos.append((nombre_estudio, caso))
            ids.append("{}::{}".format(nombre_estudio, caso["luminario"]["archivo_ies"]))
    return casos, ids


_CASOS, _IDS = _cargar_casos()


def _error_pct(esperado: float, obtenido: float) -> float:
    if esperado == 0:
        return float("inf") if obtenido != 0 else 0.0
    return 100.0 * (obtenido - esperado) / esperado


@pytest.mark.parametrize("nombre_estudio,caso", _CASOS, ids=_IDS)
def test_regresion_sead(nombre_estudio, caso):
    geo = caso["geometria"]
    esperado = caso["esperado"]

    ruta_ies = _busca_ies(caso["luminario"]["archivo_ies"])
    if ruta_ies is None:
        pytest.skip(
            "falta la fotometria {}. Los .ies no se versionan; copialos a "
            "assets/ o catalogo/ para correr la regresion.".format(
                caso["luminario"]["archivo_ies"])
        )

    foto = ies.lee(ruta_ies)

    v = Vialidad(
        num_carriles=geo["num_carriles"],
        ancho_carril=geo["ancho_carril"],
        camellon=geo["camellon"],
        disposicion=_disposicion_aceptada_por_el_motor(geo["posicion_poste"]),
        altura_montaje=geo["altura_montaje"],
        interpostal=geo["interpostal"],
        retranqueo=geo["retranqueo"],
        largo_brazo=geo["largo_brazo"],
    )

    malla = calc.calcula(v, foto, geo["llf"], modo=calc.MODO_CORRECTO)

    err_prom = _error_pct(esperado["promedio"], malla.promedio)
    err_min = _error_pct(esperado["minimo"], malla.minimo)
    err_max = _error_pct(esperado["maximo"], malla.maximo)

    detalle = (
        "\nEstudio: {}\nLuminario: {}\n"
        "Promedio: excel={:.4f} motor={:.4f} error={:.3f}%\n"
        "Minimo:   excel={:.4f} motor={:.4f} error={:.3f}%\n"
        "Maximo:   excel={:.4f} motor={:.4f} error={:.3f}%\n"
    ).format(
        nombre_estudio, caso["luminario"]["archivo_ies"],
        esperado["promedio"], malla.promedio, err_prom,
        esperado["minimo"], malla.minimo, err_min,
        esperado["maximo"], malla.maximo, err_max,
    )

    assert abs(err_prom) <= TOL_PROMEDIO_PCT, detalle
    assert abs(err_min) <= TOL_MINIMO_PCT, detalle
    assert abs(err_max) <= TOL_MAXIMO_PCT, detalle
