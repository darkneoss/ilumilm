"""Pruebas para engine/nom.py (NOM-013-ENER-2013)."""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine import nom


def test_r2_secundaria_residencial_a_tutorial():
    """Caso del tutorial: R2 + Vías secundarias residencial Tipo A."""
    requisito = nom.TABLA_R2[nom.VIALIDAD_SECUNDARIA_RESIDENCIAL_A]
    assert requisito.iluminancia_minima_promedio_lx == 9
    assert requisito.uniformidad_maxima == 6
    assert nom.dpea_maximo(requisito, 7) == 0.64


def test_r2_principales_ejes():
    """R2 + Vías principales y ejes viales, ancho 15 m."""
    requisito = nom.TABLA_R2[nom.VIALIDAD_PRINCIPALES_EJES]
    assert requisito.iluminancia_minima_promedio_lx == 17
    assert requisito.uniformidad_maxima == 3
    assert nom.dpea_maximo(requisito, 15) == 0.97


def test_falla_solo_por_uniformidad():
    """Ilumina bien y DPEA bien, pero la uniformidad excede el máximo."""
    resultado = nom.evaluar(
        eprom=9.0,
        emin=1.0,  # 9/1 = 9, excede el máximo de 6 para residencial A
        dpea_calculado=0.5,
        vialidad="residencial tipo a",
        pavimento="R2",
        ancho_m=7,
    )
    assert resultado.nivel_iluminacion.cumple is True
    assert resultado.uniformidad.cumple is False
    assert resultado.dpea.cumple is True
    assert resultado.cumple is False


def test_falla_solo_por_dpea():
    """Ilumina bien y uniformidad bien, pero el DPEA excede el máximo."""
    resultado = nom.evaluar(
        eprom=9.0,
        emin=2.0,  # 9/2 = 4.5, cumple (<=6)
        dpea_calculado=0.90,  # excede 0.64
        vialidad="residencial tipo a",
        pavimento="R2",
        ancho_m=7,
    )
    assert resultado.nivel_iluminacion.cumple is True
    assert resultado.uniformidad.cumple is True
    assert resultado.dpea.cumple is False
    assert resultado.cumple is False


def test_falla_solo_por_iluminacion():
    """DPEA y uniformidad cumplen, pero Eprom queda por debajo del mínimo."""
    resultado = nom.evaluar(
        eprom=5.0,  # menor que el mínimo de 9
        emin=2.0,  # 5/2 = 2.5, cumple
        dpea_calculado=0.5,
        vialidad="residencial tipo a",
        pavimento="R2",
        ancho_m=7,
    )
    assert resultado.nivel_iluminacion.cumple is False
    assert resultado.uniformidad.cumple is True
    assert resultado.dpea.cumple is True
    assert resultado.cumple is False


def test_cumple_todo():
    resultado = nom.evaluar(
        eprom=9.0,
        emin=2.0,
        dpea_calculado=0.5,
        vialidad="residencial tipo a",
        pavimento="R2",
        ancho_m=7,
    )
    assert resultado.cumple is True


def test_emin_cero_no_revienta():
    resultado = nom.evaluar(
        eprom=9.0,
        emin=0.0,
        dpea_calculado=0.5,
        vialidad="residencial tipo a",
        pavimento="R2",
        ancho_m=7,
    )
    assert resultado.uniformidad.cumple is False
    assert resultado.uniformidad.valor_obtenido == float("inf")


def test_columnas_dpea_los_cuatro_rangos():
    """Los cuatro rangos de ancho seleccionan la columna correcta de DPEA."""
    requisito = nom.TABLA_R1[nom.VIALIDAD_AUTOPISTAS_CARRETERAS]
    # < 9.0
    assert nom.dpea_maximo(requisito, 8.0) == requisito.dpea_ancho_menor_9
    # > 9.0 y <= 10.5 (criterio: frontera 9.0 exacta pertenece a esta columna)
    assert nom.dpea_maximo(requisito, 9.0) == requisito.dpea_ancho_9_a_10_5
    assert nom.dpea_maximo(requisito, 10.0) == requisito.dpea_ancho_9_a_10_5
    assert nom.dpea_maximo(requisito, 10.5) == requisito.dpea_ancho_9_a_10_5
    # > 10.5 y <= 12.0
    assert nom.dpea_maximo(requisito, 11.0) == requisito.dpea_ancho_10_5_a_12
    assert nom.dpea_maximo(requisito, 12.0) == requisito.dpea_ancho_10_5_a_12
    # > 12.0
    assert nom.dpea_maximo(requisito, 12.5) == requisito.dpea_ancho_mayor_12


def test_funcion_dpea():
    """dpea(watts, ancho, largo) -> W/m², sin incluir aceras/camellones (responsabilidad del llamador)."""
    resultado = nom.dpea(watts_conectados=100.0, ancho_m=7.0, largo_m=10.0)
    assert abs(resultado - (100.0 / 70.0)) < 1e-9


def test_resolver_vialidad_tolerante():
    assert nom.resolver_vialidad("residencial tipo a") == nom.VIALIDAD_SECUNDARIA_RESIDENCIAL_A
    assert nom.resolver_vialidad("Tipo A") == nom.VIALIDAD_SECUNDARIA_RESIDENCIAL_A
    assert nom.resolver_vialidad("Vías secundarias residencial Tipo A") == nom.VIALIDAD_SECUNDARIA_RESIDENCIAL_A


def test_resolver_pavimento_tolerante():
    assert nom.resolver_pavimento("r2") == nom.PAVIMENTO_R2
    assert nom.resolver_pavimento("R 2") == nom.PAVIMENTO_R2
    assert nom.resolver_pavimento("pavimento R3") == nom.PAVIMENTO_R3


def test_tablas_r2_y_r3_identicas():
    """No es un bug: la norma da valores idénticos para R2 y R3."""
    assert nom.TABLA_R2 == nom.TABLA_R3


if __name__ == "__main__":
    import traceback

    pruebas = [obj for nombre, obj in list(globals().items()) if nombre.startswith("test_")]
    fallas = 0
    for prueba in pruebas:
        try:
            prueba()
            print(f"OK   {prueba.__name__}")
        except AssertionError:
            fallas += 1
            print(f"FAIL {prueba.__name__}")
            traceback.print_exc()
    print(f"\n{len(pruebas) - fallas}/{len(pruebas)} pruebas pasaron")
    if fallas:
        sys.exit(1)
