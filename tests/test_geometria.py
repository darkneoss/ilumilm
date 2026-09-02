"""Colocacion de luminarios: cuantos hay, donde y hacia donde miran.

Lo que fija este archivo es la decision de que **los dos luminarios por poste
existen solo en central doble**, donde los pone la disposicion misma. El
`selectedFixturesPerPole` del original --un multiplicador que aplicaria a
cualquier disposicion-- no se implementa; ver el docstring de
`posiciones_luminarios` y reference/sead_vba_1.8.1/README.md.

La razon de probarlo aqui y no confiar en la regresion: la regresion mide la
iluminancia, y la iluminancia es una suma. Dos colocaciones distintas pueden
dar el mismo total --de hecho el port tenia central doble espejeado y cuadraba
igual--, asi que hace falta mirar las posiciones de frente.
"""
from __future__ import annotations

import pytest

from engine.cli import _luminarios_por_tramo
from engine.geometry import Vialidad, posiciones_luminarios

S = 40.0


def _via(disposicion, carriles=4, camellon=0.0):
    return Vialidad(num_carriles=carriles, ancho_carril=3.5, camellon=camellon,
                    disposicion=disposicion, altura_montaje=8.0, interpostal=S,
                    retranqueo=0.2, largo_brazo=1.8)


@pytest.mark.parametrize("disposicion,esperados", [
    ("unilateral", 5),          # un poste por fase, 5 fases
    ("tresbolillo", 10),        # dos postes por fase, alternados
    ("bilateral opuesta", 10),  # dos postes por fase, enfrentados
    ("central doble", 10),      # UN poste por fase con DOS brazos
])
def test_cuantos_luminarios_hay(disposicion, esperados):
    assert len(posiciones_luminarios(_via(disposicion))) == esperados


def test_en_central_doble_los_dos_brazos_comparten_poste_y_miran_al_revés():
    """Es el unico caso de dos luminarios por poste, y no hay que declararlo:
    lo pone la disposicion. Los dos brazos salen del mismo eje --misma X-- y
    quedan simetricos respecto al centro de la seccion, cada uno alumbrando su
    sentido."""
    v = _via("central doble", carriles=4, camellon=2.0)
    a, b = posiciones_luminarios(v)[:2]

    assert a.x == b.x == 0.0                      # el mismo poste
    assert {a.orientacion, b.orientacion} == {1, -1}   # uno hacia cada sentido

    centro = v.ancho_calzada / 2.0
    assert (a.y + b.y) / 2.0 == pytest.approx(centro)
    assert abs(a.y - centro) == pytest.approx(v.largo_brazo)
    assert abs(b.y - centro) == pytest.approx(v.largo_brazo)


def test_el_retranqueo_no_mueve_los_brazos_del_camellón():
    """Un poste central no tiene retranqueo: los dos brazos salen del eje. El
    original lo suma y lo resta igual (su propio comentario admite que sobra);
    replicarlo daba 8 % de error en Emin contra los estudios de referencia."""
    sin_r = posiciones_luminarios(
        Vialidad(4, 3.5, 2.0, "central doble", 8.0, S, 0.0, 1.8))
    con_r = posiciones_luminarios(
        Vialidad(4, 3.5, 2.0, "central doble", 8.0, S, 3.7, 1.8))
    assert sin_r == con_r


@pytest.mark.parametrize("disposicion,esperado", [
    ("unilateral", 1),
    ("tresbolillo", 2),
    ("bilateral opuesta", 2),
    ("central doble", 2),
])
def test_luminarios_por_tramo_para_el_dpea(disposicion, esperado):
    """El DPEA se calcula sobre la carga conectada del tramo, asi que este
    numero es criterio de cumplimiento y no un dato informativo. En central
    doble son dos SIN declarar nada, que es justo la confusion que hay que
    evitar: pedir "dos luminarios por poste" ahi lo dejaria en cuatro."""
    v = _via(disposicion)
    assert _luminarios_por_tramo(v.disposicion) == esperado
    # y coincide con lo que de verdad se coloco por fase de poste
    assert len(posiciones_luminarios(v)) == esperado * 5


@pytest.mark.parametrize("disposicion", ["unilateral", "tresbolillo",
                                         "bilateral opuesta"])
def test_los_postes_de_orilla_alumbran_hacia_el_centro(disposicion):
    """`orientacion` decide con que mitad de la fotometria se ilumina. Si un
    luminario mira al lado contrario, la calzada recibe la distribucion de la
    acera: fue el peor error de toda la validacion, +55 % en Emax."""
    v = _via(disposicion)
    centro = v.ancho_calzada / 2.0
    for lum in posiciones_luminarios(v):
        hacia_el_centro = 1 if lum.y < centro else -1
        assert lum.orientacion == hacia_el_centro


def test_los_brazos_del_camellón_alumbran_ALEJÁNDOSE_del_eje():
    """Y aqui la regla es la contraria, que es facil de leer al reves.

    Un poste de orilla apunta hacia el centro de la calzada. Uno del camellon
    ya esta en el centro: cada brazo alumbra su propio sentido de circulacion,
    o sea hacia afuera. Invertir esto haria que los dos brazos se peleen el
    mismo carril y dejen los extremos a oscuras.
    """
    v = _via("central doble", camellon=2.0)
    centro = v.ancho_calzada / 2.0
    for lum in posiciones_luminarios(v):
        hacia_afuera = -1 if lum.y < centro else 1
        assert lum.orientacion == hacia_afuera


# ---------------------------------------------------------------------------
# Lo que la entrada no acepta, y por que importa que avise
# ---------------------------------------------------------------------------

def test_declarar_luminarios_por_poste_es_un_error_y_no_un_campo_ignorado():
    """Aceptar en silencio un campo que no se usa es el defecto exacto de la
    herramienta que se sustituye: usa la inclinacion en el calculo y no la
    reporta, y eso costo una validacion entera. Un parametro que el usuario
    escribe y el motor descarta sin avisar es peor que un error."""
    import json
    import tempfile
    from pathlib import Path

    from engine.cli import ErrorEstudio, ejecuta

    entrada = {
        "vialidad": {"num_carriles": 2, "ancho_carril": 3.5, "camellon": 0.0,
                     "disposicion": "unilateral", "altura_montaje": 8.0,
                     "interpostal": 35.0, "retranqueo": 0.2, "largo_brazo": 1.8},
        "nom": {"clasificacion_vialidad": "vias_secundarias_residencial_tipo_a",
                "pavimento": "R2"},
        "luminarios": [{"archivo": "V1050UN2M50.ies", "luminarios_por_poste": 2}],
    }
    with tempfile.TemporaryDirectory() as d:
        ruta = Path(d) / "entrada.json"
        ruta.write_text(json.dumps(entrada), encoding="utf-8")
        with pytest.raises(ErrorEstudio, match="central doble"):
            ejecuta(ruta)
