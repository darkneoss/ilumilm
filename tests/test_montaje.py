"""Montaje del luminario: dos luminarios por poste y angulo de separacion.

Este camino NO tiene corrida de referencia. Las 24 corridas del Excel son
todas de un luminario por poste con angulo 0, asi que aqui no se puede
comparar contra nada medido; lo que se prueba son invariantes que se deducen de
la geometria y que fallarian si el port estuviera mal. Ver el aviso de
`reference/VALIDACION.md`.

La red que si es medicion esta en test_regresion_sead.py: con el montaje por
omision los 24 casos siguen cuadrando al 0.008 %, lo que prueba que el camino
nuevo es *neutro* cuando no se usa.
"""
from __future__ import annotations

import math

import pytest

from engine import calc, ies
from engine.geometry import Montaje, Vialidad, posiciones_luminarios

CATALOGO = "catalogo/V2100UN2M50.ies"


def _via(disposicion="unilateral"):
    return Vialidad(num_carriles=2, ancho_carril=3.5, camellon=0.0,
                    disposicion=disposicion, altura_montaje=8.0,
                    interpostal=40.0, retranqueo=0.2, largo_brazo=1.8)


# ---------------------------------------------------------------------------
# Lo que el montaje por omision no debe cambiar
# ---------------------------------------------------------------------------

def test_el_montaje_por_omision_es_el_corriente():
    m = Montaje()
    assert m.neutro
    assert (m.inclinacion, m.luminarios_por_poste, m.angulo_separacion) == (0.0, 1, 0.0)


@pytest.mark.parametrize("disposicion", ["unilateral", "tresbolillo",
                                         "central doble", "bilateral opuesta"])
def test_pasar_el_montaje_neutro_es_igual_que_no_pasar_nada(disposicion):
    v = _via(disposicion)
    assert posiciones_luminarios(v) == posiciones_luminarios(v, Montaje())


# ---------------------------------------------------------------------------
# Dos luminarios por poste
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("disposicion,n_simple", [
    ("unilateral", 5), ("tresbolillo", 10),
    ("central doble", 10), ("bilateral opuesta", 10),
])
def test_dos_por_poste_duplica_el_conteo(disposicion, n_simple):
    v = _via(disposicion)
    assert len(posiciones_luminarios(v)) == n_simple
    doble = posiciones_luminarios(v, Montaje(luminarios_por_poste=2))
    assert len(doble) == 2 * n_simple


def test_sin_angulo_de_separacion_los_dos_brazos_quedan_superpuestos():
    """Dos luminarios en un poste y 0 grados de separacion es fisicamente
    absurdo --uno encima del otro-- pero es el caso limite que fija la escala:
    la iluminancia tiene que salir EXACTAMENTE al doble.

    Sirve de red porque atrapa cualquier error de factor: si el reparto de
    brazos o el conteo estuvieran mal, este numero no seria 2.000000.
    """
    v = _via()
    foto = ies.lee(CATALOGO)
    simple = calc.calcula(v, foto, 0.765)
    doble = calc.calcula(v, foto, 0.765, montaje=Montaje(luminarios_por_poste=2))
    assert doble.promedio == pytest.approx(2 * simple.promedio, rel=1e-12)
    assert doble.minimo == pytest.approx(2 * simple.minimo, rel=1e-12)
    assert doble.uniformidad == pytest.approx(simple.uniformidad, rel=1e-12)


def test_el_angulo_de_separacion_abre_los_brazos_en_x_y_los_acorta_en_y():
    """`adjX_arm = ArmLength * sin(sep/2)`, `adjY_arm = ArmLength * cos(sep/2)`.

    El brazo no se alarga: se gira. Lo que gana en X lo pierde en Y, y la suma
    de cuadrados sigue siendo el largo del brazo.
    """
    v = _via()
    sep = 60.0
    ps = posiciones_luminarios(v, Montaje(luminarios_por_poste=2,
                                          angulo_separacion=sep))
    dx = v.largo_brazo * math.sin(math.radians(sep) / 2.0)
    dy = v.largo_brazo * math.cos(math.radians(sep) / 2.0)

    a, b = ps[0], ps[1]                      # los dos brazos del primer poste
    assert a.x == pytest.approx(-dx)         # el poste esta en x = 0
    assert b.x == pytest.approx(+dx)
    assert a.y == pytest.approx(dy - v.retranqueo)
    assert b.y == pytest.approx(dy - v.retranqueo)
    assert math.hypot(dx, dy) == pytest.approx(v.largo_brazo)
    # cada brazo gira hacia su lado
    assert (a.giro_z, b.giro_z) == (-sep, sep)


def test_el_giro_del_brazo_mueve_el_azimut_y_no_el_gamma():
    """Un punto perpendicular al luminario esta en C = 0 sin giro. Con el brazo
    girado g grados, ese mismo punto pasa a estar en C = |g| de la fotometria,
    y el gamma no cambia: girar sobre la vertical no sube ni baja el rayo.
    """
    for giro in (-30.0, -10.0, 15.0, 45.0):
        sin_giro = calc._angulos(0.0, 0.0, 1, 0.0, 6.0, 8.0, calc.MODO_CORRECTO)
        con_giro = calc._angulos(0.0, 0.0, 1, 0.0, 6.0, 8.0, calc.MODO_CORRECTO,
                                 0.0, giro)
        assert sin_giro[2] == pytest.approx(0.0, abs=1e-12)
        assert con_giro[2] == pytest.approx(abs(giro), abs=1e-9)
        assert con_giro[0] == pytest.approx(sin_giro[0], abs=1e-12)   # gamma geometrico
        assert con_giro[1] == pytest.approx(sin_giro[1], abs=1e-9)    # gamma de la tabla


# ---------------------------------------------------------------------------
# Lo que no se acepta, y por que
# ---------------------------------------------------------------------------

def test_mas_de_dos_por_poste_se_rechaza():
    """El reparto del original supone dos brazos simetricos (`sin(sep/2)`); su
    propio comentario admite que con tres no funcionaria. Mejor un error que un
    numero que parece bueno."""
    with pytest.raises(ValueError, match="dos brazos simetricos"):
        Montaje(luminarios_por_poste=3)


def test_el_angulo_de_separacion_sin_dos_luminarios_se_rechaza():
    """Separar un solo brazo no significa nada, y aceptarlo en silencio dejaria
    al usuario creyendo que giro algo."""
    with pytest.raises(ValueError, match="mas de un"):
        Montaje(angulo_separacion=30.0)


@pytest.mark.parametrize("kwargs", [
    {"luminarios_por_poste": 0},
    {"inclinacion": 90.0},
    {"inclinacion": -120.0},
    {"luminarios_por_poste": 2, "angulo_separacion": 200.0},
])
def test_valores_fuera_de_rango(kwargs):
    with pytest.raises(ValueError):
        Montaje(**kwargs)
