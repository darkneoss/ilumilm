"""La inclinacion del luminario: los dos gammas y el signo.

La regresion contra los estudios reales (test_regresion_sead.py, estudio-09)
ya prueba que el resultado numerico coincide con el Excel a 5 y 15 grados.
Lo que prueba este archivo es *por que* coincide, con casos donde el angulo
correcto se sabe de antemano sin correr nada. Si alguien "limpia" `_angulos`
cableando el cos^3 al gamma inclinado --el error que la especificacion del VBA
daba por hecho durante toda la validacion anterior-- la regresion falla con un
porcentaje y estas pruebas fallan diciendo cual es el bug.
"""
from __future__ import annotations

import math

import pytest

from engine.calc import MODO_CORRECTO, _angulos

H = 10.0


def _en(dx: float, dy: float, inclinacion: float):
    """Angulos de un punto a (dx, dy) del pie del luminario."""
    return _angulos(0.0, 0.0, 1, dx, dy, H, MODO_CORRECTO, inclinacion)


@pytest.mark.parametrize("dx,dy", [(0.0, 0.0), (3.0, 7.0), (-12.0, 1.5), (20.0, -0.5)])
def test_sin_inclinacion_los_dos_gammas_coinciden(dx, dy):
    """Con inclinacion 0 no hay dos angulos que distinguir."""
    gamma_geo, gamma, _ = _en(dx, dy, 0.0)
    assert gamma_geo == pytest.approx(gamma, abs=1e-12)


def test_el_punto_al_pie_del_poste_queda_a_gamma_igual_a_la_inclinacion():
    """Justo debajo del luminario: el rayo es vertical, el eje optico no.

    gamma geometrico = 0 (el rayo cae a plomo, asi que el cos^3 vale 1), pero
    para la fotometria ese punto esta a `inclinacion` grados del eje del
    luminario y del lado de la ACERA (C = 180), porque la inclinacion positiva
    apunta hacia la calzada. Ese C es el que fija el signo: con el criterio
    contrario saldria C = 0 y el reporte iluminaria la calzada con la
    distribucion de la banqueta.
    """
    for t in (5.0, 15.0, 30.0):
        gamma_geo, gamma, c = _en(0.0, 0.0, t)
        assert gamma_geo == pytest.approx(0.0, abs=1e-12)
        assert gamma == pytest.approx(t, abs=1e-9)
        assert c == pytest.approx(180.0, abs=1e-9)


def test_el_punto_al_que_apunta_el_eje_optico_queda_a_gamma_cero():
    """El caso espejo del anterior, y el que fija la magnitud.

    Un luminario a H de altura inclinado t grados apunta su eje al punto que
    esta a H*tan(t) hacia la calzada. Para la fotometria ese punto esta en el
    nadir (gamma = 0, la intensidad maxima del eje), mientras que el rayo llega
    al pavimento a t grados de la vertical.
    """
    for t in (5.0, 15.0, 30.0):
        dy = H * math.tan(math.radians(t))
        gamma_geo, gamma, _ = _en(0.0, dy, t)
        assert gamma == pytest.approx(0.0, abs=1e-9)
        assert gamma_geo == pytest.approx(t, abs=1e-9)


def test_la_inclinacion_negativa_es_el_espejo_de_la_positiva():
    """-t sobre un punto de la acera == +t sobre su reflejo en la calzada.

    Sirve de red para el signo: si la rotacion se implementara con el seno
    invertido, esta simetria se rompe. Los dos gammas son identicos; el azimut
    sale espejeado (C -> 180 - C) porque espejear `dy` cambia de lado la
    calzada, que es el lado desde el que se mide C.
    """
    for t in (5.0, 15.0):
        for dx, dy in ((0.0, 4.0), (8.0, 12.0), (-3.0, 0.5)):
            a = _en(dx, dy, t)
            b = _en(dx, -dy, -t)
            assert a[0] == pytest.approx(b[0], abs=1e-9)          # gamma geometrico
            assert a[1] == pytest.approx(b[1], abs=1e-9)          # gamma de la tabla
            assert a[2] == pytest.approx(180.0 - b[2], abs=1e-9)  # azimut espejeado


def test_un_punto_detras_del_plano_del_luminario_pasa_de_90_grados():
    """Con inclinacion grande, lo que queda atras del plano optico va a
    gamma > 90: es la rama `If HPrime < 0 Then gammaTemp + 180` del VBA, y sin
    ella la fotometria se consultaria en el hemisferio equivocado."""
    _, gamma, _ = _en(0.0, -H / math.tan(math.radians(20.0)) - 5.0, 20.0)
    assert gamma > 90.0
