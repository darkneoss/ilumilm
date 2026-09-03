"""Humo del reporte: que se genere y que no se le caigan las anclas del script.

El reporte es el entregable, y hasta ahora no tenia ni una prueba. Estas no
juzgan como se ve --eso hay que mirarlo en un navegador-- sino que el HTML
salga completo y que sigan ahi las dos cosas de las que depende el JavaScript.

La segunda nacio de un bug real: al agregar la tabla de fotometria a los datos
de planificacion, el script dejo de funcionar por completo. Buscaba las filas
de la comparativa con `querySelectorAll('tbody tr')`, sin acotar, y se llevaba
tambien las filas de la tabla nueva; al buscarles celdas que no tienen, moria
antes de escribir un solo numero. Una prueba de Python no puede ejecutar ese
JavaScript, pero si puede vigilar el ancla de la que ahora depende.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from engine import report
from engine.cli import ejecuta

RAIZ = Path(__file__).resolve().parent.parent
IES = RAIZ / "catalogo" / "V1050UN2M50.ies"

ENTRADA = {
    "vialidad": {"num_carriles": 2, "ancho_carril": 3.5, "camellon": 1.0,
                 "disposicion": "central doble", "altura_montaje": 8.0,
                 "interpostal": 35.0, "retranqueo": 0.0, "largo_brazo": 1.8,
                 "banqueta": 3.0},
    "nom": {"clasificacion_vialidad": "vias_secundarias_residencial_tipo_a",
            "pavimento": "R2"},
    "luminarios": [{"archivo": "V1050UN2M50.ies", "inclinacion": 5.0}],
}


@pytest.fixture(scope="module")
def html(tmp_path_factory):
    if not IES.exists():
        pytest.skip("falta catalogo/V1050UN2M50.ies; los .ies no se versionan")
    d = tmp_path_factory.mktemp("estudio")
    ruta = d / "entrada.json"
    ruta.write_text(json.dumps(ENTRADA), encoding="utf-8")
    return report.html(ejecuta(ruta))


def test_el_ancla_de_la_comparativa_sigue_ahi(html):
    """Si esto falla, el reporte se ve bien y no hace NADA al mover los
    controles: sin `#comparativa` el script no encuentra las filas."""
    assert 'id="comparativa"' in html


def test_las_filas_con_celdas_de_calculo_estan_dentro_de_la_comparativa(html):
    """La otra mitad del mismo bug: que ninguna tabla ajena traiga filas que el
    script pueda confundir con las de la comparativa."""
    celda = '<td class="n c-eprom"'          # el marcado, no la referencia del script
    cuerpo = html[html.index('id="comparativa"'):]
    dentro = cuerpo[:cuerpo.index("</tbody>")]
    assert html.count(celda) == dentro.count(celda) > 0


def test_los_datos_de_planificacion_van_completos(html):
    """Incluidos los dos esquemas acotados: la seccion transversal (que lleva
    `data-seccion`, de donde el script la toma para seguir a la altura) y la
    planta, cuya cota de interpostal dice "entre postes"."""
    for texto in ("Datos de planificación", "Perfil de la vía pública",
                  "Disposición de los luminarios", "Camino peatonal",
                  "Factor de mantenimiento", "Inclinación del brazo",
                  'data-seccion="1"', "entre postes"):
        assert texto in html, texto


def test_la_inclinacion_se_imprime_aunque_sea_cero():
    """Es la leccion de la validacion: la herramienta que se sustituye usa la
    inclinacion en el calculo y no la escribe, y dos estudios con
    inclinaciones distintas salian indistinguibles."""
    if not IES.exists():
        pytest.skip("falta la fotometria")
    import tempfile

    entrada = json.loads(json.dumps(ENTRADA))
    entrada["luminarios"] = [{"archivo": "V1050UN2M50.ies"}]   # sin inclinacion
    with tempfile.TemporaryDirectory() as d:
        ruta = Path(d) / "entrada.json"
        ruta.write_text(json.dumps(entrada), encoding="utf-8")
        h = report.html(ejecuta(ruta))
    assert "Inclinación del brazo" in h
    assert re.search(r"Inclinación del brazo</span><span class=\"v\">0\.0°", h)


def test_en_central_doble_no_se_finge_un_retranqueo(html):
    """El motor lo ignora a proposito en un poste central; imprimir el numero
    invitaria a creer que se uso."""
    assert "no aplica (poste central)" in html
    assert "Brazos desde el eje del camellón" in html


def test_cada_fila_de_la_comparativa_apunta_a_su_detalle(html):
    """La fila lleva al detalle con un ancla de verdad, no con un `onclick`:
    asi funciona sin JavaScript, se puede abrir en otra pestana y el teclado la
    alcanza. El script solo amplia el area sensible al resto de la fila.

    Se comprueban los dos extremos del salto porque un `href` sin destino no
    falla a la vista: simplemente no lleva a ningun lado.
    """
    destinos = re.findall(r'<a class="mnombre" href="#(lum-\d+)"', html)
    anclas = re.findall(r'<section class="detalle" id="(lum-\d+)"', html)
    assert destinos, "ninguna fila enlaza a su detalle"
    assert destinos == anclas


def test_el_pavimento_se_puede_mover_en_el_reporte(html):
    """El pavimento no entra en el calculo de iluminancia: solo elige la tabla
    de umbrales de la norma. Por eso se resuelve en el navegador, como la
    clasificacion, y por eso el reporte tiene que llevar las CUATRO tablas y no
    solo la del estudio."""
    assert 'id="pav"' in html
    datos = json.loads(re.search(r"const REQ = (\{.*?\});\n", html, re.S).group(1))
    assert set(datos) == {"R1", "R2", "R3", "R4"}
    # y los umbrales de verdad cambian entre tablas, que es lo que hace util el
    # control: si fueran iguales, el selector seria decorado
    eprom = {p: t["vias_primarias_y_colectoras"]["eprom_min"] for p, t in datos.items()}
    assert len(set(eprom.values())) > 1, eprom


def test_el_pavimento_ya_no_esta_en_el_bloque_de_planificacion(html):
    """Regla del bloque: solo lo que los controles NO mueven. Un dato repetido
    ahi diria un valor distinto al del selector en cuanto alguien lo tocara."""
    plan = html[html.index('<section class="plan">'):html.index('id="comparativa"')]
    assert ">Pavimento<" not in plan


def test_el_reporte_dice_con_que_version_se_calculo(html):
    """Un reporte que no dice que codigo lo produjo es infalsificable pero
    inauditable, que es el mismo defecto de la herramienta que se sustituye
    --usa la inclinacion y no la escribe--. La version es lo unico que enlaza
    una memoria de calculo entregada con el motor que la genero."""
    from engine import __version__
    assert "Calculado con <strong>ilumilm %s</strong>" % __version__ in html


def test_el_reporte_distingue_quien_calculo_de_quien_dibujo():
    """`resultados.json` se puede volver a renderizar mas tarde con un motor
    mas nuevo, y los numeros seguirian siendo los del viejo. Cuando difieren se
    dicen las dos versiones: lo que hay que poder auditar es de donde salieron
    los numeros, no quien pinto la pagina."""
    from engine import __version__, report

    if not IES.exists():
        pytest.skip("falta la fotometria")
    import tempfile

    with tempfile.TemporaryDirectory() as d:
        ruta = Path(d) / "entrada.json"
        ruta.write_text(json.dumps(ENTRADA), encoding="utf-8")
        datos = ejecuta(ruta)

    assert datos["version_motor"] == __version__

    viejo = dict(datos, version_motor="0.9.0")
    h = report.html(viejo)
    assert "ilumilm 0.9.0" in h and "el reporte lo generó la %s" % __version__ in h

    sin_version = {k: v for k, v in datos.items() if k != "version_motor"}
    assert "anterior a la 1.0.0" in report.html(sin_version)
