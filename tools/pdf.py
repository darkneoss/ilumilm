r"""Convierte un reporte.html en PDF, congelando los parametros moviles.

    python tools/pdf.py estudios/<nombre>/reporte.html \
        --clasificacion "Vias primarias" --interpostal 40 --altura 10

En pantalla el reporte deja mover clasificacion, interpostal y altura. El PDF
no puede: es un entregable fijo. Por eso este script exige elegirlos antes
(o toma los del propio estudio) y los deja escritos en el papel, en el bloque
que solo aparece al imprimir.

Requiere Playwright con Chromium instalado:
    python -m playwright install chromium
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright


def _fija(page, clasif, interpostal, altura) -> dict:
    """Mueve los controles del reporte y devuelve lo que quedo seleccionado.

    Se toca el control real y se dispara su evento, en vez de reescribir el
    DOM: asi el PDF pasa por exactamente el mismo codigo de pintado que la
    pantalla, y no hay una segunda ruta que pueda desviarse.
    """
    return page.evaluate(
        """([clasif, s, h]) => {
        const sel = document.getElementById('clasif');
        if (clasif) {
          const n = clasif.trim().toLowerCase();
          const op = [...sel.options].find(o =>
            o.value.toLowerCase() === n || o.text.toLowerCase().includes(n));
          if (!op) throw new Error('clasificación no encontrada: ' + clasif +
            ' | Opciones: ' + [...sel.options].map(o => o.text).join(' / '));
          sel.value = op.value;
          sel.dispatchEvent(new Event('change'));
        }
        const mueve = (id, valor, eje) => {
          if (valor == null) return;
          const i = eje.findIndex(x => Math.abs(x - valor) < 1e-6);
          if (i < 0) throw new Error('valor no calculado: ' + valor +
            ' | Disponibles: ' + eje.join(', '));
          const r = document.getElementById(id);
          r.value = i;
          r.dispatchEvent(new Event('input'));
        };
        mueve('interpostal', s, B.interpostales);
        mueve('altura', h, B.alturas);
        return {
          clasificacion: sel.options[sel.selectedIndex].text,
          interpostal: B.interpostales[+document.getElementById('interpostal').value],
          altura: B.alturas[+document.getElementById('altura').value],
        };
      }""",
        [clasif, interpostal, altura],
    )


def genera(html: Path, destino: Path, clasif, interpostal, altura) -> dict:
    with sync_playwright() as pw:
        navegador = pw.chromium.launch()
        page = navegador.new_page()
        page.goto(html.resolve().as_uri(), wait_until="networkidle")
        elegido = _fija(page, clasif, interpostal, altura)
        page.emulate_media(media="print")
        page.pdf(path=str(destino), format="A4", print_background=True,
                 margin={"top": "14mm", "bottom": "14mm",
                         "left": "14mm", "right": "14mm"})
        navegador.close()
    return elegido


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Reporte HTML -> PDF")
    ap.add_argument("html", type=Path)
    ap.add_argument("-o", "--salida", type=Path,
                    help="por omisión, reporte.pdf junto al HTML")
    ap.add_argument("--clasificacion",
                    help="clasificación de vialidad; acepta parte del nombre")
    ap.add_argument("--interpostal", type=float)
    ap.add_argument("--altura", type=float)
    a = ap.parse_args(argv)

    if not a.html.exists():
        print("No existe el archivo: {}".format(a.html), file=sys.stderr)
        return 1
    destino = a.salida or a.html.with_suffix(".pdf")

    try:
        elegido = genera(a.html, destino, a.clasificacion,
                         a.interpostal, a.altura)
    except Exception as exc:  # noqa: BLE001
        print("Error: {}".format(exc), file=sys.stderr)
        return 1

    print("PDF escrito en {}".format(destino))
    print("  clasificación: {}".format(elegido["clasificacion"]))
    print("  interpostal:   {} m".format(elegido["interpostal"]))
    print("  altura:        {} m".format(elegido["altura"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
