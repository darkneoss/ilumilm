"""Catálogo local de archivos fotométricos .ies.

Evita tener que rastrear rutas de archivos .ies a mano en cada estudio: se
indexa una vez un directorio de catálogo (por omisión `catalogo/`) y luego se
puede buscar por modelo, fabricante o potencia.

El índice se guarda como `catalogo/index.json` (UTF-8, indentado). Cada
entrada trae:

    archivo      -> nombre del archivo .ies (relativo al directorio)
    catalogo     -> clave LUMCAT / nombre del modelo
    fabricante   -> keyword MANUFAC
    watts_declarados -> campo watts del .ies, o None si viene en 0
    watts_texto  -> potencia rescatada de texto libre (heurística), o None
    lumenes      -> lúmenes por lámpara, o "absoluta" si el archivo es de
                    fotometría absoluta (lumenes_lampara < 0)
    simetria     -> "rotacional" | "cuadrante" | "bilateral" | "ninguna"
    n_angulos_v  -> número de ángulos verticales
    n_angulos_h  -> número de ángulos horizontales
    eficacia_lm_w -> lúmenes / watts, cuando ambos se pueden determinar

Un .ies que no logre parsearse NO tumba la indexación completa: se registra
bajo la clave "errores" del índice, con el nombre de archivo y el mensaje de
error, y la indexación continúa con el resto.
"""
from __future__ import annotations

import json
import unicodedata
from pathlib import Path
from typing import Any, Dict, List

from . import ies

NOMBRE_INDICE = "index.json"


def _sin_acentos(texto: str) -> str:
    """Quita acentos y pasa a minúsculas, para comparaciones tolerantes."""
    descompuesto = unicodedata.normalize("NFKD", texto)
    sin_acentos = "".join(c for c in descompuesto if not unicodedata.combining(c))
    return sin_acentos.lower()


def _entrada_desde_fotometria(nombre_archivo: str, foto: "ies.Fotometria") -> Dict[str, Any]:
    """Arma la entrada de índice para una `Fotometria` ya parseada."""
    watts_decl = foto.watts
    watts_texto = foto.watts_del_texto()
    watts_para_eficacia = watts_decl if watts_decl is not None else watts_texto

    if foto.absoluta:
        lumenes: Any = "absoluta"
        eficacia = None
    else:
        lumenes = foto.lumenes_lampara
        if watts_para_eficacia:
            eficacia = round(lumenes / watts_para_eficacia, 1)
        else:
            eficacia = None

    return {
        "archivo": nombre_archivo,
        "catalogo": foto.catalogo,
        "fabricante": foto.fabricante,
        "watts_declarados": watts_decl,
        "watts_texto": watts_texto,
        "lumenes": lumenes,
        "simetria": foto.simetria,
        "n_angulos_v": len(foto.angulos_v),
        "n_angulos_h": len(foto.angulos_h),
        "eficacia_lm_w": eficacia,
    }


def indexa(dir_catalogo: str = "catalogo") -> List[Dict[str, Any]]:
    """Recorre los .ies de `dir_catalogo`, los parsea y guarda `index.json`.

    Devuelve la lista de entradas indexadas (sin incluir los errores; esos
    quedan solo dentro del archivo, bajo la clave "errores").
    """
    directorio = Path(dir_catalogo)
    directorio.mkdir(parents=True, exist_ok=True)

    entradas: List[Dict[str, Any]] = []
    errores: List[Dict[str, str]] = []

    for ruta in sorted(directorio.glob("*.ies")):
        try:
            foto = ies.lee(ruta)
        except Exception as exc:  # noqa: BLE001 - un .ies malo no debe tumbar el resto
            errores.append({"archivo": ruta.name, "error": str(exc)})
            continue
        entradas.append(_entrada_desde_fotometria(ruta.name, foto))

    salida = {"luminarios": entradas, "errores": errores}
    ruta_indice = directorio / NOMBRE_INDICE
    ruta_indice.write_text(
        json.dumps(salida, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return entradas


def carga(dir_catalogo: str = "catalogo") -> List[Dict[str, Any]]:
    """Lee `index.json`; reindexa si no existe o si hay .ies más nuevos.

    "Más nuevo" se determina por fecha de modificación del archivo .ies
    contra la del propio `index.json`.
    """
    directorio = Path(dir_catalogo)
    ruta_indice = directorio / NOMBRE_INDICE

    if not ruta_indice.exists():
        return indexa(dir_catalogo)

    mtime_indice = ruta_indice.stat().st_mtime
    hay_mas_nuevo = any(
        p.stat().st_mtime > mtime_indice for p in directorio.glob("*.ies")
    )
    if hay_mas_nuevo:
        return indexa(dir_catalogo)

    datos = json.loads(ruta_indice.read_text(encoding="utf-8"))
    return datos.get("luminarios", [])


def busca(termino: str, dir_catalogo: str = "catalogo") -> List[Dict[str, Any]]:
    """Búsqueda tolerante por modelo, fabricante o potencia.

    Sin distinguir acentos ni mayúsculas. Un término como "50 W" busca la
    potencia (declarada o de texto) redondeada a entero; cualquier otro
    término se busca como subcadena en catálogo/fabricante/archivo.
    """
    clave = _sin_acentos(termino.strip())
    entradas = carga(dir_catalogo)

    # ¿El término trae una potencia tipo "50w" o "50 w"?
    import re
    m = re.match(r"^(\d+(?:[.,]\d+)?)\s*w(?:atts?)?$", clave)
    if m:
        objetivo = float(m.group(1).replace(",", "."))
        resultado = []
        for e in entradas:
            for campo in ("watts_declarados", "watts_texto"):
                w = e.get(campo)
                if w is not None and abs(w - objetivo) < 0.5:
                    resultado.append(e)
                    break
        return resultado

    resultado = []
    for e in entradas:
        campos = [e.get("catalogo", ""), e.get("fabricante", ""), e.get("archivo", "")]
        if any(clave in _sin_acentos(str(c)) for c in campos):
            resultado.append(e)
    return resultado


def _tabla(entradas: List[Dict[str, Any]]) -> str:
    """Formatea las entradas como tabla legible de ancho fijo."""
    if not entradas:
        return "(catálogo vacío)"
    encabezados = ["Archivo", "Catálogo", "Fabricante", "W", "Lúmenes", "lm/W", "Simetría"]
    filas = []
    for e in entradas:
        w = e.get("watts_declarados")
        if w is None:
            w = e.get("watts_texto")
        lm = e.get("lumenes")
        ef = e.get("eficacia_lm_w")
        filas.append([
            e.get("archivo", ""),
            e.get("catalogo", ""),
            e.get("fabricante", ""),
            "{:.1f}".format(w) if isinstance(w, (int, float)) else "?",
            "{:.0f}".format(lm) if isinstance(lm, (int, float)) else str(lm),
            "{:.1f}".format(ef) if isinstance(ef, (int, float)) else "?",
            e.get("simetria", ""),
        ])
    anchos = [max(len(str(h)), *(len(f[i]) for f in filas)) for i, h in enumerate(encabezados)]
    def formatea(fila):
        return "  ".join(str(c).ljust(anchos[i]) for i, c in enumerate(fila))
    lineas = [formatea(encabezados), formatea(["-" * a for a in anchos])]
    lineas.extend(formatea(f) for f in filas)
    return "\n".join(lineas)


if __name__ == "__main__":
    entradas_ = carga()
    print(_tabla(entradas_))
    datos_ = json.loads((Path("catalogo") / NOMBRE_INDICE).read_text(encoding="utf-8"))
    if datos_.get("errores"):
        print("\nArchivos con error al parsear:")
        for err in datos_["errores"]:
            print("  {}: {}".format(err["archivo"], err["error"]))
