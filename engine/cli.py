r"""Ejecuta un estudio completo de iluminación vial contra la NOM-013-ENER-2013.

Uso:
    python -m engine.cli estudios/<nombre>/entrada.json

Escribe, junto al JSON de entrada:
    - malla.csv        (una copia de la malla del PRIMER luminario evaluado en
                         modo "correcto"; cada luminario trae también su
                         propia malla completa dentro de resultados.json)
    - resultados.json  (comparación de todos los luminarios, ordenada)

Formato de `entrada.json`
--------------------------
{
  "vialidad": {
    "num_carriles": 2,
    "ancho_carril": 3.5,
    "camellon": 0.0,
    "disposicion": "unilateral",
    "altura_montaje": 8.0,
    "interpostal": 35.0,
    "retranqueo": 0.2,
    "largo_brazo": 1.8
  },
  "nom": {
    "clasificacion_vialidad": "Vías secundarias residencial Tipo A",
    "pavimento": "R2"
  },
  "perdidas": { "lld": 0.85, "ldd": 0.90, "bf": 1.0 },
  "modo_azimut": "correcto",
  "luminarios": [
    { "archivo": "V1050UN2M50.ies" },
    { "archivo": "V2100UN2M50.ies", "watts": 100.0 }
  ]
}

Notas de diseño:

- `luminarios[].archivo` puede ser una ruta a un .ies (absoluta o relativa al
  directorio del propio `entrada.json`) o el nombre de un archivo ya indexado
  en `catalogo/` (se busca ahí si la ruta directa no existe).
- `watts` es opcional y, si viene, SOBREESCRIBE los watts del archivo .ies
  (declarados o de texto) para el cálculo de DPEA. Útil cuando el .ies no
  declara watts o cuando se quiere evaluar un driver distinto al de fábrica.
- `modo_azimut` es opcional, por omisión "correcto" (ver docstring de
  `engine.calc`). Se acepta también "cuadrante", solo para diagnóstico.

Sobre el DPEA y la disposición de postes
-----------------------------------------
La NOM-013 pide el DPEA sobre la "carga total conectada" de un tramo. Cuántos
luminarios hay por tramo interpostal depende de la disposición (ver
`engine.geometry.DISPOSICIONES`):

    unilateral        -> 1 luminario por tramo (un solo lado de postes)
    tresbolillo        -> 2 luminarios por tramo (postes alternados en ambos lados)
    bilateral opuesta -> 2 luminarios por tramo (un poste enfrentado en cada lado)
    central doble     -> 2 luminarios por tramo (doble brazo en el camellón, un
                          poste central sirve ambos sentidos con 2 luminarias)

Este número (1 o 2) se multiplica por los watts del luminario para obtener la
carga conectada por tramo, y esa carga es la que entra a `nom.dpea` junto con
el ancho sin camellón y el interpostal como largo del tramo.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from . import calc, catalogo, ies, nom
from .geometry import Vialidad

# Disposiciones (ya normalizadas por Vialidad, forma interna del VBA) que
# tienen 2 luminarios por tramo interpostal en vez de 1.
_DISPOSICIONES_DOBLES = {"Staggered", "Opposite", "Median mounted"}


class ErrorEstudio(Exception):
    """Error de datos de entrada o de proceso, para mostrar en español y sin
    traceback crudo al usuario."""


def _luminarios_por_tramo(disposicion_normalizada: str) -> int:
    """1 para unilateral, 2 para tresbolillo/bilateral/central doble."""
    return 2 if disposicion_normalizada in _DISPOSICIONES_DOBLES else 1


def _lee_entrada(ruta_json: Path) -> Dict[str, Any]:
    if not ruta_json.exists():
        raise ErrorEstudio(f"No se encontró el archivo de entrada: {ruta_json}")
    try:
        return json.loads(ruta_json.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ErrorEstudio(f"El JSON de entrada no es válido: {exc}") from exc


def _campo(d: Dict[str, Any], clave: str, contenedor: str) -> Any:
    if clave not in d:
        raise ErrorEstudio(f"Falta el campo '{clave}' dentro de '{contenedor}' en la entrada")
    return d[clave]


def _arma_vialidad(datos: Dict[str, Any]) -> Vialidad:
    v = _campo(datos, "vialidad", "entrada.json")
    try:
        return Vialidad(
            num_carriles=int(_campo(v, "num_carriles", "vialidad")),
            ancho_carril=float(_campo(v, "ancho_carril", "vialidad")),
            camellon=float(v.get("camellon", 0.0)),
            disposicion=str(_campo(v, "disposicion", "vialidad")),
            altura_montaje=float(_campo(v, "altura_montaje", "vialidad")),
            interpostal=float(_campo(v, "interpostal", "vialidad")),
            retranqueo=float(_campo(v, "retranqueo", "vialidad")),
            largo_brazo=float(_campo(v, "largo_brazo", "vialidad")),
        )
    except (TypeError, ValueError) as exc:
        raise ErrorEstudio(f"Datos de vialidad inválidos: {exc}") from exc


def _resuelve_ruta_ies(nombre_o_ruta: str, dir_entrada: Path) -> Path:
    """Busca el .ies como ruta directa; si no existe, en el catálogo local."""
    candidata = Path(nombre_o_ruta)
    if not candidata.is_absolute():
        candidata = dir_entrada / nombre_o_ruta
    if candidata.exists():
        return candidata

    # Segundo intento: nombre de archivo dentro de catalogo/
    dir_catalogo = Path("catalogo")
    candidata_catalogo = dir_catalogo / Path(nombre_o_ruta).name
    if candidata_catalogo.exists():
        return candidata_catalogo

    raise ErrorEstudio(
        f"No se encontró el luminario '{nombre_o_ruta}' (se buscó como ruta "
        f"y dentro de '{dir_catalogo}/')"
    )


def _evalua_luminario(
    v: Vialidad,
    dato_luminario: Dict[str, Any],
    llf_total: float,
    modo: str,
    vialidad_nom: str,
    pavimento_nom: str,
    dir_entrada: Path,
) -> Dict[str, Any]:
    nombre_o_ruta = _campo(dato_luminario, "archivo", "luminarios[]")
    ruta = _resuelve_ruta_ies(nombre_o_ruta, dir_entrada)

    try:
        foto = ies.lee(ruta)
    except Exception as exc:  # noqa: BLE001
        raise ErrorEstudio(f"No se pudo interpretar el archivo .ies '{ruta.name}': {exc}") from exc

    watts_sobrescritos = dato_luminario.get("watts")
    if watts_sobrescritos is not None:
        watts = float(watts_sobrescritos)
        origen_watts = "sobrescrito en entrada.json"
    else:
        watts = foto.watts if foto.watts is not None else foto.watts_del_texto()
        origen_watts = "declarado en .ies" if foto.watts is not None else "texto libre del .ies"
        if watts is None:
            raise ErrorEstudio(
                f"El archivo '{ruta.name}' no declara watts ni se pudieron rescatar de "
                "texto libre; indica 'watts' en la entrada para este luminario"
            )

    malla = calc.calcula(v, foto, llf_total, modo)
    comparacion_modos = calc.comparar_modos(v, foto, llf_total)

    n_luminarios_tramo = _luminarios_por_tramo(v.disposicion)
    watts_conectados_tramo = watts * n_luminarios_tramo
    dpea_calculado = nom.dpea(
        watts_conectados_tramo, v.ancho_sin_camellon, v.interpostal
    )

    resultado_nom = nom.evaluar(
        eprom=malla.promedio,
        emin=malla.minimo,
        dpea_calculado=dpea_calculado,
        vialidad=vialidad_nom,
        pavimento=pavimento_nom,
        ancho_m=v.ancho_sin_camellon,
    )

    return foto, watts, {
        "archivo": str(ruta),
        "catalogo": foto.catalogo,
        "fabricante": foto.fabricante,
        "watts": watts,
        "origen_watts": origen_watts,
        "n_luminarios_por_tramo": n_luminarios_tramo,
        "watts_conectados_tramo": watts_conectados_tramo,
        "dpea_w_m2": dpea_calculado,
        "malla": {
            "modo": malla.modo,
            "promedio_lx": malla.promedio,
            "minimo_lx": malla.minimo,
            "maximo_lx": malla.maximo,
            "uniformidad": malla.uniformidad,
            "xs": malla.xs,
            "ys": malla.ys,
            "e": malla.e,
        },
        "comparacion_modos": comparacion_modos,
        "nom": {
            "vialidad": resultado_nom.vialidad,
            "pavimento": resultado_nom.pavimento,
            "ancho_m": resultado_nom.ancho_m,
            "cumple": resultado_nom.cumple,
            "criterios": {
                "nivel_iluminacion": vars(resultado_nom.nivel_iluminacion),
                "uniformidad": vars(resultado_nom.uniformidad),
                "dpea": vars(resultado_nom.dpea),
            },
        },
    }


# ---------------------------------------------------------------------------
# Barrido de interpostal x altura de montaje
# ---------------------------------------------------------------------------

# Rangos usuales de alumbrado vial. Se incluyen siempre los valores del propio
# estudio, aunque no caigan en la retícula, para que el reporte pueda volver a
# la configuración especificada de forma exacta.
INTERPOSTALES = [20.0 + 2.5 * i for i in range(13)]      # 20 a 50 m
ALTURAS = [6.0 + 0.5 * i for i in range(13)]             # 6 a 12 m


def _eje(valores: List[float], propio: float) -> List[float]:
    """Retícula del barrido con el valor del estudio insertado en su lugar."""
    if any(abs(x - propio) < 1e-9 for x in valores):
        return list(valores)
    return sorted(valores + [propio])


def _barrido(
    v: Vialidad,
    fotos: List[Any],
    watts: List[float],
    llf_total: float,
    modo: str,
) -> Dict[str, Any]:
    """Precalcula la malla para cada combinación de interpostal y altura.

    La clasificación de vialidad se resuelve en el navegador porque solo mueve
    umbrales, pero la interpostal y la altura cambian la geometría, los ángulos
    y el área del tramo, así que hay que volver a correr el motor. A 0.6 ms por
    corrida el barrido completo cuesta menos de un segundo.

    Las mallas se guardan con un decimal, que es exactamente la precisión con
    que el reporte las dibuja. Los agregados van con precisión completa, para
    que un veredicto justo en el límite no dependa del redondeo del mapa.
    """
    interpostales = _eje(INTERPOSTALES, v.interpostal)
    alturas = _eje(ALTURAS, v.altura_montaje)
    n_tramo = _luminarios_por_tramo(v.disposicion)

    mallas: List[List[List[List[float]]]] = []
    stats: List[List[List[List[float]]]] = []
    xs_por_interpostal: List[List[float]] = []

    for s_ in interpostales:
        fila_m, fila_s = [], []
        for h in alturas:
            vv = Vialidad(v.num_carriles, v.ancho_carril, v.camellon, v.disposicion,
                          h, s_, v.retranqueo, v.largo_brazo)
            celda_m, celda_s = [], []
            for foto in fotos:
                m = calc.calcula(vv, foto, llf_total, modo)
                celda_m.append([round(x, 1) for fila in m.e for x in fila])
                celda_s.append([m.promedio, m.minimo, m.maximo, m.uniformidad])
            fila_m.append(celda_m)
            fila_s.append(celda_s)
        mallas.append(fila_m)
        stats.append(fila_s)
        # Las X de la ventana evaluada dependen del interpostal; las Y no.
        vv = Vialidad(v.num_carriles, v.ancho_carril, v.camellon, v.disposicion,
                      v.altura_montaje, s_, v.retranqueo, v.largo_brazo)
        m = calc.calcula(vv, fotos[0], llf_total, modo)
        xs_por_interpostal.append([round(x, 3) for x in m.xs])

    return {
        "interpostales": interpostales,
        "alturas": alturas,
        "i_interpostal_estudio": interpostales.index(v.interpostal),
        "i_altura_estudio": alturas.index(v.altura_montaje),
        "xs_por_interpostal": xs_por_interpostal,
        "ys": [round(y, 3) for y in calc.valores_y(v.num_carriles, v.ancho_carril, v.camellon)],
        "watts_por_tramo": [w * n_tramo for w in watts],
        "ancho_sin_camellon": v.ancho_sin_camellon,
        "mallas": mallas,
        "stats": stats,
    }


def ejecuta(ruta_json: Path) -> Dict[str, Any]:
    """Corre el estudio completo y devuelve el dict que se escribe a JSON."""
    dir_entrada = ruta_json.parent
    datos = _lee_entrada(ruta_json)

    v = _arma_vialidad(datos)

    datos_nom = _campo(datos, "nom", "entrada.json")
    vialidad_nom = _campo(datos_nom, "clasificacion_vialidad", "nom")
    pavimento_nom = _campo(datos_nom, "pavimento", "nom")
    try:
        vialidad_nom = nom.resolver_vialidad(vialidad_nom)
        pavimento_nom = nom.resolver_pavimento(pavimento_nom)
    except ValueError as exc:
        raise ErrorEstudio(str(exc)) from exc

    perdidas = datos.get("perdidas", {})
    llf_total = calc.llf(
        lld=float(perdidas.get("lld", 0.85)),
        ldd=float(perdidas.get("ldd", 0.90)),
        bf=float(perdidas.get("bf", 1.0)),
    )

    modo = datos.get("modo_azimut", calc.MODO_CORRECTO)
    if modo not in (calc.MODO_CUADRANTE, calc.MODO_CORRECTO):
        raise ErrorEstudio(
            f"modo_azimut desconocido: '{modo}' (usa 'correcto' o 'cuadrante')"
        )

    lista_luminarios = _campo(datos, "luminarios", "entrada.json")
    if not lista_luminarios:
        raise ErrorEstudio("La entrada no trae ningún luminario en 'luminarios'")

    evaluados: List[Any] = []
    for dato_luminario in lista_luminarios:
        evaluados.append(
            _evalua_luminario(
                v, dato_luminario, llf_total, modo, vialidad_nom, pavimento_nom, dir_entrada
            )
        )

    # Orden: primero los que cumplen, y entre ellos menor potencia primero.
    evaluados.sort(key=lambda e: (not e[2]["nom"]["cumple"], e[2]["watts"]))
    fotos = [e[0] for e in evaluados]
    watts_lista = [e[1] for e in evaluados]
    resultados = [e[2] for e in evaluados]

    salida = {
        "entrada": str(ruta_json),
        "vialidad": {
            "num_carriles": v.num_carriles,
            "ancho_carril": v.ancho_carril,
            "camellon": v.camellon,
            "disposicion": v.disposicion,
            "altura_montaje": v.altura_montaje,
            "interpostal": v.interpostal,
            "retranqueo": v.retranqueo,
            "largo_brazo": v.largo_brazo,
            "ancho_calzada": v.ancho_calzada,
            "ancho_sin_camellon": v.ancho_sin_camellon,
        },
        "nom": {"clasificacion_vialidad": vialidad_nom, "pavimento": pavimento_nom},
        "perdidas": {
            "lld": perdidas.get("lld", 0.85),
            "ldd": perdidas.get("ldd", 0.90),
            "bf": perdidas.get("bf", 1.0),
            "llf_total": llf_total,
        },
        "modo_azimut": modo,
        "resultados": resultados,
        "barrido": _barrido(v, fotos, watts_lista, llf_total, modo),
        "recomendacion": resultados[0]["catalogo"] if resultados[0]["nom"]["cumple"] else None,
    }
    return salida


def main(argv: Optional[List[str]] = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) != 1:
        print("Uso: python -m engine.cli estudios/<nombre>/entrada.json", file=sys.stderr)
        return 2

    ruta_json = Path(argv[0])
    try:
        salida = ejecuta(ruta_json)
    except ErrorEstudio as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    dir_salida = ruta_json.parent
    ruta_resultados = dir_salida / "resultados.json"
    ruta_resultados.write_text(
        json.dumps(salida, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # malla.csv: la del primer resultado (el recomendado / el mejor ordenado).
    primero = salida["resultados"][0]
    xs = primero["malla"]["xs"]
    ys = primero["malla"]["ys"]
    e = primero["malla"]["e"]
    malla = calc.Malla(xs=xs, ys=ys, e=e, modo=primero["malla"]["modo"])
    ruta_csv = dir_salida / "malla.csv"
    ruta_csv.write_text(malla.a_csv(), encoding="utf-8")

    print(f"Estudio calculado: {len(salida['resultados'])} luminario(s) evaluados")
    print(f"  -> {ruta_resultados}")
    print(f"  -> {ruta_csv} (malla del luminario mejor clasificado: {primero['catalogo']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
