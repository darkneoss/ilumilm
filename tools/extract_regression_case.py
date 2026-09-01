"""
Extrae, a partir de los VALORES EN CACHE del libro
"assets/SEAD street lighting tool.xls" (leido con xlrd, que NUNCA evalua
formulas -- solo lee lo que Excel calculo la ultima vez que se guardo el
archivo), un caso de regresion para el motor de calculo de iluminancia.

Genera:
  - reference/caso_regresion_sead.json

Uso:
  python tools/extract_regression_case.py

Requiere: xlrd (probado con 2.0.2)

NOTA IMPORTANTE (ver reference/CASO_REGRESION.md para el detalle):
La hoja "FixtureData" contiene, en el momento en que se guardo el archivo,
los datos del luminario "A - HPS 70W Type II" (LLF, fotometria), NO los del
luminario "ATB0_20B_LED_E53_XXXXX_R2.ies" (34 W) que fue el que realmente se
uso para producir la malla de iluminancia cacheada en "Illuminance Calcs".
Ese luminario LED no existe en la libreria de 38 luminarios del libro, así
que su fotometria y su LLF NO se pueden recuperar de este archivo. Este
script deja esos campos explicitamente en null.
"""

import json
import xlrd

XLS_PATH = "assets/SEAD street lighting tool.xls"
OUT_PATH = "reference/caso_regresion_sead.json"


def cell_or_none(sheet, row, col):
    """Devuelve el valor de una celda, o None si esta vacia ('')."""
    v = sheet.cell_value(row, col)
    if v == "":
        return None
    return v


def extraer_geometria(book):
    """
    Hoja 'Road Geometry'. Las columnas C (indice 2) y D (indice 3) son
    'Baseline' y 'Upgrade' respectivamente (fila de encabezado en la fila 8,
    indice 7: 'Description | Baseline | Upgrade | Units').

    Los rangos con nombre confirmados via book.name_map (todos apuntan a
    filas index 9-19, columnas index 2/3 de esta hoja):
      BNumLanes/UNumLanes         fila idx 9  (Excel fila 10)
      BLaneWidth/ULaneWidth       fila idx 10 (Excel fila 11)
      BMedianWidth/UMedianWidth   fila idx 11 (Excel fila 12)
      BMountingHeight/UMountingHeight  fila idx 16 (Excel fila 17)
      BPoleSpacing/UPoleSpacing   fila idx 17 (Excel fila 18)
      BPoleSetback/UPoleSetback   fila idx 18 (Excel fila 19)
      BArmLength/UArmLength       fila idx 19 (Excel fila 20)

    'Pole Placement' (fila idx 15, Excel 16) no tiene rango nombrado propio
    con valor cacheado en C/D -- solo su version traducida/codificada esta
    en BFixtureArrangement/UFixtureArrangement (columnas AA/AB, indice 26/27
    de la misma fila 16 (Excel 17, 'Luminaire Mounting Height'), lo cual es
    la tabla oculta de codigos de 'fixture arrangement', NO de pole
    placement -- ver nota en el JSON de salida).

    TODAS estas celdas C/D estan vacias ('') en el archivo cacheado: la
    corrida real no dejo un guardado de las cantidades de geometria del
    camino. Se reportan como null.
    """
    sh = book.sheet_by_name("Road Geometry")

    filas_nombradas = {
        "numero_de_carriles": 9,
        "ancho_de_carril_m": 10,
        "ancho_de_camellon_m": 11,
        "altura_de_montaje_m": 16,
        "distancia_interpostal_m": 17,
        "retranqueo_de_poste_m": 18,
        "largo_de_brazo_m": 19,
    }

    geometria = {"baseline": {}, "upgrade": {}}
    celdas = {"baseline": {}, "upgrade": {}}
    for clave, fila in filas_nombradas.items():
        geometria["baseline"][clave] = cell_or_none(sh, fila, 2)
        geometria["upgrade"][clave] = cell_or_none(sh, fila, 3)
        celdas["baseline"][clave] = f"Road Geometry!{xlrd.formula.colname(2)}{fila+1}"
        celdas["upgrade"][clave] = f"Road Geometry!{xlrd.formula.colname(3)}{fila+1}"

    # Codigo de disposicion de postes / fixture arrangement (tabla oculta de
    # codificacion en columnas AA/AB de la fila "Luminaire Mounting Height").
    # Valor cacheado = 42 en ambas columnas, que NO corresponde a ninguna de
    # las 4 opciones de la tabla de traduccion (Translation!B134:B137 =
    # Single-side/Staggered/Median mounted/Opposite). Es el codigo "sin
    # seleccionar" (placeholder) del libro. Se reporta como valor bruto y se
    # marca no interpretable.
    geometria["baseline"]["disposicion_de_postes_codigo_crudo"] = cell_or_none(sh, 16, 26)
    geometria["upgrade"]["disposicion_de_postes_codigo_crudo"] = cell_or_none(sh, 16, 27)
    geometria["baseline"]["disposicion_de_postes"] = None
    geometria["upgrade"]["disposicion_de_postes"] = None
    celdas["baseline"]["disposicion_de_postes_codigo_crudo"] = "Road Geometry!AA17"
    celdas["upgrade"]["disposicion_de_postes_codigo_crudo"] = "Road Geometry!AB17"

    # Tipo de pavimento (fila 'Road Surface Type', idx 23). Baseline (col C,
    # idx2) esta vacio; Upgrade (col D, idx3) = 'Standard Surface'. El
    # encabezado de la hoja (fila idx7) marca C=Baseline, D=Upgrade, pero
    # para esta fila en particular el libro solo usa una celda compartida
    # (D) para el tipo de pavimento -- no hay version 'baseline' separada
    # cacheada.
    geometria["baseline"]["tipo_de_pavimento"] = cell_or_none(sh, 23, 2)
    geometria["upgrade"]["tipo_de_pavimento"] = cell_or_none(sh, 23, 3)
    celdas["baseline"]["tipo_de_pavimento"] = "Road Geometry!C24"
    celdas["upgrade"]["tipo_de_pavimento"] = "Road Geometry!D24"

    return geometria, celdas


def extraer_malla_iluminancia(book):
    sh = book.sheet_by_name("Illuminance Calcs")

    encabezados = [sh.cell_value(11, c) for c in range(1, 5)]

    malla = []
    r = 12
    while r < sh.nrows:
        fila = [sh.cell_value(r, c) for c in range(1, 5)]
        if not all(isinstance(v, float) for v in fila):
            break
        malla.append(fila)
        r += 1

    plano = [v for fila in malla for v in fila]
    promedio_calculado = sum(plano) / len(plano)
    minimo_calculado = min(plano)
    maximo_calculado = max(plano)
    uniformidad_calculada = promedio_calculado / minimo_calculado

    cacheados = {
        "promedio": sh.cell_value(3, 2),
        "minimo": sh.cell_value(4, 2),
        "maximo": sh.cell_value(5, 2),
        "uniformidad_promedio_sobre_minimo": sh.cell_value(6, 2),
    }

    verificacion = {
        "promedio_coincide": abs(promedio_calculado - cacheados["promedio"]) < 1e-9,
        "minimo_coincide": abs(minimo_calculado - cacheados["minimo"]) < 1e-9,
        "maximo_coincide": abs(maximo_calculado - cacheados["maximo"]) < 1e-9,
        "uniformidad_coincide": abs(uniformidad_calculada - cacheados["uniformidad_promedio_sobre_minimo"]) < 1e-9,
        "metodo_de_agregado": "media aritmetica simple sobre las 40 celdas de la malla (10 filas x 4 columnas), sin ponderar ni excluir ninguna",
    }

    return {
        "encabezados_columna": encabezados,
        "malla": malla,
        "agregados_cacheados_en_c4_c7": cacheados,
        "agregados_recalculados_desde_la_malla": {
            "promedio": promedio_calculado,
            "minimo": minimo_calculado,
            "maximo": maximo_calculado,
            "uniformidad_promedio_sobre_minimo": uniformidad_calculada,
        },
        "verificacion": verificacion,
    }


def extraer_llf(book):
    """
    Hoja 'FixtureData'. H6 (fila idx5, col idx7) = 'Total Light Loss Factor'
    = 0.6424, para la fila 'Upgrade' que en el momento del guardado apuntaba
    al luminario de la LIBRERIA 'A - HPS 70W Type II' (B6 = fila idx5,col
    idx1), NO al 'ATB0_20B_LED_E53_XXXXX_R2.ies' (34 W) usado en la corrida
    real cacheada en 'Illuminance'/'Illuminance Calcs'.

    Los tres factores (LLD, LDD, BF) para ESE luminario (HPS 70W Type II) se
    confirmaron en la hoja 'Fixtures', fila Excel 39 (idx38), columnas P/Q/R
    (idx15/16/17): LLD=0.73, LDD=0.88, BF=1.0. Producto = 0.6424, coincide
    exactamente con H6.

    El luminario LED evaluado (34 W, 'ATB0...') NO aparece en absoluto en la
    libreria de 38 luminarios de la hoja 'Fixtures' (se recorrieron las 37
    filas de luminarios reales: son todas HPS/MH/LED de la libreria estandar,
    ninguna de 34 W ni con ese nombre de archivo IES). Por lo tanto su LLD,
    LDD, BF y LLF reales NO se pueden recuperar de este archivo.
    """
    sh_fd = book.sheet_by_name("FixtureData")
    sh_fx = book.sheet_by_name("Fixtures")

    llf_h6 = sh_fd.cell_value(5, 7)
    nombre_en_h6 = sh_fd.cell_value(5, 1)

    fila_hps70 = 38  # idx de 'A - HPS 70W Type II' en Fixtures
    lld = sh_fx.cell_value(fila_hps70, 15)
    ldd = sh_fx.cell_value(fila_hps70, 16)
    bf = sh_fx.cell_value(fila_hps70, 17)
    producto = lld * ldd * bf

    # Confirmar que ningun luminario de la libreria es el evaluado.
    nombres_libreria = []
    for r in range(37, 74):
        v = sh_fx.cell_value(r, 2)
        if v:
            nombres_libreria.append(v)
    luminario_evaluado_en_libreria = any(
        "ATB0" in n or "34" in str(sh_fx.cell_value(r, 9)) for r, n in enumerate(nombres_libreria, start=37)
    )

    return {
        "advertencia": (
            "H6 en 'FixtureData' corresponde al luminario de libreria "
            f"'{nombre_en_h6}', NO al luminario LED de 34 W evaluado en la corrida cacheada."
        ),
        "llf_cacheado_en_fixturedata_h6": llf_h6,
        "luminario_al_que_realmente_corresponde_h6": nombre_en_h6,
        "factores_del_luminario_de_h6": {
            "lld_lamp_lumen_depreciation": lld,
            "ldd_luminaire_dirt_depreciation": ldd,
            "bf_ballast_factor": bf,
            "producto_lld_x_ldd_x_bf": producto,
            "coincide_con_h6": abs(producto - llf_h6) < 1e-9,
            "celdas_de_origen": {
                "lld": "Fixtures!P39",
                "ldd": "Fixtures!Q39",
                "bf": "Fixtures!R39",
                "llf_h6": "FixtureData!H6",
                "nombre": "FixtureData!B6",
            },
        },
        "luminario_evaluado_encontrado_en_libreria_fixtures": luminario_evaluado_en_libreria,
        "llf_del_luminario_led_evaluado": None,
        "lld_ldd_bf_del_luminario_led_evaluado": None,
    }


def extraer_fotometria(book):
    """
    Hoja 'FixtureData'. Estructura de bloque por luminario (confirmada
    leyendo el bloque de 'A - HPS 70W Type II', el UNICO cuya fotometria
    esta presente en el archivo cacheado):

      Sea R la fila (1-based) donde esta el NOMBRE del luminario dentro de
      la hoja FixtureData (para el primer luminario, R = FxLib_start + 1,
      donde FxLib_start es el rango nombrado que apunta a FixtureData!A41).
      A partir de ahi, usando Offset=1 (rango 'Offset'), Hrow=10 y Vrow=9
      (filas idx dentro del bloque, relativas al inicio):

        R + 0  -> Nombre del luminario
        R + 1  -> Fabricante
        R + 2  -> Tipo (LED/HPS/MH...)
        R + 3  -> Modelo
        R + 4  -> Wattage
        R + 5, R + 6 -> (sin uso en el bloque de datos; distribution type y
                          comentario viven en la hoja 'Fixtures', no aqui)
        R + 7  -> Numero de angulos verticales (columnas de la matriz de
                   candelas) = rango nombrado NumVert
        R + 8  -> Numero de angulos horizontales (filas de la matriz) =
                   rango nombrado NumHoriz
        R + 9  -> Fila de angulos VERTICALES (gamma), valores en columnas
                   B..(1+NumVert)  [rango nombrado VertDegrees]
        R + 10 .. R + 9 + NumHoriz
               -> NumHoriz filas: columna A = angulo HORIZONTAL (phi) de esa
                  fila [rango HorizDegrees], columnas B en adelante =
                  intensidad en candelas para cada angulo vertical de esa
                  fila [rango Candelaarray]

      Tamano total del bloque = 10 + NumHoriz filas. El siguiente luminario
      empieza inmediatamente despues (fila R + 10 + NumHoriz).

      Para 'A - HPS 70W Type II': NumVert=51, NumHoriz=24, bloque completo =
      FixtureData filas 42 a 75 (1-based).

    ADVERTENCIA: la hoja 'FixtureData' en el archivo cacheado SOLO contiene
    el bloque fotometrico del luminario 'A - HPS 70W Type II' (el que estaba
    seleccionado en Baseline/Upgrade -- BFixtureChoice=UFixtureChoice=3 -- al
    momento de guardar). El luminario realmente evaluado en la corrida
    cacheada ('ATB0_20B_LED_E53_XXXXX_R2.ies', LED, 34 W) NO tiene su bloque
    de angulos/candelas en ningun lugar de este libro: no esta en la
    libreria de 38 luminarios de 'Fixtures', y 'FixtureData' guarda
    unicamente los datos del ultimo luminario cargado/seleccionado en la
    interfaz -- no un historial de todos los luminarios alguna vez
    evaluados. Por lo tanto la fotometria del luminario evaluado NO es
    recuperable de este archivo, y se reporta como null.
    """
    sh = book.sheet_by_name("FixtureData")

    r0 = 41  # idx de la fila de nombre del primer luminario ('A - HPS...')
    numvert = int(sh.cell_value(r0 + 7, 1))
    numhoriz = int(sh.cell_value(r0 + 8, 1))

    vert_degrees = [sh.cell_value(r0 + 9, 1 + c) for c in range(numvert)]

    horiz_degrees = []
    candela_matrix = []
    for i in range(numhoriz):
        fila_idx = r0 + 10 + i
        horiz_degrees.append(sh.cell_value(fila_idx, 0))
        candela_matrix.append([sh.cell_value(fila_idx, 1 + c) for c in range(numvert)])

    return {
        "advertencia": (
            "Este bloque corresponde al luminario 'A - HPS 70W Type II' "
            "(el que estaba seleccionado al guardar el archivo), NO al "
            "luminario LED de 34 W ('ATB0_20B_LED_E53_XXXXX_R2.ies') "
            "evaluado en la corrida cacheada. La fotometria del luminario "
            "LED evaluado no esta presente en ningun lugar de este libro."
        ),
        "luminario_de_este_bloque": sh.cell_value(r0, 1),
        "celdas_de_origen": {
            "nombre": "FixtureData!B42",
            "num_angulos_verticales": "FixtureData!B49",
            "num_angulos_horizontales": "FixtureData!B50",
            "angulos_verticales": "FixtureData!B51:AZ51",
            "angulos_horizontales_y_matriz_candelas": "FixtureData!A52:AZ75",
        },
        "numero_de_angulos_verticales": numvert,
        "numero_de_angulos_horizontales": numhoriz,
        "angulos_verticales_grados": vert_degrees,
        "angulos_horizontales_grados": horiz_degrees,
        "matriz_de_candelas_filas_horiz_columnas_vert": candela_matrix,
        "fotometria_del_luminario_led_evaluado": {
            "nombre_ies": "ATB0_20B_LED_E53_XXXXX_R2.ies",
            "tipo": "LED",
            "watts": 34.0,
            "angulos_verticales_grados": None,
            "angulos_horizontales_grados": None,
            "matriz_de_candelas": None,
            "nota": "No recuperable del archivo cacheado; ver advertencia arriba.",
        },
    }


def main():
    book = xlrd.open_workbook(XLS_PATH, formatting_info=False)

    geometria, celdas_geometria = extraer_geometria(book)
    malla = extraer_malla_iluminancia(book)
    llf = extraer_llf(book)
    fotometria = extraer_fotometria(book)

    salida = {
        "meta": {
            "origen": XLS_PATH,
            "descripcion": (
                "Caso de regresion extraido de los valores cacheados de una corrida real "
                "del libro Excel SEAD Street Lighting Tool, hoja por hoja. Generado con "
                "tools/extract_regression_case.py, que lee unicamente valores (xlrd no evalua formulas)."
            ),
            "luminario_evaluado_segun_hoja_illuminance": {
                "nombre_ies": "ATB0_20B_LED_E53_XXXXX_R2.ies",
                "tipo": "LED",
                "watts": 34.0,
                "celda_origen": "Illuminance!B4:D4 (fila 'RESULTS - 1')",
            },
        },
        "geometria_entrada": geometria,
        "geometria_entrada_celdas_origen": celdas_geometria,
        "malla_iluminancia_esperada": malla,
        "factor_de_perdidas_llf": llf,
        "fotometria_luminario": fotometria,
    }

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, indent=2)

    print(f"Escrito: {OUT_PATH}")
    print("Verificacion de la malla:", malla["verificacion"])
    print("Verificacion del LLF:", llf["factores_del_luminario_de_h6"]["coincide_con_h6"])


if __name__ == "__main__":
    main()
