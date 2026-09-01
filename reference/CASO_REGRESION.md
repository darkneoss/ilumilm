# Caso de regresión SEAD — de dónde sale cada dato

Fuente única: `assets/SEAD street lighting tool.xls`, leído con `xlrd`
(`formatting_info=False`), que solo expone **valores en caché** — la última
corrida que Excel guardó — nunca fórmulas. Script que regenera el JSON:
`tools/extract_regression_case.py`. Salida: `reference/caso_regresion_sead.json`.

## 1. Geometría de entrada (hoja "Road Geometry")

Se localizaron las celdas exactas vía `book.name_map` (rangos con nombre):

| Rango con nombre | Celda Baseline | Celda Upgrade | Valor cacheado |
|---|---|---|---|
| BNumLanes / UNumLanes | Road Geometry!C10 | !D10 | *(vacío)* |
| BLaneWidth / ULaneWidth | !C11 | !D11 | *(vacío)* |
| BMedianWidth / UMedianWidth | !C12 | !D12 | *(vacío)* |
| BMountingHeight / UMountingHeight | !C17 | !D17 | *(vacío)* |
| BPoleSpacing / UPoleSpacing | !C18 | !D18 | *(vacío)* |
| BPoleSetback / UPoleSetback | !C19 | !D19 | *(vacío)* |
| BArmLength / UArmLength | !C20 | !D20 | *(vacío)* |
| BFixtureArrangement / UFixtureArrangement | !AA17 | !AB17 | 42 (ambas) |
| Road Surface Type (sin rango nombrado, fila fija) | !C24 | !D24 | *(vacío)* / "Standard Surface" |

**Hallazgo principal:** todas las celdas numéricas de geometría (carriles,
ancho de carril, camellón, altura de montaje, separación de postes,
retranqueo, largo de brazo) están **vacías** en el archivo cacheado. La hoja
"Road Geometry" nunca se rellenó para esta corrida, o se limpió después de
correrla.

El código 42 de `BFixtureArrangement`/`UFixtureArrangement` **no** corresponde
a ninguna de las 4 opciones de la tabla de traducción real
(`Translation!B134:B137` = Single-side / Staggered / Median mounted /
Opposite). Es, con alta probabilidad, el valor centinela que usa el libro
para "sin seleccionar" (aparece también como relleno en filas enteras de la
hoja "Illuminance", ver más abajo). Por eso se deja como
`disposicion_de_postes: null` y solo se reporta el código crudo.

`Road Surface Type` tiene una peculiaridad: aunque el encabezado de la hoja
(fila 8) declara C=Baseline y D=Upgrade, esta fila en particular solo usa la
celda D ("Standard Surface"); C está vacía. No hay un valor "baseline"
independiente cacheado para el tipo de pavimento.

**Conclusión de esta sección: la geometría de entrada NO es recuperable de
este archivo.** Ningún valor numérico de geometría quedó guardado.

## 2. Malla de iluminancia (hoja "Illuminance Calcs")

- C4 = Promedio = 4.9699814242012375
- C5 = Mínimo = 3.717512649408743
- C6 = Máximo = 6.832462419176224
- C7 = Uniformidad (Iprom/Imin) = 1.336910427188915
- Encabezados de fila 12: "Lane 1 - 1/4 lane", "Lane 1 - 3/4 lane", "Lane 2 - 1/4 lane", "Lane 2 - 3/4 lane"
- Malla completa: filas 13 a 22 (10 filas × 4 columnas = 40 puntos), volcada íntegra en el JSON (`malla_iluminancia_esperada.malla`).

**Verificación numérica (confirmada, ver salida del script):**

```
promedio_coincide: True
minimo_coincide:   True
maximo_coincide:   True
uniformidad_coincide: True
```

La media aritmética simple de las 40 celdas reproduce exactamente
4.9699814242012375; el mínimo y máximo de la malla reproducen exactamente
C5 y C6. **Confirmado: el agregado es media aritmética simple sobre toda la
malla, sin ponderar filas/columnas y sin excluir ningún punto.** La
uniformidad es simplemente promedio/mínimo, también confirmado exactamente.

Esta es la pieza más importante del caso de regresión y quedó completamente
validada.

## 3. Factor de pérdidas — LLF (hoja "FixtureData")

- `FixtureData!H6` ("Total Light Loss Factor") = 0.6424
- `FixtureData!G6` (Watts) = 70
- `FixtureData!B6` (nombre) = "A - HPS 70W Type II"

Es decir: **H6 no corresponde al luminario LED de 34 W evaluado** en la
corrida cacheada (`ATB0_20B_LED_E53_XXXXX_R2.ies`), sino al luminario de
librería "A - HPS 70W Type II", que era el que estaba seleccionado en
Baseline/Upgrade (`FixtureData!C18=C19=3`, el índice 3 de la lista de
luminarios) en el momento en que se guardó el archivo — después, aparentemente,
de haber corrido la simulación con el LED.

Se localizaron LLD/LDD/BF para ese luminario HPS en la hoja "Fixtures",
fila 39 (columnas P/Q/R):

- LLD (Lamp Lumen Depreciation) = 0.73 — `Fixtures!P39`
- LDD (Luminaire Dirt Depreciation) = 0.88 — `Fixtures!Q39`
- BF (Ballast Factor) = 1.0 — `Fixtures!R39`
- Producto = 0.73 × 0.88 × 1.0 = **0.6424** → coincide exactamente con H6.

**Confirmado numéricamente: LLF = LLD × LDD × BF**, pero para el luminario
HPS 70W Type II, no para el LED evaluado.

Se recorrió toda la librería de 38 luminarios de la hoja "Fixtures" (filas
37–73): son todos HPS, MH y LED "genéricos" de catálogo (los LED de librería
usan LLD=0.70, LDD=0.88, BF=1.0 → LLF=0.616). **Ninguno de los 38 tiene 34 W
ni se llama "ATB0..."**. El luminario LED evaluado nunca formó parte de la
librería guardada en el archivo.

**Conclusión: el LLF real del luminario evaluado (LLD, LDD, BF individuales
y su producto) NO es recuperable de este archivo.** Se deja `null` en el
JSON, junto con la fórmula ya verificada (aplicable en general, solo que con
datos de otro luminario).

## 4. Fotometría (hoja "FixtureData")

Estructura de bloque descifrada usando los rangos con nombre `Boffset`
(=1, `FixtureData!B7`), `Hrow` (=10, `!B3`), `Vrow` (=9, `!B2`),
`NumVert`/`NumHoriz` (por luminario) y `FxLib_start` (`!A41`, la celda
etiqueta "First Row"):

Para el luminario cuyo nombre está en la fila `R` (1-based) de `FixtureData`:

| Offset desde R | Contenido |
|---|---|
| +0 | Nombre del luminario |
| +1 | Fabricante |
| +2 | Tipo (LED/HPS/MH) |
| +3 | Modelo |
| +4 | Wattage |
| +5, +6 | (sin uso en este bloque) |
| +7 | NumVert (nº de ángulos verticales = columnas de la matriz de candelas) |
| +8 | NumHoriz (nº de ángulos horizontales = filas de la matriz) |
| +9 | Fila de ángulos VERTICALES (γ), en columnas B en adelante |
| +10 .. +9+NumHoriz | NumHoriz filas: columna A = ángulo HORIZONTAL (φ) de esa fila; columnas B en adelante = candelas |

Tamaño de bloque = 10 + NumHoriz filas; el siguiente luminario empieza justo
después.

Se extrajo el bloque **completo** del primer (y único) luminario presente:
"A - HPS 70W Type II", filas 42–75 de `FixtureData` — NumVert=51,
NumHoriz=24, 51 ángulos verticales, 24 ángulos horizontales, matriz de
24×51 candelas. Todo esto está en el JSON (`fotometria_luminario`).

**Pero éste tampoco es el luminario evaluado.** `FixtureData` guarda
únicamente la fotometría del luminario que estaba cargado/seleccionado al
momento de guardar (no un historial de todos los que alguna vez se
evaluaron), y ese era de nuevo "A - HPS 70W Type II" — el mismo que aparece
en H6. El LED "ATB0_20B_LED_E53_XXXXX_R2.ies" (34 W) no tiene su bloque de
ángulos ni su matriz de candelas en ningún lugar del libro.

**Conclusión: la fotometría del luminario LED evaluado (ángulos, matriz de
candelas) NO es recuperable de este archivo.** Se deja `null` en el JSON,
con la estructura del bloque ya verificada usando el luminario HPS que sí
está presente, por si se quiere usar como fixture de prueba de la
estructura del parser (no del resultado de iluminancia).

## Resumen — ¿tenemos un caso de regresión completo?

**No, está incompleto.** Falta exactamente una pieza, pero es una pieza
central: **la fotometría (ángulos + matriz de candelas) y el LLF real del
luminario LED "ATB0_20B_LED_E53_XXXXX_R2.ies" (34 W)** que efectivamente se
usó para producir la malla de iluminancia cacheada. Ninguno de los dos está
en el archivo — el archivo quedó guardado después de que el usuario cambiara
la selección de luminario en la interfaz a "A - HPS 70W Type II", sobrescribiendo
en `FixtureData`/`Fixtures` los datos del luminario con el que realmente se
corrió el cálculo, sin volver a correr el cálculo con el HPS.

Lo que **sí está completo y verificado numéricamente**:
- La malla de iluminancia esperada (40 puntos) y sus agregados (promedio,
  mínimo, máximo, uniformidad), con el método de agregado confirmado
  (media aritmética simple).
- La fórmula LLF = LLD × LDD × BF, verificada con datos de un luminario
  (aunque no el evaluado).
- La estructura completa del bloque fotométrico en `FixtureData` (offsets,
  filas de ángulos, matriz de candelas), verificada con un bloque real.
- Que la geometría de entrada (Road Geometry) simplemente no fue guardada:
  no es un problema de parseo, las celdas están vacías.

**Lo que falta para completar el caso** no se puede extraer de este
archivo: haría falta el archivo .ies original del luminario LED evaluado
(`ATB0_20B_LED_E53_XXXXX_R2.ies`) y su LLD/LDD/BF reales, o una nueva
corrida del Excel que se guarde inmediatamente después de calcular con ese
luminario, sin cambiar la selección de librería antes de guardar. Mientras
tanto, la malla de iluminancia y los agregados sirven como caso de
regresión válido para el **postproceso** (agregación de la malla), pero NO
para validar el motor de cálculo fotométrico de punta a punta, porque
faltan sus entradas fotométricas reales.
