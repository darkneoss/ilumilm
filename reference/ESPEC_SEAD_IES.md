# Especificación: cómo `ReadISO.bas` interpreta archivos IES (SEAD Street Lighting Tool v1.7.6)

Fuente única: `reference/sead_vba_1.7.6/ReadISO.bas`, subrutina `ReadISOfile` (líneas 229-604). Todas las citas de línea refieren a ese archivo.

## 1. Parseo de la cabecera

El archivo se lee línea por línea con `oFS.ReadLine` en un bucle `Do Until oFS.AtEndOfStream` (líneas 267-468). No hay una fase de "parseo de cabecera" separada del resto: todo se resuelve dentro del mismo bucle, comparando el número de línea (`row`) contra posiciones calculadas dinámicamente.

**Keywords entre corchetes** (líneas 281-291): se buscan por coincidencia de prefijo con `Left(sText, N)`, sólo tres keywords son reconocidos:
- `[MANUFAC]` → `Manufac`
- `[DISTRIBUTION]` → `Distribution`
- `[LUMCAT]` → `LumCat`

Cualquier otro keyword (`[TEST]`, `[ISSUEDATE]`, `[LUMINAIRE]`, `[LAMP]`, `[BALLAST]`, `[LAMPCAT]`, `[OTHER]`, etc.) simplemente se ignora — no rompe el parseo, pero tampoco se guarda ni se muestra en ningún lado.

**Línea TILT** (líneas 295-308): se detecta con `Left(sText, 4) = "TILT"`. En cuanto se encuentra esta línea:
- Si el texto es exactamente `TILT=INCLUDE` se fija `TiltRows = 4`; para cualquier otro valor de TILT (`TILT=NONE`, `TILT=<archivo>`, etc.) `TiltRows` se queda en su valor por defecto de VBA, `0` (nunca se inicializa explícitamente a 0, ver sección de rarezas).
- Se marca `NoKeywords = True` (nombre engañoso: en realidad significa "ya pasamos la sección de keywords/TILT").
- Se calculan, a partir del número de línea actual (`row`) de la línea TILT:
  - `DataRow1 = row + TiltRows + 1` → línea de 10 valores.
  - `DataRow2 = row + TiltRows + 2` → línea de 3 valores.
  - `VertAngleRow = row + TiltRows + 3` → primera línea donde pueden empezar los ángulos verticales.

**Línea de 10 valores** (líneas 313-348, `row = DataRow1`): se tokeniza manualmente buscando espacios con `InStr(sText, " ")`, cortando el string en cada espacio (bucle `Do While Len(sText) > 0`). Espacios duplicados se ignoran (`If sValue <> ""`). Los primeros 7 tokens se asignan por posición:
1. `NumLamps`
2. `Lumens` (`CLng`)
3. `CandelaMult` (`Val`)
4. `NumVertAngles` (`CInt`)
5. `NumHorizAngles` (`CInt`)
6. `PhotoType`
7. `UnitType`

Los tokens 8, 9 y 10 (ancho, largo, alto) están **comentados** (líneas 337-339: `'If VarCount1 = 8 Then Width = sValue` etc.) — se leen y descartan, no se guardan en ninguna variable ni se usan después.

Inmediatamente después de leer esta línea se dimensionan los arrays (línea 344-346):
```
ReDim VertAngles(NumVertAngles)
ReDim HorizAngles(NumHorizAngles)
ReDim CandelaValues(NumHorizAngles, NumVertAngles)
```

**Línea de 3 valores** (líneas 351-374, `row = DataRow2`): mismo tokenizador manual. Los 3 tokens:
1. `BallastFactor` (sin conversión de tipo, se queda como `Variant`)
2. `FutureUse` (ignorado después, sólo se lee)
3. `InputWatts` (`CLng`) — éste es el único valor que sí se usa más adelante.

## 2. ¿Escala las candelas?

Sí, únicamente por el multiplicador (`CandelaMult`, token 3 de la línea de 10 valores). La escritura final del valor de candela es (línea 578):

```vba
tempVal = CandelaValues(i, j)
tempVal = format(tempVal, "#.0###")   ' redondeo a 4 decimales vía formato de texto
tempVal = Val(tempVal)
Sheet2.Cells(...) = tempVal * CandelaMult
```

- **No** hay ninguna escala por `Lumens/1000` en ningún punto del código (búsqueda exhaustiva del archivo: `Lumens` sólo aparece en la asignación de la línea 331 y nunca vuelve a usarse en ningún cálculo — ni siquiera se escribe a las hojas de salida). El valor de lúmenes declarado se lee y se descarta funcionalmente.
- **No** hay escala por `BallastFactor` de la línea de 3 valores en `ReadISO.bas`: `BallastFactor` se lee (línea 367) pero nunca se multiplica contra las candelas dentro de esta subrutina. Sí se escribe luego a `Sheet13.Cells(LastRowSheet13 + 1, 18) = BF` (línea 557), pero esa variable `BF` es el "Ballast Factor" **de los valores por defecto/entrada del usuario para el tipo de luminaria** (`Range(FixtureType & "BF")`, línea 120/205), completamente distinto del `BallastFactor` leído del archivo IES — son dos variables independientes que coinciden sólo de nombre.
- **Fotometría absoluta (`Lumens = -1`)**: no se detecta ni se maneja de forma especial en ningún punto. El valor `-1` se lee como cualquier otro número y, como se explicó, ni siquiera se usa después. El multiplicador (`CandelaMult`) se sigue aplicando exactamente igual, sin importar si `Lumens` es -1 o un número positivo.

## 3. Conversión de unidades (pies → metros)

**No existe ninguna conversión de unidades en el archivo.** `UnitType` (token 7 de la línea de 10) se lee y se guarda en la variable `UnitType`, pero:
- Nunca se compara contra `1` o `2`.
- Nunca se usa en ninguna operación aritmética ni condicional en el resto de la subrutina.
- Los campos de ancho/largo/alto (tokens 8-10, donde normalmente se aplicaría la conversión) están comentados y ni siquiera se capturan (ver sección 1).

En consecuencia, sea `units_type=1` (pies) o `2` (metros), el VBA no hace nada distinto: no convierte dimensiones del luminario (porque ni siquiera las lee) ni ningún otro valor.

## 4. Manejo del rango de ángulos horizontales (simetría)

Esto es lo más importante y la respuesta es tajante: **el VBA no expande ni resuelve ninguna simetría. Copia los ángulos horizontales y verticales exactamente como vienen en el archivo, sea cual sea su rango (0-0, 0-90, 0-180 o 0-360), y copia la matriz de candelas tal cual, con las mismas dimensiones `NumHorizAngles x NumVertAngles` que declara la cabecera.**

Evidencia:
- El array se dimensiona exactamente a `NumHorizAngles` x `NumVertAngles` (línea 346), sin ningún factor de expansión.
- El volcado final a la hoja de cálculo (líneas 571-580) recorre `For i = 1 To NumHorizAngles` / `For j = 1 To NumVertAngles` y escribe literalmente `HorizAngles(i)`, `VertAngles(j)` y `CandelaValues(i,j) * CandelaMult` — no hay ningún código que detecte el máximo ángulo horizontal (90/180/360) ni que refleje/duplique columnas para "completar" el círculo.
- No existe ninguna otra rutina en `ReadISO.bas` que resuelva la simetría "al consultar" (no hay funciones de interpolación/lookup en este archivo; la subrutina sólo lee y escribe a hojas de Excel).

Conclusión: la resolución de simetría (si existe en algún lugar del sistema SEAD) tendría que vivir en otro módulo que consuma los datos ya volcados en `Sheet2` (candela data cruda, sin expandir) — no en `ReadISO.bas`. Este archivo entrega los datos "en bruto" tal como el fabricante los declaró, sin normalizarlos a un rango canónico de 0-360°.

## 5. Origen de los watts del luminario

`InputWatts` viene exclusivamente del tercer token de la **línea de 3 valores** del archivo IES (línea 370: `InputWatts = CLng(sValue)`, dentro del bloque `If VarCount2 = 3`). No se lee de ningún keyword `[LAMP]`/`[BALLAST]`/`[WATTS]`, ni se solicita al usuario en ningún punto de `ReadISOfile`, `ReadSingleISOfile` ni `ReadMultipleISOFile`.

Ese valor se propaga:
- A `Sheet13.Cells(LastRowSheet13 + 1, 10) = InputWatts` (línea 547, columna de watts en la hoja "Fixtures").
- A `Sheet2.Cells(StartRow + 4, 2) = InputWatts` (línea 564, en el bloque de "intro information" de la ficha del luminario).

No se pudo inspeccionar `FixtureCalcs.bas` en este encargo (no forma parte del archivo revisado); si ese módulo usa watts en cálculos de energía, previsiblemente los toma de la columna 10 de `Sheet13` (donde `ReadISO.bas` los deposita) y no vuelve a leer el archivo IES.

## 6. Soporte de `TILT=INCLUDE`

Sí, se detecta (línea 296: `If sText = "TILT=INCLUDE" Then TiltRows = 4`), y el único efecto es desplazar 4 líneas adicionales el punto donde empiezan `DataRow1`/`DataRow2`/`VertAngleRow`, asumiendo que esas 4 líneas contienen los datos de tilt (lamp-to-luminaire geometry, número de pares ángulo/factor y sus valores) en el formato estándar LM-63.

**El contenido de esas 4 líneas de tilt nunca se lee ni se parsea** — el código sólo las "salta" contando líneas; no valida que sean 4 líneas o que el número de pares de tilt declarado coincida con las líneas reales. Si el bloque TILT=INCLUDE real ocupa un número de líneas distinto de 4 (por ejemplo si tiene muchos pares ángulo-factor que no caben en una sola línea de continuación), el desplazamiento fijo de 4 desalinea todo el resto del parseo silenciosamente.

## 7. Lectura de arrays de ángulos y candelas: ¿por línea o como flujo de tokens?

Como **flujo de tokens continuo**, no por línea fija. El código no asume que cada array quepa en una única línea de texto:

- Para ángulos verticales (líneas 379-404): mientras `row >= VertAngleRow` y `VertAngleFilled = False`, cada línea leída se tokeniza y cada token numérico se añade a `VertAngles()` incrementando `AngleCount`, sin importar cuántas líneas físicas ocupe — sigue consumiendo líneas sucesivas del archivo hasta acumular exactamente `NumVertAngles` valores, momento en el que fija `VertAngleFilled = True` y calcula `HorizAngleRow = row + 1` (siguiente línea).
- Para ángulos horizontales (líneas 407-434): mismo patrón, activado una vez `VertAngleFilled = True`, hasta completar `NumHorizAngles` valores; entonces fija `HorizAngleFilled = True` y `CandelaRow = row + 1`.
- Para las candelas (líneas 436-465): una vez `HorizAngleFilled = True`, cada línea restante del archivo se tokeniza y cada valor numérico (`IsNumeric(sValue)`) se acomoda en `CandelaValues(i, j)` incrementando `j` hasta `NumVertAngles`, momento en que resetea `j = 1` y avanza `i = i + 1` (líneas 453-458) — es decir, reconstruye la matriz fila-por-fila (una fila = un ángulo horizontal, con `NumVertAngles` candelas) sin que las líneas físicas del archivo tengan que alinearse con las filas lógicas de la matriz.

Esto significa que el parser **sí tolera correctamente** que los arrays de ángulos y de candelas estén partidos en líneas arbitrarias del archivo (como ocurre en los 6 `.ies` de `assets/`), porque trabaja con un flujo de tokens acumulativo y sólo corta cuando el contador (`AngleCount`, o `i`/`j`) llega al total esperado — nunca por fin de línea.

## Ambigüedades y rarezas (comportamiento real vs. estándar LM-63)

1. **`Lumens` (lúmenes por lámpara) se lee y se descarta.** Nunca se usa para escalar candelas, ni se reporta en ninguna hoja de salida. En fotometría absoluta (`Lumens = -1`) el estándar LM-63 indica que el multiplicador de candela ya contiene toda la información necesaria y que -1 es sólo una bandera informativa — en eso el VBA "acierta por accidente" (no necesita hacer nada especial porque nunca usó `Lumens`), pero no es por diseño consciente: simplemente ignora el campo siempre, sea -1 o cualquier otro valor.

2. **`BallastFactor` del archivo IES no se aplica a las candelas.** Según LM-63, el valor final debería ser candela_bruta × multiplicador × ballast_factor × (lumens_por_lampara/1000 en fotometría relativa). El VBA sólo aplica el multiplicador; ignora tanto el ballast factor del archivo como el factor lumens/1000. Esto produce valores de candela sistemáticamente distintos del estándar cuando el ballast factor del archivo ≠ 1 o cuando la fotometría es relativa con lúmenes/lámpara ≠ 1000.

3. **`UnitType` (pies/metros) se lee y nunca se usa.** No hay conversión de unidades en ningún punto — ni de las dimensiones del luminario (que tampoco se capturan, ver punto 4) ni de ninguna otra magnitud. Un archivo en pies (`units_type=1`) y uno en metros (`units_type=2`) se tratan exactamente igual.

4. **Ancho/largo/alto del luminario (tokens 8-10 de la línea de 10) se leen y se tiran** — están explícitamente comentados en el código (líneas 337-339), simplemente avanzan el cursor de tokenización pero no se guardan.

5. **`TiltRows` no se inicializa a 0 explícitamente antes del `If`.** Como es una variable `Integer` declarada al tope de la subrutina (línea 236: `Dim TiltRows As Integer`), VBA la inicializa por defecto a 0 al declarar, así que el comportamiento accidental coincide con el correcto para `TILT=NONE`/`TILT=<nombre_archivo>` — pero es frágil: si esta subrutina fuera invocada más de una vez reutilizando estado (no es el caso aquí, es de ámbito local), o si alguien reordena el código, dependen implícitamente del valor por defecto del lenguaje en vez de asignarlo.

6. **El bloque `TILT=INCLUDE` se salta con un desplazamiento fijo de 4 líneas**, sin parsear ni validar su contenido real (ángulo de inclinación de la luminaria, número de pares, y los pares ángulo/multiplicador). Si el bloque real no ocupa exactamente 4 líneas, el resto de la lectura (línea de 10 valores, línea de 3, ángulos, candelas) queda desalineado silenciosamente — no hay ninguna validación cruzada (p. ej. comprobar que el conteo de candelas leídas cuadra con `NumHorizAngles * NumVertAngles`) que hubiera detectado el error.

7. **No hay resolución de simetría de ninguna clase.** Un archivo con `horiz_angles` que llega sólo a 90° o 180° (simetría de cuadrante o bilateral) se almacena tal cual, con esa cantidad reducida de columnas — cualquier consumo posterior de estos datos (cálculos de iluminancia, gráficos polares) tendría que saber reconstruir la simetría por su cuenta; `ReadISO.bas` no ofrece ni información explícita sobre qué tipo de simetría aplica (más allá de que se puede inferir del último valor de `HorizAngles`).

8. **Límite de 256 columnas de Excel** (líneas 510-527): si `NumVertAngles >= 256` el código verifica si la versión de Excel/modo de compatibilidad permite más columnas; si no, marca `TooManyColumns = True` y **descarta todo el archivo sin escribir nada**, incluyendo `InputWatts`, `Manufac`, etc. — esto no es un límite del formato IES sino una limitación de la hoja de cálculo destino, pero afecta directamente qué archivos LM-63 legítimos logra procesar la herramienta (p. ej. `V1050UN2M50.ies` con `NumVertAngles=361` cae en este chequeo).

9. **Tokenización manual con `InStr`/`Left`/`Right` en vez de `Split`.** Funciona, pero es ineficiente y fue reescrita idéntica cuatro veces (línea de 10 valores, línea de 3, ángulos verticales, ángulos horizontales, candelas) — puro código duplicado, sin relevancia funcional pero notable como "rareza" de calidad.

10. **`PhotoType` (tipo fotométrico: 1=C, 2=B, 3=A) se lee y nunca se usa** para nada — ni para interpretar el significado de los ángulos horizontales/verticales, ni para ningún cálculo. Esto es coherente con el punto 3: el archivo trata todos los tipos fotométricos de la misma manera genérica (ángulo vertical × ángulo horizontal × candela), sin adaptar la interpretación geométrica al sistema C, B o A.

## Contraste con los archivos reales (`assets/*.ies`)

Extraído parseando los 6 archivos con un script Python de un solo uso (no incluido en el repo), replicando exactamente el algoritmo de `ReadISO.bas` (localizar línea TILT, `TiltRows=4` si `TILT=INCLUDE`, línea de 10 valores, línea de 3 valores, luego flujo de tokens para ángulos).

| Archivo | n_vert | n_horiz | rango vertical (°) | rango horizontal (°) | units_type | lúmenes declarados | multiplicador | ballast factor | watts declarados |
|---|---|---|---|---|---|---|---|---|---|
| V1050UN2M50.ies | 361 | 73 | 0 – 180 | 0 – 360 | 2 (metros) | 6597.61 | 0.965 | 1.000 | 49.94 |
| V1070UN2M50.ies | 361 | 73 | 0 – 180 | 0 – 360 | 2 (metros) | 9456.8 | 1 | 1.000 | 71.9 |
| V2100UN2M50.ies | 73 | 19 | 0 – 180 | 0 – 180 | 1 (pies) | -1 (fotometría absoluta) | 1 | 1 | 95.986 |
| V2130UN2M50.ies | 73 | 19 | 0 – 180 | 0 – 180 | 1 (pies) | -1 (fotometría absoluta) | 1 | 1 | 124.484 |
| V3160UN2M50.ies | 361 | 73 | 0 – 180 | 0 – 360 | 2 (metros) | 22404 | 1 | 1.000 | 164.6 |
| V3200UN3M50.ies | 361 | 73 | 0 – 180 | 0 – 360 | 3 | 26115 | 1 | 1.000 | 199.3 |

Notas sobre la tabla:
- Los archivos con 73 ángulos horizontales cubriendo 0-360° (`V1050`, `V1070`, `V3160`, `V3200`) no tienen ninguna simetría declarada (rango completo) — el VBA los carga tal cual, sin cambios.
- `V2100` y `V2130` cubren horizontalmente sólo 0-180° con 19 puntos — esto es simetría bilateral (plano de simetría en 0°/180°); dado el punto 4/7 de este documento, `ReadISO.bas` los deja exactamente así, con 19 columnas, sin expandir a 360°.
- `V2100`/`V2130` declaran `units_type=1` (pies) y lúmenes `-1` (fotometría absoluta) — según lo documentado en las secciones 2 y 3, ninguno de estos dos valores provoca ninguna rama de código distinta en `ReadISO.bas`: se leen, se guardan en variables que no vuelven a usarse (excepto que si `Lumens` hubiera sido usado, que no lo es) y el multiplicador de candela (1 en ambos casos) se aplica igual que en los demás archivos.
- `V3200UN3M50.ies` reporta `units_type=3`, un valor no estándar en LM-63 (los válidos son 1=pies, 2=metros) — de nuevo irrelevante para el VBA porque nunca inspecciona este campo.
