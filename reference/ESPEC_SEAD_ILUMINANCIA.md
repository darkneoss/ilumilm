# Especificación del algoritmo de Iluminancia (método IES) — SEAD Street Lighting Tool v1.7.6

Fuente: VBA extraído en `reference/sead_vba/`. Todas las citas son `archivo.bas:línea`.
Alcance: SOLO método IES, SOLO iluminancia (se ignora luminancia, CIE, financiero, gráficas, formularios, traducción).

---

## 1. Resumen del flujo (`Sub finalMatrices`, `IlluminanceAndLuminance.bas:7-220` rama IES `:131-219`)

1. Leer geometría de la vía (carriles, ancho de carril, mediana, altura de montaje, separación de postes, retranqueo de poste, longitud de brazo, arreglo de postes) desde `Road Geometry` (`:32-56`).
2. `LLF = Sheets("FixtureData").Range("H6").Value` (`:67`) — ver sección 4.
3. `ngp = TotalGridLength(IES) / GridSpace(IES)` (`:75`).
4. `outputXY = makeGrid(...)` → matriz de coordenadas X (a lo largo de la vía) e Y (transversal) de TODOS los puntos de rejilla posibles (`:79`).
5. `fixtureX(), fixtureY() = FixturePosition(...)` → posiciones de todos los postes candidatos, calculadas sobre una rejilla que cubre `gridlength = TotalGridLength(IES)` (`:85-87`).
6. Para cada poste `k` (`For k = LBound(fixtureX) To UBound(fixtureX)`, `:144-174`):
   a. `phi = anglePhi(...)`, `gammaArray = angleGamma(...)` (ángulos sin tilt; los "WithTilt" se calculan pero solo se usan para pasar a `LintensityMatrix`, ver Ambigüedad A1).
   b. `larray = LintensityMatrix(...)` → intensidad luminosa I (cd) interpolada de la tabla fotométrica, evaluada en `phiArrayForITable`/`gammaArrayForITable` (los ángulos CON tilt, que con tilt=0 son iguales a los sin tilt) (`:153`).
   c. `tempArray1 = Illum(larray, gammaArray, LLF, FixtureHeight)` → iluminancia que aporta el poste `k` en cada punto de la ventana de rejilla (`:168`). Nota: aquí se usa el `gammaArray` SIN tilt (no el `gammaArrayForITable`).
   d. Se guarda `illuminanceFixture(k) = tempArray1`.
7. Se suma la contribución de TODOS los postes punto a punto: `LsumArray(j,k) = Σ_i illuminanceFixture(i)(j,k)` (`:190-203`). No hay ningún filtro de distancia ni de "5H" para IES (ese filtro solo existe en la rama CIE, `:377-381`) — ver Ambigüedad A2.
8. `LsumArray` se vuelca a la hoja `Illuminance Calcs` a partir de `B13` (`:206`).

---

## 2. Rejilla de medición (`MakeMeasurementGrid.bas`)

### 2.1 Separación de rejilla — `GridSpace` (`:90-108`)
```
GridSpace(IES, polespacing) = min(polespacing / 10, 5)
```
Confirmado, `:92-96`.

### 2.2 Longitud total de rejilla — `TotalGridLength` (`:109-115`)
```
TotalGridLength(IES, polespacing) = 4 * polespacing
```
Confirmado, `:111`.

### 2.3 Número de puntos de rejilla
```
ngp = TotalGridLength / GridSpace          (IlluminanceAndLuminance.bas:75)
```
`ngp` es el índice más alto (0-based) del arreglo `Xvalues`: `ReDim Xvalues(numberOfGridPoints)` con `numberOfGridPoints = ngp` (`MakeMeasurementGrid.bas:28`), o sea `Xvalues` tiene `ngp+1` elementos, índices `0..ngp`.

### 2.4 Valores X (`makeGrid`, `:33-41`)
```
Xvalues(0) = GridSpacing / 2
Xvalues(i) = Xvalues(i-1) + GridSpacing     para i = 1..ngp
```
(Nota: aunque el código distingue el caso `calculationmethod = "CIE"` con un `If`, ambas ramas hacen exactamente lo mismo, `Xvalues(0) = GridSpacing/2`; el `If` es inocuo — ver Ambigüedad A3.)

### 2.5 Número de valores Y (`:12-23`)
```
NumberOfGridsPerLane = 2                    (IES, :13)
NumberOfYvalues = 2 * NumberOfLanes         (el If Mod 2 en :19-23 no cambia nada; ambas ramas son idénticas)
```
Arreglo `Yvalues` con índices `0 .. NumberOfYvalues-1`.

### 2.6 Valores Y (`:54-63`, rama IES)
```
Yvalues(0) = lanewidth / 4
Para i = 1 .. NumberOfYvalues-1:
    Yvalues(i) = Yvalues(i-1) + lanewidth/2
    Si Yvalues(i) >= medianYvalue Y flagformedian:
        Yvalues(i) += MedianLength
        flagformedian = False        ' el ajuste por mediana se aplica UNA SOLA VEZ, en el primer cruce
```
donde:
```
Si NumberOfLanes es par:  medianYvalue = (NumberOfLanes/2) * lanewidth ; flagformedian = True
Si NumberOfLanes es impar: medianYvalue = 0 ; flagformedian = False (nunca se suma la mediana a Y)
```
(`:44-50`). Es decir, con número impar de carriles la mediana NUNCA se inserta en la coordenada Y de la rejilla — ver Ambigüedad A4.

`makeGrid` devuelve `ArrayXY = {Xvalues, Yvalues}` (`:79-84`), y esto es lo que en `finalMatrices` se llama `outputXY` (con `outputXY(0)` = arreglo de X, `outputXY(1)` = arreglo de Y).

---

## 3. Posiciones de postes — `FixturePosition` (`FixturePositions.bas:22-89`)

```
Si PoleConfiguration = "Single-side":
    numberoffixtures = gridlength / polespacing
Sino:
    numberoffixtures = (gridlength / polespacing) * 2

FPArrayX, FPArrayY dimensionados 0 .. CInt(numberoffixtures)+1   (:37-38)
```

**Single-side** (`:41-45`, la única configuración que el usuario ya validó y la única requerida para el objetivo):
```
Para i = 0 .. UBound(FPArrayX):
    FPArrayX(i) = i * polespacing
    FPArrayY(i) = 0 - polesetback + ArmLength
```
Y=0 es el borde de vía más cercano al poste (confirmado). Todas las demás configuraciones (`Opposite`, `Median mounted`, `Staggered`) alternan lado en índices pares/impares — no se detalla más porque el encargo es sobre single-side/IES genérico, pero el código de `finalMatrices` es agnóstico a la configuración: siempre itera `For k = LBound(fixtureX) To UBound(fixtureX)`, es decir **todos** los postes generados por `FixturePosition` participan en la suma (sección 1.7), sin importar si están dentro o fuera de la ventana de evaluación — los que caen lejos simplemente producen valores de iluminancia muy pequeños (nunca se descartan explícitamente en la rama IES).

`gridlength` pasado a `FixturePosition` es `TotalGridLength(IES) = 4*polespacing` (`IlluminanceAndLuminance.bas:85-87`), o sea para Single-side hay `4*polespacing/polespacing + 2 = 6` postes (índices 0..5) en X = 0, 35, 70, 105, 140, 175 (con polespacing=35).

---

## 4. LLF (Light Loss Factor)

`LLF = Sheets("FixtureData").Range("H6").Value` (`IlluminanceAndLuminance.bas:67`, reafirmado en `:351` y `:380` en la rama CIE).

**No existe ninguna fórmula VBA que calcule H6** — es una celda de la hoja de cálculo `FixtureData` (fuera del alcance del VBA extraído). Lo único que el VBA muestra sobre sus componentes es en `ReadISO.bas`:
- `LLD`, `LDD`, `BF` se leen de la librería de luminarios / hoja `Sheet21` (`ReadISO.bas:13-15, 118-130, 203-205`) y se escriben a `Sheet13` (hoja "Fixtures"), columnas 16, 17, 18 respectivamente (`:555-557`) — es decir **Lamp Lumen Depreciation, Luminaire Dirt Depreciation, Ballast Factor**.
- No hay una línea VBA que multiplique `LLD*LDD*BF` y lo deposite en `FixtureData!H6`; eso debe ser una fórmula de hoja de cálculo (probablemente `=LLD*LDD*BF` referenciando la fila de `Sheet13` del luminario activo), que no está disponible en los `.bas`. 
- **Recomendación para el reimplemento en Python**: tratar `LLF = LLD * LDD * BF` como hipótesis razonable (es el producto estándar en fotometría de vías) pero flagearlo como supuesto no verificado en el código fuente, y exponer `LLF` como parámetro de entrada directo (igual que hace el VBA, que simplemente lee un valor ya calculado).
- El multiplicador de candela del archivo IES (`CandelaMult`, tercer valor de la línea de datos del IES) **ya se aplica una sola vez, en el momento de importar el archivo**, directamente sobre la tabla de candelas que queda grabada en `FixtureData` (`ReadISO.bas:578`: `Sheet2.Cells(...) = tempVal * CandelaMult`). Por lo tanto, en tiempo de cálculo (`LintensityCalc`) la tabla ya está en candelas absolutas y **no se vuelve a aplicar ningún multiplicador, factor de lúmenes ni de balastro** — solo `LLF` se aplica, y únicamente dentro de `Illum` (sección 6).

---

## 5. Cálculo de ángulos (`AngleCalculations.bas`)

### 5.1 Ventana de evaluación (idéntica en `angleGamma`, `anglePhi`, `angleBeta`, `LintensityMatrix` — repetida literalmente en cada función)
```
Si IES:
    iStart = WorksheetFunction.Match(polespacing,     outputX, True)          ' match_type=1 (aproximado, ascendente)
    iEnd   = WorksheetFunction.Match(2 * polespacing,  outputX, True) - 1
```
(`AngleCalculations.bas:16-17` y repetido en `:61-62, :138-139, :219-221, :333-334`, y en `LuminanceIntensity.bas:25-26`).

**Semántica de `Match(valor, array, 1)` cuando `array` es un arreglo VBA `Variant` 0-based**: Excel `WorksheetFunction.Match` con `match_type=1` requiere el arreglo ascendente y devuelve la **posición relativa dentro del arreglo, 1-based**, del mayor elemento `<= valor` (o el índice del último elemento si todos son menores). Es decir, si el elemento que cumple la condición está en la posición VBA `outputX(n)` (con `n` 0-based), `Match` devuelve `n+1`.

El código VBA **reutiliza ese resultado directamente como índice del arreglo 0-based** (`For i = iStart To iEnd ... outputX(i) ...`), **sin restarle 1**. Esto produce un corrimiento de +1 posición respecto de lo que uno esperaría de una búsqueda "aproximada hacia abajo": en la práctica, `iStart` (usado como índice VBA) apunta al **primer punto de rejilla con `outputX(i) > polespacing`**, no al que está justo debajo. Ver el ejemplo numérico abajo y la Ambigüedad A5.

### 5.2 Ejemplo numérico pedido
`polespacing = 35`, `GridSpacing = 3.5` (`= min(35/10,5) = 3.5`), por lo tanto:
```
outputX(i) = 1.75 + 3.5*i     para i = 0..ngp   (ngp = 4*35/3.5 = 40)
outputX(0)=1.75, outputX(1)=5.25, outputX(2)=8.75, ... outputX(9)=33.25, outputX(10)=36.75, ...
                                                   ... outputX(19)=68.25, outputX(20)=71.75, ...
```
- `Match(35, outputX, 1)`: el mayor valor `<=35` es `outputX(9)=33.25` (posición 1-based = 10). ⇒ Excel devuelve **10**.
  `iStart = 10` (usado directamente como índice VBA) ⇒ `outputX(iStart) = outputX(10) = 36.75`.
- `Match(70, outputX, 1)`: el mayor valor `<=70` es `outputX(19)=68.25` (posición 1-based = 20). ⇒ Excel devuelve **20**.
  `iEnd = 20 - 1 = 19` ⇒ `outputX(iEnd) = outputX(19) = 68.25`.

**Ventana resultante: `i = 10..19`, es decir los puntos X = 36.75, 40.25, 43.75, 47.25, 50.75, 54.25, 57.75, 61.25, 64.75, 68.25` (10 puntos).**

Nótese que:
- El punto más cercano "por debajo" de `polespacing` (33.25) **queda fuera** de la ventana.
- El límite superior nominal `2*polespacing=70` tampoco se alcanza; el último punto es 68.25.
- La ventana cubre exactamente **un `polespacing` de longitud** (10 puntos × 3.5 = 35), pero desplazada hacia adelante en `GridSpacing` respecto de lo que una lectura ingenua de "Match aproximado" sugeriría. Esto es consistente en todos los casos donde `polespacing` no es múltiplo exacto de `GridSpacing` (lo cual es casi siempre, dado `GridSpacing = min(polespacing/10,5)`, salvo cuando `polespacing<=50` y es múltiplo de `polespacing/10`, que es trivialmente siempre cierto en ese rango... en realidad para `polespacing<=50`, `GridSpacing=polespacing/10` exactamente, por lo que `polespacing` SÍ es múltiplo exacto de `GridSpacing` en ese caso, y el off-by-one se comporta distinto — ver Ambigüedad A5b).

### 5.3 `numberOfY` para el bucle interno
```
numberOfY = UBound(outputY)     (AngleCalculations.bas:30, etc.)
```
o sea el bucle interno recorre `j = 0 .. UBound(outputY)` (TODOS los valores Y de la rejilla, sin ventana).

### 5.4 `angleGamma` (`:5-48`)
```
distY = fixtureY                                    ' Y del poste (constante en todo el bucle)
dist  = Distance(fixtureX, distY, outputX(i), outputY(j))
                                                     ' = sqrt((fixtureX-outputX(i))² + (fixtureY-outputY(j))²)
Si dist ≠ 0:  gamma(i,j) = atan(dist / FixtureHeight) * 180/π
Si dist = 0:  gamma(i,j) = 0
```
Confirmado, `:38-44`. `dist` es la distancia horizontal real (2D) entre poste y punto de rejilla, no solo `|X|`.

### 5.5 `anglePhi`, caso NO "Median mounted" (`:181-191`, el caso relevante para Single-side)
```
distY = fixtureY
Si (distY - outputY(j)) ≠ 0:
    phi(i,j) = atan( |fixtureX - outputX(i)| / |distY - outputY(j)| ) * 180/π      ' en [0,90)
Sino:
    phi(i,j) = 90
```
Confirmado — coincide con lo que ya tenías, en rango `[0,90]`, con `90` cuando `outputY(j) == fixtureY` (dY=0).

### 5.6 `Illum` — fórmula final de iluminancia (`IlluminanceAndLuminance.bas:442-456`)
```
E(i,j) = larray(i,j) * cos(gammaArray(i,j) * π/180)^3 * LLF / FixtureHeight^2
```
Confirmado exactamente. Usa el `gammaArray` **sin tilt** (parámetro `gammaArray`, no `gammaArrayForITable`), a pesar de que `larray` fue interpolado usando los ángulos **con tilt** (`phiArrayForITable`, `gammaArrayForITable`) — ver Ambigüedad A1.

---

## 6. Intensidad luminosa `larray` — interpolación (`LuminanceIntensity.bas`)

### 6.1 Función invocadora: `LintensityMatrix` (`:6-87`)
- Ubica la tabla fotométrica dentro de `FixtureData`:
  ```
  firstRow = FixtureData!B16 + 9      ' fila donde están los ángulos verticales (encabezado)
  lastRow  = FixtureData!B17
  lastCol  = última columna con datos en la fila firstRow
  ```
  (`:52-54`)
- `tablePhi1()` = valores de la columna A, filas `firstRow+1 .. lastRow` (`:58,64-74`) → a pesar del nombre "Phi", por el comentario de `:52` ("Row with angles (vert angles)") esta columna en realidad corresponde a los **ángulos horizontales/C-plane** del archivo IES (ver Ambigüedad A6 sobre el nombre confuso).
- `tableGamma1()` = valores de la fila `firstRow`, columnas `2 .. lastCol` (`:59,68-70`) → corresponde a los **ángulos verticales (gamma)** del archivo IES.
- `tableArray()` = bloque de candelas, filas `firstRow+1..lastRow` × columnas `2..lastCol` (`:76`), ya multiplicado por `CandelaMult` en el momento de la importación (sección 4).
- Para cada punto de la ventana `(i,j)` (misma ventana `iStart..iEnd` de la sección 5.1, con `j=0..numberOfY`):
  ```
  larray(i,j) = LintensityCalc(gridPhi(i,j), gridGamma(i,j), tablePhi1, tableGamma1, tableArray)
  ```
  donde `gridPhi`/`gridGamma` son los `phiArrayForITable`/`gammaArrayForITable` (ángulos **con tilt**, que con tilt=0 son numéricamente iguales a `phi`/`gammaArray` sin tilt) (`IlluminanceAndLuminance.bas:153`).

### 6.2 `LintensityCalc` — interpolación bilineal con vecino más cercano por abajo (`:89-185`)

**Paso 1 — localizar Phi (eje "horizontal", filas de la tabla):**
```
Phi1pos = Application.Match(gridPhi, tablePhi, 1)     ' match aproximado (mayor valor <= gridPhi), 1-based
Phi1    = tablePhi(Phi1pos)                            ' NOTA: aquí Phi1pos SÍ se usa para indexar
                                                        ' directamente tablePhi() que también es 1-based
                                                        ' (ReDim tablePhi1(phiN) con phiN=UBound, típico 1..N)
Si Phi1pos = 1:                Phi0pos = Phi1pos + 1
ElseIf Phi1pos = UBound(tablePhi): Phi0pos = Phi1pos - 1
Sino:                            Phi0pos = Phi1pos - 1
Phi0 = tablePhi(Phi0pos)

Si Phi1 = Phi0: K1 = 0
Sino:           K1 = (gridPhi - Phi1) / (Phi0 - Phi1)
K2 = 1 - K1
```
(`:102-125`)

**Paso 2 — localizar Gamma (eje "vertical", columnas de la tabla), análogo:**
```
Gamma1pos = Application.Match(gridGamma, tableGamma, 1)
Gamma1    = tableGamma(Gamma1pos)
Si Gamma1pos = 1:                   Gamma0pos = Gamma1pos + 1
ElseIf Gamma1pos = UBound(tableGamma): Gamma0pos = Gamma1pos - 1
Sino:                                 Gamma0pos = Gamma1pos - 1
Gamma0 = tableGamma(Gamma0pos)

Si Gamma1 = Gamma0: kGamma1 = 0
Sino:               kGamma1 = (gridGamma - Gamma1) / (Gamma0 - Gamma1)
kGamma2 = 1 - kGamma1
```
(`:136-157`)

**Paso 3 — interpolación bilineal (llamada "cuadrática" en los comentarios, pero es bilineal con 2×2 puntos, no 3×3):**
```
IPhi0Gamma0 = tableArray(Phi0pos, Gamma0pos)
IPhi0Gamma1 = tableArray(Phi0pos, Gamma1pos)
IPhi1Gamma0 = tableArray(Phi1pos, Gamma0pos)
IPhi1Gamma1 = tableArray(Phi1pos, Gamma1pos)

IgridPhiGamma0 = K1*IPhi0Gamma0 + K2*IPhi1Gamma0
IgridPhiGamma1 = K1*IPhi0Gamma1 + K2*IPhi1Gamma1

I(gridPhi,gridGamma) = kGamma1*IgridPhiGamma0 + kGamma2*IgridPhiGamma1
```
(`:169-184`)

Es decir: interpolación lineal en Phi para cada uno de los dos Gamma vecinos, y luego interpolación lineal en Gamma entre esos dos resultados — bilineal estándar, con la salvedad de la elección de vecinos descrita abajo.

### 6.3 Manejo de valores fuera de rango / vecino "hacia abajo" únicamente
- `Match(x, tabla, 1)` en Excel, con la tabla ascendente, se satura en los extremos: si `x` es menor que el primer elemento de la tabla, `Match` genera **error #N/A** (no hay elemento `<= x`); si `x` es mayor que el último, `Match` devuelve la posición del último elemento (no hay extrapolación).
- Cuando `Phi1pos = 1` (el propio match ya es el primer elemento de la tabla, o sea `gridPhi < tablePhi(2)`), el código NO usa el vecino "hacia abajo" (que no existe) sino que fuerza `Phi0pos = 2` (`Phi1pos+1`), es decir usa los dos primeros elementos de la tabla — sigue siendo interpolación entre los dos vecinos más próximos, nunca extrapolación fuera de la tabla.
- Cuando `Phi1pos = UBound(tablePhi)` (`gridPhi` es mayor o igual que el último ángulo de la tabla — incluido el caso en que Phi1pos sea la última posición porque Match se saturó), `Phi0pos = Phi1pos - 1`: usa los dos últimos elementos. Es decir, si `gridPhi` cae por encima del rango de la tabla (p. ej. tabla horizontal 0–90° pero `gridPhi` calculado en 95°), **no hay clamping explícito de `gridPhi`**: `Match` devuelve la última posición y el código interpola entre los dos últimos puntos, lo que en la práctica **extrapola linealmente** en vez de recortar (clamp) al valor de borde. Esto aplica igual al eje Gamma.
- Igualmente, si `gridPhi` fuera MENOR que el primer valor de la tabla (ej. tabla que arranca en 0 y `gridPhi` calculado como negativo, lo cual no debería pasar porque `anglePhi`/`angleGamma` siempre producen valores en `[0,180]`), `Match` fallaría con error `#N/A` que se propagaría como error VBA — no hay manejo defensivo. En la práctica, dado que `anglePhi`/`angleGamma` en el código VBA siempre devuelven ángulos `>=0`, y las tablas IES para calzadas empiezan en 0°, esto no debería activarse en operación normal, pero **no hay clamp explícito en el código**, es una responsabilidad del handler de reimplementación.
- **No se aplica ningún factor adicional en `LintensityCalc`** — la tabla ya contiene el resultado final en candelas (candela × CandelaMult, ver sección 4); no hay multiplicador de lúmenes ni de balastro en este punto.

### 6.4 Rango angular horizontal del archivo IES (0-90 / 0-180 / 0-360)
El código **no distingue explícitamente** el tipo de simetría horizontal (`NumHorizAngles`/tipo de fotometría) al momento de interpolar: simplemente usa la tabla `tablePhi1()` tal cual quedó importada por `ReadISO.bas` (que sí lee `NumHorizAngles` y todos los ángulos horizontales reales del archivo, sea cual sea su rango: 0-90, 0-180 o 0-360). `LintensityCalc` no hace ningún "mapeo" o "espejo" de `gridPhi` (que siempre viene en `[0,90]` desde `anglePhi`, sección 5.5) hacia el rango completo de la tabla si ésta fuera, p.ej., 0-360: simplemente busca `gridPhi` (en `[0,90]`) directamente dentro de `tablePhi1`. **Esto es correcto solo si la tabla fotométrica es de tipo 0-90° (con simetría cuádruple, típico de luminarias de vialidad simétricas), y produce resultados incorrectos/silenciosos si el archivo IES importado tiene rango horizontal 0-180 o 0-360** (el código buscaría `gridPhi<=90` dentro de una tabla que va hasta 180 o 360, encontrando solo el primer cuadrante) — ver Ambigüedad A7.

---

## 7. Agregados de reporte (promedio, mínimo, máximo, uniformidad)

**No existe código VBA para el método IES que calcule promedio/mínimo/máximo/uniformidad de iluminancia.** `finalMatrices` en la rama IES únicamente escribe la matriz `LsumArray` en la hoja `Illuminance Calcs!B13` (`IlluminanceAndLuminance.bas:206`) y dos encabezados de columnas (`:213-219`). No hay ningún `WorksheetFunction.Average/Min/Max` para IES en los `.bas` disponibles — esos cálculos deben vivir como **fórmulas de hoja de cálculo** en `Illuminance Calcs` / `MResults` / `Illuminance` (fuera del alcance del VBA extraído), no fueron encontradas en ningún `.bas` del repo (`grep` sin resultados para "Average"/"uniformity" asociado a iluminancia IES).

Como referencia del patrón que el propio autor usa para el caso análogo (**luminancia, rama CIE**, `IlluminanceAndLuminance.bas:319-326`), que si es visible en VBA:
```
promedio    = WorksheetFunction.Average(RsumArray)          ' promedio aritmético simple de TODOS los puntos de la matriz
minimo      = WorksheetFunction.Min(RsumArray)
maximo      = WorksheetFunction.Max(RsumArray)
Si minimo = 0:  uniformidad = "" (blanco)
Sino:           uniformidad = minimo / promedio             ' Min/Avg (comentario en :304 dice "changed from max/min ***")
```
**Recomendación**: para la reimplementación en Python del método IES, replicar este mismo patrón (promedio simple sobre todos los puntos de `LsumArray` dentro de la ventana `iStart..iEnd` × todos los `j`; uniformidad = min/avg) por ser el único patrón documentado en el propio código fuente, dejándolo explícitamente marcado como **inferido por analogía, no confirmado para IES** (Ambigüedad A8). Nótese además que en `:304` hay un comentario explícito del autor original indicando que la fórmula de uniformidad fue *cambiada* de `max/min` a `min/max` en algún momento — indicio de que esta cifra fue fuente de errores/ambigüedad incluso para los desarrolladores originales.

---

## 8. Pseudocódigo consolidado (método IES, iluminancia, un solo poste lateral "Single-side")

```
GridSpacing   = min(polespacing/10, 5)
GridLength    = 4 * polespacing
ngp           = GridLength / GridSpacing            # entero exacto si polespacing/GridSpacing lo es

# Rejilla
X[i] = GridSpacing/2 + i*GridSpacing                for i in 0..ngp
Y[0] = lanewidth/4
for i in 1..(2*NumberOfLanes - 1):
    Y[i] = Y[i-1] + lanewidth/2
    if Y[i] >= medianYvalue and not median_ya_insertada and NumberOfLanes par:
        Y[i] += MedianLength
        median_ya_insertada = True
# (si NumberOfLanes es impar, la mediana NUNCA se inserta en Y)

# Postes (Single-side)
numberoffixtures = GridLength / polespacing
for k in 0..(int(numberoffixtures)+1):
    fixtureX[k] = k * polespacing
    fixtureY[k] = -polesetback + ArmLength

# Ventana de evaluación en X (con el off-by-one real del VBA)
iStart = index_0based_de( primer X[i] > polespacing )      # ver sección 5.1-5.2
iEnd   = index_0based_de( último X[i] <= 2*polespacing ) - 1   # igual, con el mismo corrimiento

LsumArray[i][j] = 0   para i in iStart..iEnd, j in 0..len(Y)-1

for k in postes:
    for i in iStart..iEnd:
        for j in 0..len(Y)-1:
            dist_horiz = sqrt((fixtureX[k]-X[i])**2 + (fixtureY[k]-Y[j])**2)
            gamma = 0 if dist_horiz==0 else atan(dist_horiz/FixtureHeight) * 180/pi
            dY = fixtureY[k] - Y[j]
            phi = 90 if dY==0 else atan(abs(fixtureX[k]-X[i]) / abs(dY)) * 180/pi   # en [0,90]

            I = interp_bilineal(tablaFotometrica, phi, gamma)   # sección 6, vecino inferior + extrapola en bordes

            E = I * cos(radians(gamma))**3 * LLF / FixtureHeight**2
            LsumArray[i][j] += E

promedio    = mean(LsumArray)
minimo      = min(LsumArray)
maximo      = max(LsumArray)
uniformidad = "" if minimo==0 else minimo/promedio     # inferido por analogía, sección 7
```

---

## 9. Ambigüedades y rarezas (no resolver "limpio" sin decidirlo conscientemente)

- **A1 — Tilt inconsistente en `Illum`.** `larray` se calcula con los ángulos "ForITable" (con tilt), pero `Illum` recibe el `gammaArray` SIN tilt para el `cos(gamma)^3` (`IlluminanceAndLuminance.bas:148,149,153,168`). Con tilt=0 (como está hardcodeado, `:99-101`) esto no importa numéricamente, pero es una inconsistencia estructural: si algún día se activa tilt, el resultado sería incorrecto (mezclaría I(phi',gamma') con cos(gamma) sin tilt).
- **A2 — Sin filtro de distancia en IES.** La rama IES suma la contribución de TODOS los postes de `fixtureX/fixtureY` sin excluir los que están lejos de la ventana de cálculo (a diferencia de la rama CIE que sí anula `LLF=0` para postes a más de 5H, `:377-381`). Para Single-side con 6 postes espaciados `polespacing` cubriendo `4*polespacing`, todos los postes previos/posteriores contribuyen matemáticamente (probablemente de forma insignificante por el término `1/FixtureHeight^2 * cos^3`, pero sin corte explícito).
- **A3 — Rama `If calculationmethod="CIE"` en `Xvalues(0)` es código muerto/no-op** (`MakeMeasurementGrid.bas:34-38`): ambas ramas asignan lo mismo. No afecta el resultado, pero sugiere que el autor original modificó esta lógica y quedó vestigial — posible señal de que hubo una versión con comportamiento diferente por método.
- **A4 — Mediana nunca insertada en Y si `NumberOfLanes` es impar** (`MakeMeasurementGrid.bas:44-50`). Si el caso de uso incluye número impar de carriles con mediana, la coordenada Y de la rejilla del lado lejano quedaría desplazada (le faltaría sumar `MedianLength`), afectando todos los cálculos de ese lado. Confirmar con el usuario si este caso aplica al alcance del proyecto.
- **A5 — Off-by-one de `Match` reutilizado como índice 0-based** (secciones 5.1-5.2). Esto es sistemático en `angleGamma`, `anglePhi`, `angleBeta`, `LintensityMatrix` (todas repiten el mismo patrón). El efecto neto es que la "ventana de un tramo de poste a poste" en realidad:
  - excluye el punto de rejilla más cercano (por debajo) a `polespacing`,
  - y no llega exactamente a `2*polespacing` (se detiene un `GridSpacing` antes de lo que uno esperaría, ya vimos que termina en 68.25 en vez de próximo a 70).
  **A5b** — cuando `polespacing <= 50`, `GridSpacing = polespacing/10` exactamente, por lo que `polespacing` SÍ es múltiplo exacto de `GridSpacing` (`polespacing = 10*GridSpacing`). En ese caso, `X[9] = GridSpacing/2 + 9*GridSpacing = 9.5*GridSpacing = 0.95*polespacing`, y `X[10] = 10.5*GridSpacing = 1.05*polespacing` — **nunca cae exactamente en `polespacing`** porque la rejilla arranca desplazada `GridSpacing/2`. Por lo tanto el comportamiento de "vecino superior" descrito arriba se mantiene siempre (no hay caso de coincidencia exacta), y el ejemplo numérico de la sección 5.2 es representativo del caso general.
  **Decisión pendiente para el reimplemento**: ¿replicar literalmente este corrimiento (fidelidad exacta al Excel) o "corregir" la ventana a `[polespacing, 2*polespacing)` sin el off-by-one? Esto cambia los resultados numéricos y debe decidirse explícitames, ya que no es un detalle cosmético.
- **A6 — Nombres de tabla `tablePhi`/`tableGamma` posiblemente invertidos respecto a la convención de ejes del resto del programa.** El comentario en `LuminanceIntensity.bas:52` llama a la fila de encabezado "Row with angles (vert angles)" pero esa fila se usa como `tableGamma1` (eje columnas) mientras que la columna A (etiquetada como si fuera "Phi") se llama `tablePhi1` (eje filas). Dado que en el resto del programa "gamma" = ángulo vertical y "phi" = ángulo horizontal, los nombres SÍ parecen consistentes con esa convención (fila=horizontal=Phi, columna=vertical=Gamma) — pero el comentario de línea 52 sugiere lo contrario y podría ser simplemente un comentario obsoleto/erróneo del autor. No afecta el cálculo (los nombres son solo etiquetas), pero genera confusión al leer el código; se aconseja verificar con un archivo `.ies` real de ejemplo qué eje trae qué ángulos antes de fijar la convención en Python.
- **A7 — Sin manejo de rango horizontal 0-180/0-360.** Sección 6.4: si la tabla fotométrica importada no es de tipo 0-90° pura, la búsqueda de `phi` (que siempre está en `[0,90]`) dentro de la tabla completa produce resultados silenciosamente incorrectos (no hay espejo/normalización). Se debe decidir si el alcance del proyecto solo admite archivos IES 0-90° (común en luminarias viales simétricas) o si hay que añadir esa normalización en Python (el VBA original NO la tiene).
- **A8 — Fórmulas de promedio/mínimo/máximo/uniformidad para IES no existen en el VBA disponible** (sección 7): se infieren por analogía con el bloque CIE de luminancia. Se necesitaría inspeccionar las fórmulas de hoja de cálculo del archivo `.xlsm` original (`Illuminance Calcs`, `MResults`, `Illuminance`) para confirmarlas, ya que no están en los `.bas` extraídos.
- **A9 — Comentario de "interpolación cuadrática" es engañoso.** Tanto `LintensityMatrix` como su docstring (`LuminanceIntensity.bas:5`, `IlluminanceAndLuminance.bas:152`) llaman "quadratic interpolation" a lo que en realidad es una interpolación **bilineal** (2 puntos por eje, no 3). No hay ningún término cuadrático real en la fórmula (sección 6.2-6.3).
- **A10 — LLF real no verificable desde VBA** (sección 4): se asume `LLF = LLD*LDD*BF` por convención fotométrica estándar, pero la celda `FixtureData!H6` es una fórmula de Excel no incluida en los archivos `.bas`.

---

## 10. Referencias cruzadas rápidas (archivo:línea)

| Elemento | Archivo:línea |
|---|---|
| `GridSpace(IES)` | `MakeMeasurementGrid.bas:90-96` |
| `TotalGridLength(IES)` | `MakeMeasurementGrid.bas:109-111` |
| `makeGrid` (X,Y) | `MakeMeasurementGrid.bas:5-86` |
| `FixturePosition` Single-side | `FixturePositions.bas:30,41-45` |
| Ventana `iStart/iEnd` (IES) | `AngleCalculations.bas:16-17` (repetida 4 veces + `LuminanceIntensity.bas:25-26`) |
| `angleGamma` | `AngleCalculations.bas:5-48` |
| `anglePhi` (no median) | `AngleCalculations.bas:181-191` |
| `Distance` | `AngleCalculations.bas:112-114` |
| `LintensityMatrix` | `LuminanceIntensity.bas:6-87` |
| `LintensityCalc` (bilineal) | `LuminanceIntensity.bas:89-185` |
| `Illum` (fórmula final) | `IlluminanceAndLuminance.bas:442-456` |
| `LLF` lectura | `IlluminanceAndLuminance.bas:67` |
| `LLD/LDD/BF` import | `ReadISO.bas:118-130, 555-557` |
| `CandelaMult` aplicado en import | `ReadISO.bas:578` |
| Suma sobre postes | `IlluminanceAndLuminance.bas:190-203` |
| Volcado a hoja | `IlluminanceAndLuminance.bas:206` |
| Patrón agregados (analogía CIE) | `IlluminanceAndLuminance.bas:319-326` |
