# VBA del SEAD Street Lighting Tool v1.8.1

> Material de terceros. La herramienta es de la iniciativa SEAD (Clean Energy
> Ministerial) y sus macros las desarrolló p2w2. Ver
> [`reference/PROCEDENCIA.md`](../PROCEDENCIA.md) para la atribución completa y
> los términos en que se incluye.

**Esta es la versión de referencia del motor.** Es la que generó todos los
estudios reales de `assets/estudios/`, así que es contra este código contra el
que se contrasta `engine/`.

Extraído de `assets/SEAD Street Lighting TOOL.xlsm` (stream `xl/vbaProject.bin`)
con la descompresión MS-OVBA de `tools/extract_vba.py`. Solo están los 7 módulos
del cálculo que el port cita; el resto --interfaz, gráficas, traducción-- no se
redistribuye, y el árbol completo se regenera con ese mismo script. La hoja `Versions` del
libro llega hasta la 1.8.1 («IES uploaded fixes»).

`reference/sead_vba_1.7.6/` conserva la versión anterior, extraída del `.xls`.
Se queda en el repositorio porque `ESPEC_SEAD_IES.md` y
`ESPEC_SEAD_ILUMINANCIA.md` citan números de línea de **esa** versión, y porque
el contraste entre ambas explica varias decisiones del port.

## Lo que cambia entre 1.7.6 y 1.8.1

### El tilt está vivo

Es la diferencia que importa. En la 1.7.6 la inclinación se captura en la
interfaz (`Fixtures!T36`, `Add IES Files!L19`) y se guarda en `FixtureData`
columna P, pero **ningún módulo la lee**: el cálculo la tiene fija en cero.

```vba
' 1.7.6 -- IlluminanceAndLuminance.bas:99
tiltOnX = 0 / 180 * WorksheetFunction.Pi        'the up down tilt
```

La propia hoja `Versions` lo dice en la nota de la v1.7.4: «*Current version of
the tool uses tilt equation, but it is hard-coded to 0 tilt*». La v1.7.7 lo
cierra («Finalized tilt») y en la 1.8.1 el valor entra de verdad:

```vba
' 1.8.1 -- IlluminanceAndLuminance.bas:104,110
tiltDegreesX = Sheets("FixtureData").Range("selectedTilt")
tiltOnX = tiltDegreesX / 180 * WorksheetFunction.Pi
```

Cómo se aplica, que es lo que replica `engine/calc.py`:

* Solo gira sobre **X** (arriba-abajo, el cabeceo del luminario sobre el brazo).
  `tiltDegreesY` está fijo en 0 («not applicable currently», `:105`) y
  `tiltOnZ` sale del «ángulo de separación» cuando hay más de un luminario por
  poste, no de la inclinación (`FixturePositions.bas:88`).
* Los ángulos se calculan **dos veces** por cada par luminario-punto
  (`IlluminanceAndLuminance.bas:218-222`): una con la inclinación, que es la
  que consulta la tabla de intensidades, y otra sin ella, que es la del
  trayecto real de la luz.
* `Illum` usa el **γ sin inclinación** para el `cos³` (`:583`).

Esa doble cuenta es la que `ESPEC_SEAD_ILUMINANCIA.md` registró como
«Ambigüedad A1», sospechando un bug. No lo es: es correcta. El γ inclinado
responde *cuánta* luz sale del luminario en esa dirección, y el γ geométrico
la ley del cosseno sobre el pavimento. Son dos preguntas distintas y cada una
quiere su ángulo.

La rotación completa está en `AngleCalculations.bas`, funciones
`angleGammaWithTilt` y `anglePhiWithTilt`. Con `tiltOnY = tiltOnZ = 0` se
reduce a lo que implementa el motor.

### Conteo de luminarios

La 1.7.6 generaba un luminario extra por lado (hasta `x = 5S`) por un
`ReDim FPArrayX(n + 1)` con base 0. La 1.8.1 usa
`CInt(gridlength / polespacing) + 1` y `FixturePositions.bas` reescrito con
`polePhase`. Ver la nota de `posiciones_luminarios` en `engine/geometry.py`.

### Dos luminarios por poste: una función a medio terminar

Las otras dos entradas nuevas de la 1.8.1, ambas en `FixtureData`:

* `selectedFixturesPerPole` — multiplica los luminarios por poste
  (`FixturePositions.bas:31`).
* `selectedSeparationAngle` — separa en X los brazos de un mismo poste y les da
  giros opuestos sobre Z (`FixturePositions.bas:45,88-90`).

**No están portadas, a propósito** (ver abajo). Y por si hiciera falta la razón
principal: **la interfaz no las ofrece de forma usable**, y por eso ninguna
corrida de referencia las usa.

* En la biblioteca de luminarios (`Fixtures!U36:V36`) los encabezados existen y
  están traducidos, pero las celdas de captura de los luminarios agregados por
  el usuario quedan vacías.
* En la hoja de alta (`Add IES Files`) sí hay dónde escribirlas —columnas **M**
  y **N**, una por luminario— y `ReadISO.bas:132-133,146-147` las lee de ahí,
  con 1 y 0 por omisión. Pero sus encabezados son **los dos únicos en inglés**
  de esa fila («Fixtures per Pole (1 or 2)», «Separation Angle (if 2 fixtures
  per pole)») mientras las cuatro columnas vecinas están en español, y son las
  únicas sin el marcador de ayuda `?`. La inclinación sí llegó a traducirse;
  esto no.

O sea que se puede capturar, pero no se anuncia. Quien use la herramienta en
español no tiene forma de saber que existe.

### No confundirla con el doble brazo del camellón

Son mecanismos distintos y confundirlos duplica el DPEA.

En **central doble** la herramienta ya pone dos luminarios por poste: lo hace la
disposición misma, con `numPoleSides = 2` y `adjY_median = 1`, que saca los dos
brazos del eje del camellón en direcciones opuestas
(`FixturePositions.bas:60-62,80-83`). Eso es independiente de
`selectedFixturesPerPole`, y es lo que hace que el DPEA cuente dos luminarios
por tramo en esa disposición sin tocar nada más.

`selectedFixturesPerPole` es un multiplicador aparte que aplica a **cualquier**
disposición. Combinado con central doble daría cuatro brazos por poste y cuatro
veces la carga del tramo. El original lo permite; el port **no**, y por eso:

* No hay corrida de referencia contra la que validarlo, ni forma cómoda de
  conseguirla.
* El caso real de dos luminarios por poste —el camellón— ya está cubierto por la
  disposición, medido contra tres corridas.
* Lo único que aporta de nuevo es una vía silenciosa para cuadruplicar el DPEA,
  que es criterio de cumplimiento.

Declararlo en `entrada.json` da error con el mensaje que apunta a `central
doble`. Estuvo implementado en el commit `590c665` y se quitó en el siguiente;
si algún día aparece una corrida que lo use, el código está en el historial.
