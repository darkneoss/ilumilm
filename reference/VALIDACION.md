# Validación del motor contra estudios reales

Contraste de `engine/` contra 24 corridas reales de la herramienta Excel de
referencia, extraídas por `tools/extraer_casos_sead.py` y congeladas en
`reference/casos_sead.json`. La suite vive en `tests/test_regresion_sead.py`.

Los estudios van identificados como `estudio-NN`: los archivos de origen llevan
el nombre de la vialidad en el nombre de archivo, y ese dato es de cliente y no
pertenece al repositorio. La anonimización ocurre en la extracción, no después,
para que una regeneración no vuelva a introducirlos.

Los `.xlsx` de origen viven en `assets/estudios/`, que no se versiona. Por eso
`tools/extraer_casos_sead.py` no produce nada en un clon: lo que se versiona es
su salida ya anonimizada, `casos_sead.json`, que es todo lo que la suite
necesita.

Se excluyen las filas «Línea base» de cada estudio: usan un luminario genérico
precargado en el Excel cuyo archivo .ies no tenemos.

## Resultados

| Estudio | Luminario | Geometría | Tilt | Eprom ref | Eprom motor | err | Emin ref | Emin motor | err | Emax ref | Emax motor | err |
|---|---|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|
| estudio-01 | V1070UN2M50 | 2x2.50 cam 0.0 h7.5 S39.80 unilateral | 0° | 15.506 | 15.506 | -0.002% | 6.929 | 6.929 | -0.000% | 39.414 | 39.411 | -0.006% |
| estudio-01 | V2100UN2M50 | 2x2.50 cam 0.0 h7.5 S39.80 unilateral | 0° | 17.669 | 17.669 | -0.000% | 5.293 | 5.293 | -0.000% | 38.339 | 38.339 | +0.000% |
| estudio-02 | V1070UN2M50 | 2x4.09 cam 0.0 h7 S33.15 unilateral | 0° | 18.239 | 18.240 | +0.002% | 10.434 | 10.434 | -0.002% | 44.753 | 44.750 | -0.007% |
| estudio-02 | V2100UN2M50 | 2x4.09 cam 0.0 h7 S33.15 unilateral | 0° | 22.878 | 22.878 | +0.000% | 8.559 | 8.559 | +0.000% | 39.790 | 39.790 | -0.000% |
| estudio-03 | V2100UN2M50 | 4x2.68 cam 4.0 h9 S27.00 central doble | 0° | 31.814 | 31.814 | +0.000% | 20.685 | 20.685 | +0.000% | 39.210 | 39.210 | +0.000% |
| estudio-03 | V2130UN2M50 | 4x2.68 cam 4.0 h9 S27.00 central doble | 0° | 35.702 | 35.702 | +0.000% | 22.122 | 22.122 | +0.000% | 66.653 | 66.653 | +0.000% |
| estudio-04 | V1070UN2M50 | 2x4.71 cam 0.0 h7.5 S35.83 unilateral | 0° | 15.307 | 15.307 | +0.004% | 8.937 | 8.937 | +0.003% | 39.916 | 39.917 | +0.003% |
| estudio-04 | V2100UN2M50 | 2x4.71 cam 0.0 h7.5 S35.83 unilateral | 0° | 18.917 | 18.917 | +0.000% | 7.776 | 7.776 | +0.000% | 35.231 | 35.231 | +0.000% |
| estudio-05 | V1070UN2M50 | 4x3.36 cam 0.0 h6 S30.21 unilateral | 0° | 15.297 | 15.297 | -0.001% | 1.222 | 1.222 | +0.003% | 54.715 | 54.712 | -0.006% |
| estudio-05 | V2100UN2M50 | 4x3.36 cam 0.0 h6 S30.21 unilateral | 0° | 17.514 | 17.514 | -0.000% | 0.893 | 0.893 | -0.000% | 59.297 | 59.297 | +0.000% |
| estudio-06 | V2130UN2M50 | 2x4.35 cam 0.0 h7.5 S36.46 unilateral | 0° | 24.970 | 24.970 | -0.000% | 12.694 | 12.694 | -0.000% | 64.523 | 64.523 | -0.000% |
| estudio-06 | V3160UN2M50 | 2x4.35 cam 0.0 h7.5 S36.46 unilateral | 0° | 35.001 | 35.001 | +0.001% | 17.047 | 17.047 | -0.000% | 90.936 | 90.940 | +0.004% |
| estudio-07 | V2130UN2M50 | 6x3.33 cam 0.0 h7.5 S30.00 unilateral | 0° | 16.450 | 16.450 | +0.000% | 0.578 | 0.578 | +0.000% | 71.527 | 71.527 | +0.000% |
| estudio-07 | V3160UN2M50 | 6x3.33 cam 0.0 h7.5 S30.00 unilateral | 0° | 22.746 | 22.746 | +0.000% | 0.724 | 0.724 | +0.006% | 96.252 | 96.254 | +0.003% |
| estudio-08 | V1070UN2M50 | 4x3.54 cam 0.0 h10 S45.00 unilateral | 0° | 7.626 | 7.626 | +0.002% | 2.888 | 2.888 | +0.004% | 22.850 | 22.852 | +0.008% |
| estudio-08 | V2100UN2M50 | 4x3.54 cam 0.0 h10 S45.00 unilateral | 0° | 9.820 | 9.820 | +0.000% | 1.506 | 1.506 | +0.000% | 20.761 | 20.761 | +0.000% |
| estudio-09 | V1070UN2M50 | 4x3.54 cam 0.0 h10 S45.00 bilateral opuesta | 5° | 15.591 | 15.592 | +0.001% | 9.194 | 9.193 | -0.001% | 27.812 | 27.813 | +0.002% |
| estudio-09 | V2100UN2M50 | 4x3.54 cam 0.0 h10 S45.00 bilateral opuesta | 15° | 18.462 | 18.462 | +0.000% | 10.020 | 10.020 | +0.000% | 30.193 | 30.193 | +0.000% |
| estudio-10 | V1070UN2M50 | 3x3.50 cam 0.0 h7 S40.00 tresbolillo | 0° | 26.902 | 26.901 | -0.001% | 14.733 | 14.733 | -0.005% | 47.695 | 47.695 | +0.000% |
| estudio-10 | V2100UN2M50 | 3x3.50 cam 0.0 h7 S40.00 tresbolillo | 0° | 32.009 | 32.009 | +0.000% | 9.523 | 9.523 | +0.000% | 61.149 | 61.149 | +0.000% |
| estudio-11 | V1070UN2M50 | 3x3.50 cam 0.0 h7 S40.00 central doble | 0° | 19.944 | 19.944 | -0.002% | 8.198 | 8.197 | -0.002% | 46.983 | 46.987 | +0.008% |
| estudio-11 | V2100UN2M50 | 3x3.50 cam 0.0 h7 S40.00 central doble | 0° | 22.764 | 22.764 | +0.000% | 7.509 | 7.509 | +0.000% | 47.763 | 47.763 | +0.000% |
| estudio-12 | V1070UN2M50 | 3x3.50 cam 1.0 h7 S40.00 central doble | 0° | 19.828 | 19.828 | +0.000% | 8.061 | 8.061 | -0.004% | 50.087 | 50.084 | -0.007% |
| estudio-12 | V2100UN2M50 | 3x3.50 cam 1.0 h7 S40.00 central doble | 0° | 22.524 | 22.524 | +0.000% | 7.214 | 7.214 | +0.000% | 47.565 | 47.565 | +0.000% |

## Error máximo observado

| Métrica | Error máximo |
|---|--:|
| Promedio | 0.004 % |
| Minimo | 0.006 % |
| Maximo | 0.008 % |
| Uniformidad | 0.006 % |

Trece de los veinticuatro casos coinciden con la referencia a cero exacto. El
residuo de los otros once está siempre en los `.ies` grandes (V1070, V3160:
~250 KB, malla angular fina) y es de interpolación, no de geometría.

## Cobertura

**Las cuatro disposiciones de la NOM-013 están validadas contra medición:**
unilateral, tresbolillo, central doble y bilateral opuesta. Ya no queda ninguna
apoyada solo en el argumento de que comparte ruta de código con otra.

**Fuera de alcance: el multiplicador de luminarios por poste**
(`selectedFixturesPerPole` y su ángulo de separación). No se implementa, y la
decisión es deliberada: la interfaz del Excel no lo ofrece de forma usable
—columnas sin traducir en la hoja de alta—, ninguna de las 24 corridas lo usa, y
en central doble se confundiría con el doble brazo del camellón, que sí existe y
lo pone la disposición. Declararlo en `entrada.json` da error en vez de
ignorarse. Estuvo implementado brevemente (commit `590c665`) y se quitó por eso;
si algún día hace falta, ahí está el código y lo que falta es una corrida.

Inclinaciones validadas: **0°, 5° y 15°**. Las dos distintas de cero, en
bilateral opuesta y sobre la misma geometría que su control a 0° (estudios 08 y
09), que es lo que permite atribuirle la diferencia a la inclinación y no a la
geometría.

El estudio-11 añade dos condiciones que ningún otro cubría, ambas del caso
central doble:

* **camellón de 0.0 m con montaje central**, o sea los dos brazos saliendo del
  eje de la calzada sin camellón físico. El único caso central doble anterior
  (estudio-03) tenía 4 m.
* **retranqueo de 0.2 m que el motor ignora a propósito.** En un poste central
  no tiene sentido físico —los dos brazos salen del mismo eje— y el motor lo
  descarta; el VBA lo suma y lo resta igual, con un comentario propio admitiendo
  que sobra. Que la corrida cuadre al 0.008 % con retranqueo distinto de cero
  confirma la decisión por segunda vez y con otra geometría.

## La malla comprimida con carriles impares: «Ambigüedad A4» medida

El estudio-12 es el mismo caso central doble del 11 pero con **camellón de 1 m
y tres carriles**, es decir un número impar. La especificación tenía esto
marcado como pendiente de confirmar (`ESPEC_SEAD_ILUMINANCIA.md`, A4): el VBA
solo inserta el camellón en las coordenadas Y de la rejilla cuando el número de
carriles es **par**, porque con impares pone `medianYvalue = 0` y baja la
bandera.

La corrida lo confirma: cuadra al 0.007 % **replicando la rareza**. Con 3
carriles de 3.5 m y 1 m de camellón, la rejilla cubre 10.5 m cuando la sección
física mide 11.5, y a la vez los postes sí se colocan en el eje contando el
camellón. La malla y los postes están medidos con reglas distintas.

Esto no es un defecto del port: es fidelidad, que es el objetivo declarado. Pero
sí es un límite que conviene tener presente al entregar un estudio con carriles
impares y camellón, porque el tramo evaluado no es exactamente la sección
transversal real. El motor reproduce lo que la herramienta que se sustituye
entregó, y eso es lo que permite defender los estudios ya emitidos; si algún día
se quiere corregir, tiene que ser una decisión consciente y con su propia
validación, no un arreglo silencioso.

## Cuatro correcciones que salieron de esta validación

1. **Orientación del luminario.** Cada luminario necesita saber hacia qué lado
   mira: en las disposiciones de dos luminarios por tramo, la mitad apunta al
   lado contrario. Sin esto, la calzada se ilumina con la distribución del lado
   de la acera. Error antes de corregir: hasta +55 % en Emax y −78 % en Emin.
2. **Retranqueo en central doble.** No aplica a un poste central: los dos brazos
   salen del eje del camellón. Aplicarlo daba 8 % de error en Emin.
3. **Número de luminarios por lado.** Eran uno de más (hasta x = 5·S en vez de
   4·S), heredado de un `ReDim` base 0 de la v1.7.6 que la 1.8.1 corrigió. El
   sobrante aportaba poco, pero quitarlo bajó el error máximo de la suite de
   0.145 % a 0.008 % y volvió exactos ocho de los dieciséis casos.
4. **La inclinación del luminario faltaba por completo** (ver abajo).

## La inclinación: cómo se encontró

El estudio-09 fue el primero con **bilateral opuesta**, justo la disposición que
estaba sin validar, y no cuadraba: hasta 6 % en Eprom, 14 % en Emin y **24 % en
uniformidad**. Con esa disposición recién estrenada, la lectura obvia era un bug
en su ruta de código.

No lo era. Faltaba un parámetro de entrada.

La herramienta captura la inclinación **por luminario** (`Fixtures!T36`,
guardada en `FixtureData!selectedTilt`) y **no la escribe en el reporte de
salida**: en el `.xlsx` la cadena «Inclinación ( grados)» solo aparece en la hoja
de traducciones. Dos corridas con inclinaciones distintas producen archivos
indistinguibles salvo por los números. Por eso pasó inadvertida durante toda la
validación anterior: los siete primeros estudios se corrieron con 0°, que es el
valor por omisión.

El port tampoco la tenía, porque se hizo contra el VBA de la **v1.7.6**, donde
la inclinación se captura en la interfaz pero está *muerta* en el cálculo
(`tiltOnX = 0 / 180 * Pi`, a mano). La v1.7.7 la conectó («Finalized tilt» en la
hoja `Versions`) y la 1.8.1 —la que generó todos estos estudios— la usa de
verdad. Ver `reference/sead_vba_1.8.1/README.md`.

La confirmación fue un barrido: implementada la ecuación de la 1.8.1, se probaron
tres inclinaciones por caso.

| Corrida | Luminario | Tilt | Eprom | Emin | Emax | Unif. |
|---|---|--:|--:|--:|--:|--:|
| estudio-09 | V1070 | 0° | 2.180 % | 4.087 % | 5.574 % | 1.988 % |
| | | **+5°** | **0.001 %** | **0.001 %** | **0.002 %** | **0.002 %** |
| | | −5° | 6.760 % | 13.138 % | 11.073 % | 7.342 % |
| estudio-09 | V2100 | 0° | 6.382 % | 14.501 % | 5.041 % | 24.425 % |
| | | **+15°** | **0.000 %** | **0.000 %** | **0.000 %** | **0.000 %** |
| | | −15° | 10.101 % | 19.726 % | 11.780 % | 11.991 % |
| estudio-08 | V1070 | **0°** | **0.002 %** | 0.004 % | 0.008 % | 0.002 % |
| | | +5° | 2.232 % | 19.118 % | 1.208 % | 14.176 % |
| estudio-08 | V2100 | **0°** | **0.000 %** | **0.000 %** | **0.000 %** | **0.000 %** |
| | | +15° | 5.999 % | 122.276 % | 26.996 % | 57.710 % |

Cuatro métricas independientes, dos luminarios y dos ángulos, todo al cuarto
decimal. Los 5° y 15° que declaró quien corrió el estudio quedan confirmados por
medición, y el signo positivo —inclinación hacia la calzada— también: el negativo
se desvía entre 7 % y 20 %.

Con la inclinación puesta, **bilateral opuesta cuadra al 0.002 %**. Su ruta de
código siempre estuvo bien.

Lección de método: cuando una disposición recién validada no cuadra, el sospechoso
natural es su código, pero conviene descartar primero que no falte una entrada.
Aquí el 24 % de error no estaba en el motor: estaba en un campo del Excel que el
propio Excel no reporta.
