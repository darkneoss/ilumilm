# Validación del motor contra estudios reales

Contraste de `engine/` contra 14 corridas reales de la herramienta Excel de
referencia, extraídas por `tools/extraer_casos_sead.py` y congeladas en
`reference/casos_sead.json`. La suite vive en `tests/test_regresion_sead.py`.

Los estudios van identificados como `estudio-NN`: los archivos de origen llevan
el nombre de la vialidad en el nombre de archivo, y ese dato es de cliente y no
pertenece al repositorio. La anonimización ocurre en la extracción, no después,
para que una regeneración no vuelva a introducirlos.

Se excluyen las filas «Línea base» de cada estudio: usan un luminario genérico
precargado en el Excel cuyo archivo .ies no tenemos.

## Resultados

| Estudio | Luminario | Geometría | Eprom ref | Eprom motor | err | Emin ref | Emin motor | err | Emax ref | Emax motor | err |
|---|---|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|
| estudio-01 | V1070UN2M50 | 2x2.50 cam 0.0 h8 S39.80 unilateral | 15.506 | 15.506 | -0.001% | 6.929 | 6.929 | +0.002% | 39.414 | 39.411 | -0.005% |
| estudio-01 | V2100UN2M50 | 2x2.50 cam 0.0 h8 S39.80 unilateral | 17.669 | 17.669 | +0.002% | 5.293 | 5.294 | +0.006% | 38.339 | 38.340 | +0.002% |
| estudio-02 | V1070UN2M50 | 2x4.09 cam 0.0 h7 S33.15 unilateral | 18.239 | 18.240 | +0.003% | 10.434 | 10.434 | +0.000% | 44.753 | 44.751 | -0.006% |
| estudio-02 | V2100UN2M50 | 2x4.09 cam 0.0 h7 S33.15 unilateral | 22.878 | 22.879 | +0.003% | 8.559 | 8.559 | +0.007% | 39.790 | 39.791 | +0.003% |
| estudio-03 | V2100UN2M50 | 4x2.68 cam 4.0 h9 S27.00 central doble | 31.814 | 31.814 | +0.000% | 20.685 | 20.685 | +0.000% | 39.210 | 39.210 | +0.000% |
| estudio-03 | V2130UN2M50 | 4x2.68 cam 4.0 h9 S27.00 central doble | 35.702 | 35.702 | +0.000% | 22.122 | 22.122 | +0.000% | 66.653 | 66.653 | +0.000% |
| estudio-04 | V1070UN2M50 | 2x4.71 cam 0.0 h8 S35.83 unilateral | 15.307 | 15.308 | +0.005% | 8.937 | 8.938 | +0.005% | 39.916 | 39.917 | +0.004% |
| estudio-04 | V2100UN2M50 | 2x4.71 cam 0.0 h8 S35.83 unilateral | 18.917 | 18.918 | +0.003% | 7.776 | 7.777 | +0.006% | 35.231 | 35.232 | +0.003% |
| estudio-05 | V1070UN2M50 | 4x3.36 cam 0.0 h6 S30.21 unilateral | 15.297 | 15.297 | +0.001% | 1.222 | 1.222 | +0.015% | 54.715 | 54.712 | -0.006% |
| estudio-05 | V2100UN2M50 | 4x3.36 cam 0.0 h6 S30.21 unilateral | 17.514 | 17.515 | +0.005% | 0.893 | 0.894 | +0.097% | 59.297 | 59.298 | +0.002% |
| estudio-06 | V2130UN2M50 | 2x4.35 cam 0.0 h8 S36.46 unilateral | 24.970 | 24.971 | +0.001% | 12.694 | 12.694 | +0.002% | 64.523 | 64.524 | +0.001% |
| estudio-06 | V3160UN2M50 | 2x4.35 cam 0.0 h8 S36.46 unilateral | 35.001 | 35.002 | +0.002% | 17.047 | 17.048 | +0.002% | 90.936 | 90.940 | +0.005% |
| estudio-07 | V2130UN2M50 | 6x3.33 cam 0.0 h8 S30.00 unilateral | 16.450 | 16.451 | +0.005% | 0.578 | 0.579 | +0.145% | 71.527 | 71.528 | +0.002% |
| estudio-07 | V3160UN2M50 | 6x3.33 cam 0.0 h8 S30.00 unilateral | 22.746 | 22.747 | +0.004% | 0.724 | 0.725 | +0.114% | 96.252 | 96.256 | +0.004% |

## Error máximo observado

| Métrica | Error máximo |
|---|--:|
| Promedio | 0.005 % |
| Minimo | 0.145 % |
| Maximo | 0.006 % |
| Uniformidad | 0.139 % |

## Cobertura

Disposiciones presentes en los estudios y por tanto validadas:
**central doble**, **unilateral**.

**Sin validar: tresbolillo y bilateral opuesta.** No aparecen en ninguno de los
estudios disponibles. Comparten la ruta de código con central doble, que sí quedó
validada, pero eso es un argumento de construcción, no una medición.

## Dos correcciones que salieron de esta validación

1. **Orientación del luminario.** Cada luminario necesita saber hacia qué lado
   mira: en las disposiciones de dos luminarios por tramo, la mitad apunta al
   lado contrario. Sin esto, la calzada se ilumina con la distribución del lado
   de la acera. Error antes de corregir: hasta +55 % en Emax y −78 % en Emin.
2. **Retranqueo en central doble.** No aplica a un poste central: los dos brazos
   salen del eje del camellón. Aplicarlo daba 8 % de error en Emin.
