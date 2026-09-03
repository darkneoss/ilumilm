"""Motor de estudios de iluminacion vial contra la NOM-013-ENER-2013.

La version se estampa en `resultados.json` y en el pie del reporte, y es la
unica forma de saber que codigo calculo un estudio ya entregado. Si dentro de
un ano alguien cuestiona una memoria de calculo, la respuesta util es "la
calculo la version X", con esa version etiquetada en el repositorio y su
registro de validacion al lado.

Que se numere:

* El PRIMER numero cambia si cambian los numeros que produce el motor para una
  misma entrada. Un estudio calculado con 1.x y otro con 2.x no son
  comparables, y eso hay que poder verlo sin leer el diff.
* El SEGUNDO, si se agrega algo que antes no se podia calcular o reportar,
  dejando intactos los resultados anteriores.
* El TERCERO, para arreglos que no mueven ningun numero.

No hay proceso de release ni changelog automatico: se etiqueta a mano cuando
hay algo que congelar. Ver reference/VALIDACION.md para el estado de la
validacion de cada version.
"""

__version__ = "1.0.0"
