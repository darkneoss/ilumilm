# Instrucciones para agentes

Qué es el proyecto está en [`README.md`](README.md). Esto es lo otro: lo que un
agente haría mal si nadie se lo dice.

*(Este archivo lo leen Claude Code, Hermes y otros entornos, que es la razón
del nombre: el motor no está atado a ninguno.)*

> **In English.** These are project rules for coding agents, kept in Spanish
> because they are inseparable from the code they govern: they point at Spanish
> docstrings, at `RAREZA SEAD` markers in the source and at Spanish reference
> documents. Anyone —or any agent— able to act on these rules is reading Spanish
> anyway. The short version: **this engine's job is to reproduce the numbers of
> the Excel tool it replaces, quirks included, not to improve on them.** Do not
> "fix" anything marked `RAREZA SEAD`, and never relax a test tolerance to make
> a case pass.

## La regla que manda sobre todas

**El objetivo es reproducir los números de la herramienta que se sustituye, no
mejorarlos.** Los estudios ya entregados con ese Excel tienen que seguir
defendiéndose, y para eso el motor debe coincidir con él, rarezas incluidas.

En el código hay comportamientos marcados `RAREZA SEAD` que son objetivamente
mejorables: la malla se comprime con carriles impares y camellón, el arreglo de
postes no es simétrico respecto al tramo evaluado, el retranqueo se ignora en un
poste central. **No los arregles.** Están replicados a propósito, medidos y
documentados en `reference/VALIDACION.md`. Cambiar uno rompe la razón de ser del
proyecto, aunque las pruebas siguieran pasando.

Si algo parece un bug del original, primero mira si ya está documentado como
rareza o como ambigüedad resuelta en `reference/ESPEC_SEAD_ILUMINANCIA.md`.

## La suite de regresión es el contrato

`tests/test_regresion_sead.py` contrasta 24 corridas reales con tolerancia de
**0.05 %**. No es holgura de sobra: el error real es de 0.008 % y la tolerancia
estuvo en 0.5 % hasta que se descubrió que ese margen escondía un poste fantasma
durante toda la validación.

- **Nunca relajes una tolerancia para que pase un caso.** Se deja fallando.
- Corre `python -m pytest tests/ -q` después de tocar cualquier cosa de
  `engine/`. Es menos de un segundo.
- Los casos viven en `reference/casos_sead.json`, que se regenera con
  `tools/extraer_casos_sead.py` desde archivos que **no** están en el clon.

## Al tocar el reporte, ábrelo en un navegador

`engine/report.py` genera HTML con JavaScript, y **las pruebas de Python no ven
ese JavaScript**. Ya pasó: agregar una tabla rompió el reporte entero —el script
tomaba las filas de la comparativa con un selector sin acotar— y la página se
veía perfecta sin hacer absolutamente nada al mover los controles.

`tests/test_reporte.py` vigila las anclas de las que depende el script, pero no
sustituye abrirlo.

Dos reglas del reporte que se ganaron a golpes:

- **El bloque de datos de planificación lleva solo lo que los controles no
  mueven.** Un dato que diga un número distinto al del control de arriba es una
  trampa para quien audite el estudio.
- **La inclinación del brazo se imprime siempre, incluso en cero.** La
  herramienta original la usa en el cálculo y no la escribe en su reporte: dos
  estudios con inclinaciones distintas salían indistinguibles, y eso costó media
  jornada de depuración.

## Un campo de entrada que no se usa debe dar error

Nunca ignores en silencio un campo de `entrada.json`. Ese es exactamente el
defecto del Excel que se sustituye. Si el motor no soporta algo, que lo diga con
un mensaje que explique la alternativa —`engine/cli.py` ya lo hace con
`luminarios_por_poste`—.

## Dependencias

`engine/` es **biblioteca estándar y nada más**, y así se queda: corre en
cualquier Python 3.10+ sin instalar nada. `tools/` sí puede usar externas
(`olefile`, `openpyxl`, `xlrd`, `playwright`), porque solo hacen falta para
regenerar material de `reference/` o producir el PDF.

## El VBA de referencia

`reference/sead_vba_1.8.1/` es la autoridad: es la versión que generó todas las
corridas de contraste. `reference/sead_vba_1.7.6/` está solo porque las
especificaciones citan sus números de línea. De cada una se conservan los 7
módulos del cálculo; ver [`reference/PROCEDENCIA.md`](reference/PROCEDENCIA.md),
que también explica en qué términos se redistribuye ese código ajeno.

## La versión se estampa, y qué la mueve

`engine/__version__` viaja a `resultados.json` y al pie del reporte. Es lo único
que enlaza una memoria de cálculo ya entregada con el código que la produjo, así
que **no la toques a la ligera** y respeta lo que significa cada número:

- **El primero** cambia si cambian los números que el motor produce para una
  misma entrada. Dos estudios calculados con mayores distintos no son
  comparables, y eso debe verse sin leer el diff.
- **El segundo**, si se puede calcular o reportar algo nuevo sin mover lo
  anterior.
- **El tercero**, para arreglos que no mueven ningún número.

Se etiqueta a mano cuando hay algo que congelar. No montes CI ni changelog
automático.

## Datos de cliente

`assets/` y `estudios/` no se versionan y no deben versionarse. Los nombres de
vialidad son datos de cliente. La anonimización ocurre **en la extracción**
(`tools/extraer_casos_sead.py`), no después, para que una regeneración no vuelva
a introducirlos.

## Idioma

Todo va en español —código, comentarios, reporte, documentación— y es
deliberado: la norma es mexicana, el reporte lo lee un cliente mexicano, y
palabras como *interpostal*, *retranqueo*, *camellón* o *DPEA* son el vocabulario
de la propia norma. La única excepción es el resumen en inglés del README.

En el código, los comentarios y docstrings van **sin acentos** (ASCII); el texto
que ve el usuario, con acentos y bien escrito.
