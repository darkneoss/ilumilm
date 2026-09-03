---
name: estudio-iluminacion
description: Use when the user wants a street/roadway lighting study, photometric calculation, or NOM-013-ENER-2013 compliance check for public lighting (alumbrado público). Computes point-by-point illuminance from .ies photometric files, evaluates Eprom, uniformity and DPEA against the norm, and publishes an HTML report. Triggers on "estudio de iluminación", "alumbrado público", "NOM-013", "cumple la norma", "archivo IES", "luminario vial", "DPEA", "sustitución de luminarios", "memoria de cálculo".
---

# Estudio de iluminación vial (NOM-013-ENER-2013)

Motor propio de cálculo fotométrico, validado contra 24 corridas reales de la
herramienta Excel que sustituye: error máximo 0.004 % en Eprom y 0.006 % en Emin,
con las cuatro disposiciones de la NOM medidas (ver `reference/VALIDACION.md`).

## Qué necesitas del usuario

Pregunta solo lo que falte; si ya lo dijo, no lo vuelvas a preguntar. Propón
siempre los valores por omisión, que cubren el caso típico.

**Geometría de la vialidad**
- número de carriles y ancho de carril [m]
- ancho de camellón [m] (0 si no hay)
- disposición de postes: unilateral / tresbolillo / central doble / bilateral opuesta
- altura de montaje [m], distancia interpostal [m]
- retranqueo del poste [m] (del poste a la orilla de la calzada)
- largo del brazo [m]
- ancho de banqueta [m], **opcional y solo cosmético**: completa el perfil de la
  vía en el reporte y no entra en ningún cálculo. No lo preguntes; úsalo si el
  usuario lo menciona al describir la sección.

**Clasificación y pavimento**
- una de las 7 clasificaciones de la NOM (ver `engine/nom.py`)
- pavimento: **R2 por omisión**, salvo que el usuario indique otra cosa. Si no
  lo tiene confirmado, dale el estudio con R2 y dile que en el reporte puede
  probar los otros tres sin volver a calcular: el pavimento solo cambia la
  tabla de umbrales, no la física.

**Pérdidas y luminarios**
- LLD (depreciación de lúmenes) **0.85**, LDD (suciedad) **0.90**, BF **1.0** para LED
- los archivos `.ies` a evaluar
- inclinación del luminario sobre el brazo [grados], **0 por omisión**, positiva
  hacia la calzada. Es por luminario, no de la vialidad. **No la preguntes si el
  usuario no la menciona** — 0 es lo normal y lo que usa la herramienta de
  referencia. Pero si la menciona, decláralas: no es un dato de dibujo, entra al
  cálculo, y 15° mueven el Eprom un 6 % y el Emín un 14 %.
- Si el usuario pide **dos luminarios por poste**, lo que quiere es
  `disposicion: central doble`: esa disposición ya cuelga dos brazos de cada
  poste, uno por sentido, y el DPEA cuenta dos luminarios por tramo sin declarar
  nada. El motor no tiene un multiplicador aparte para eso y declararlo da
  error, a propósito (ver `README.md`).

## Cómo correrlo

1. **Catálogo.** `python -m engine.catalogo` lista los `.ies` ya indexados en
   `catalogo/`. Si el usuario trae archivos nuevos, cópialos ahí y reindexa: a
   partir de la segunda corrida puede pedirlos por nombre en vez de por ruta.

2. **Confirma los watts.** El motor los toma del campo *input watts* del archivo
   IES, pero muchos archivos lo dejan en 0 o traen la potencia solo en texto
   libre. **Siempre enséñale al usuario el watt que se va a usar y pide que lo
   confirme**: de ese número depende directamente el DPEA, que es uno de los
   tres criterios de cumplimiento.

3. **Entrada.** Escribe `estudios/<nombre>/entrada.json` **antes** de calcular,
   para poder repetir la corrida variando un parámetro sin volver a preguntar
   todo. El formato está documentado en `engine/cli.py`.

4. **Cálculo.** `python -m engine.cli estudios/<nombre>/entrada.json`
   → `malla.csv` y `resultados.json` junto al archivo de entrada.

5. **Reporte.** `python -m engine.report estudios/<nombre>/resultados.json`
   → `reporte.html`. Publícalo con la herramienta Artifact y entrega el enlace.

## Lo que hay que decirle al usuario, siempre

Tres advertencias que el reporte ya incluye, pero que conviene no enterrar:

- **Es un cálculo de diseño, no un dictamen.** La malla fina (método IES: paso
  de `interpostal/10` topado a 5 m, dos puntos por carril, un tramo interpostal
  completo) no es la fórmula de 9 puntos del Apéndice C con la que
  la Unidad de Verificación mide en campo. Sirve para decidir la sustitución,
  no para sustituir la verificación.

- **Tresbolillo y bilateral opuesta no están validadas.** No aparecen en
  ninguno de los estudios de referencia disponibles. Comparten la ruta de código
  con central doble, que sí quedó validada, pero eso es un argumento de
  construcción, no una medición. Avísale al usuario si su estudio usa una de
  esas dos disposiciones.

- **El DPEA depende de la disposición.** Los watts conectados por tramo
  interpostal son 1 luminario en unilateral y 2 en las demás disposiciones.

## Estructura

```
engine/ies.py        lector LM-63, resuelve simetría horizontal
engine/geometry.py   malla y posiciones de postes (puerto del VBA)
engine/calc.py       ángulos, interpolación y suma de iluminancia
engine/nom.py        tablas de la NOM, DPEA y veredicto
engine/catalogo.py   índice de archivos .ies
engine/cli.py        corre un estudio desde entrada.json
engine/report.py     genera el HTML
reference/           VBA extraído del Excel + especificaciones y caso de regresión
```
