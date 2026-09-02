# ilumilm

Estudios de iluminación vial y evaluación de cumplimiento de la
**NOM-013-ENER-2013** a partir de archivos fotométricos `.ies`.

Sustituye a una herramienta Excel heredada, que solo corre en Office de 32
bits. Este motor corre en cualquier Python 3.10 o superior, **sin
dependencias externas** — solo biblioteca estándar.

## Estado de validación

Contrastado contra 18 corridas reales de la herramienta que reemplaza, diez de
ellas con coincidencia exacta:

| Métrica | Error máximo |
|---|--:|
| Iluminancia promedio | 0.004 % |
| Iluminancia máxima | 0.008 % |
| Iluminancia mínima | 0.006 % |
| Uniformidad | 0.006 % |

El detalle está en [`reference/VALIDACION.md`](reference/VALIDACION.md) y la
suite de regresión en `tests/`.

**Disposiciones validadas: unilateral, central doble y bilateral opuesta.**
Tresbolillo no aparece en ninguna corrida de referencia disponible. Comparte la
ruta de código con las otras tres, que sí están validadas, pero eso es un
argumento de construcción, no una medición: si vas a entregar un estudio en
tresbolillo, conviene contrastar una corrida antes.

**Inclinación del luminario validada a 0°, 5° y 15°.** Es un parámetro por
luminario, con 0° por omisión, y mueve el resultado de verdad: 15° cambian el
Eprom un 6 % y el Emín un 14 %. Si el luminario se monta cabeceado sobre el
brazo, hay que declararlo.

---

## Uso con la skill (recomendado)

En Claude Code, dentro de este repositorio, basta con pedirlo en lenguaje
natural. La skill `estudio-iluminacion` se activa sola con frases como:

> «Hazme un estudio de alumbrado para una calle de dos carriles de 3.5 m,
> postes unilaterales a 8 m de altura cada 35 m, con el luminario de 50 W»

> «¿Cumple la NOM-013 esta vialidad?»

> «Compara estos tres archivos IES en una secundaria residencial tipo A»

La skill se encarga de:

1. **Preguntar lo que falte.** Geometría, clasificación, pavimento y pérdidas.
   Propone los valores por omisión, así que solo confirmas.
2. **Confirmar los watts contigo.** El campo *input watts* del archivo IES no es
   de fiar: muchos lo dejan en 0 y ponen la potencia solo en texto libre. De ese
   número depende directamente el DPEA, que es uno de los tres criterios de
   cumplimiento, así que siempre te lo muestra antes de calcular.
3. **Guardar `entrada.json` antes de calcular**, para poder repetir la corrida
   variando un parámetro sin volver a capturar todo.
4. **Calcular y publicar el reporte** como enlace que puedes compartir.

### Archivos `.ies` nuevos

Suéltalos en `catalogo/` y pídele que reindexe. A partir de ahí los eliges por
nombre («el Construlita de 50 W tipo II») en vez de rastrear rutas, y puedes
correr el catálogo entero contra una vialidad para ver cuál es el de menor
potencia que cumple.

---

## Uso por línea de comandos

### 1. Ver el catálogo

```bash
python -m engine.catalogo
```

Indexa los `.ies` de `catalogo/` leyendo marca, modelo, watts, lúmenes y simetría
de la propia cabecera del archivo. Un `.ies` que no parsee se registra con su
error en el índice, sin tumbar la indexación.

### 2. Describir el estudio

Crea `estudios/<nombre>/entrada.json`:

```json
{
  "vialidad": {
    "num_carriles": 2,
    "ancho_carril": 3.5,
    "camellon": 0.0,
    "disposicion": "unilateral",
    "altura_montaje": 8.0,
    "interpostal": 35.0,
    "retranqueo": 0.2,
    "largo_brazo": 1.8
  },
  "nom": {
    "clasificacion_vialidad": "Vías secundarias residencial Tipo A",
    "pavimento": "R2"
  },
  "perdidas": { "lld": 0.85, "ldd": 0.90, "bf": 1.0 },
  "luminarios": [
    { "archivo": "V1050UN2M50.ies" },
    { "archivo": "V2100UN2M50.ies", "watts": 100.0, "inclinacion": 5.0 }
  ]
}
```

| Campo | Notas |
|---|---|
| `disposicion` | `unilateral`, `tresbolillo`, `central doble` o `bilateral opuesta`. Acepta también los nombres del Excel («De un solo lado», «Montado en el camellón»). |
| `ancho_carril` | Ancho de **un** carril, no el total de la calzada. |
| `retranqueo` | Del poste a la orilla de la calzada. Si el brazo lo supera, el luminario queda montado sobre el arroyo vehicular. |
| `pavimento` | `R1` a `R4`. **R2 por omisión**, salvo indicación contraria. |
| `perdidas` | Opcional. Por omisión `0.85 × 0.90 × 1.0` (LED). El factor de balastro es 1 porque el driver ya viene en la potencia nominal. |
| `luminarios[].archivo` | Ruta a un `.ies` o nombre de uno indexado en `catalogo/`. |
| `luminarios[].watts` | Opcional. **Sobrescribe** los watts del archivo. Úsalo cuando el `.ies` no los declara o evalúas otro driver. |
| `luminarios[].inclinacion` | Opcional, **0° por omisión**. Cabeceo del luminario sobre el brazo, en grados, positivo hacia la calzada. Es propiedad del luminario, no de la vialidad. No es cosmético: entra al cálculo. |
| `modo_azimut` | Opcional. `correcto` (por omisión, IESNA RP-8). `cuadrante` existe solo para diagnóstico. |

Las clasificaciones se resuelven de forma tolerante: `"tipo a"`, `"residencial
Tipo A"` y `"vias_secundarias_residencial_tipo_a"` valen lo mismo.

### 3. Calcular

```bash
python -m engine.cli estudios/<nombre>/entrada.json
```

Escribe, junto a la entrada, `resultados.json` y `malla.csv`.

### 4. Generar el reporte

```bash
python -m engine.report estudios/<nombre>/resultados.json
```

Produce `reporte.html`, autocontenido y listo para publicar o enviar.

### 5. PDF (opcional)

El reporte trae un botón **Imprimir / PDF** que imprime lo que estés viendo en
ese momento, con los valores que hayas dejado en los controles.

Para generarlo sin abrir el navegador:

```bash
python tools/pdf.py estudios/<nombre>/reporte.html     --clasificacion "Vías primarias" --interpostal 40 --altura 10
```

Los tres argumentos son opcionales; sin ellos usa los valores del propio
estudio. Requiere Playwright con Chromium (`python -m playwright install
chromium`), que no hace falta para nada más.

**En papel hay que decidir antes lo que en pantalla se movía.** El PDF no
lleva controles, así que la clasificación, la interpostal y la altura con que
se generó quedan impresas en su propio bloque, arriba de la comparativa. Un
PDF sin ese dato no se puede auditar: los umbrales de la norma dependen de la
clasificación, y esa no se deduce de la geometría.

---

## El reporte

Además de la comparativa y el veredicto por criterio, deja mover tres
parámetros sin volver a correr nada:

- **Clasificación de vialidad.** No cambia la física, solo los umbrales de la
  norma, así que se resuelve en el navegador.
- **Interpostal y altura de montaje.** Estas sí cambian el cálculo completo, así
  que vienen de un barrido precalculado: 13 × 13 combinaciones, cada una pasada
  por el mismo motor. No son valores interpolados.

Los rangos del barrido (20–50 m de interpostal, 6–12 m de altura) son dos
constantes en `engine/cli.py`; los valores del propio estudio siempre están
incluidos aunque no caigan en la retícula.

El mapa isolux dibuja el poste, el brazo y el luminario en su posición real, y
el DPEA se recalcula al mover el interpostal, porque cambia el área del tramo.

---

## Advertencias

**Es un cálculo de diseño, no un dictamen.** La malla fina del método IES —paso
de `interpostal/10` topado a 5 m, dos puntos por carril, un tramo interpostal
completo— no es la fórmula de nueve puntos del Apéndice C con la que la Unidad
de Verificación mide en campo. Sirve para elegir la sustitución; no sustituye la
verificación.

**Confirma siempre los watts** contra la ficha del fabricante. Es el dato que
más se equivoca y del que depende el DPEA por completo.

**El ancho para el DPEA excluye aceras y camellones**, como pide la sección 6.1
de la norma. El motor lo maneja solo, pero conviene saberlo al comparar contra
otras herramientas.

---

## Estructura

```
engine/ies.py        lector LM-63, resuelve simetría horizontal
engine/geometry.py   malla de cálculo y posiciones de postes
engine/calc.py       ángulos, interpolación y suma de iluminancia
engine/nom.py        tablas de la norma, DPEA y veredicto
engine/catalogo.py   índice de archivos .ies
engine/cli.py        corre un estudio desde entrada.json
engine/report.py     genera el HTML
tools/pdf.py         convierte el reporte a PDF con los parámetros fijos
catalogo/            fotometrías .ies + índice
estudios/<nombre>/   entrada.json y salidas (no se versiona: es trabajo)
reference/           VBA extraído del Excel (v1.8.1 es la de referencia,
                     v1.7.6 se conserva por las citas de las especificaciones),
                     especificaciones y validación
tools/               extracción del VBA y de los casos de regresión
tests/               NOM + regresión contra las corridas de referencia
```

`assets/` (la herramienta Excel original, la norma y las corridas de cliente) no
se versiona. Lo que el repo necesita de ese material ya está derivado a texto en
`reference/`.

## Pruebas

```bash
python -m pytest tests/ -q
```

Las pruebas de regresión leen las fotometrías de `catalogo/`. Si falta alguna,
esos casos se saltan con un mensaje que dice cuál copiar, en vez de fallar como
si el motor estuviera roto.

`tools/` sí usa dependencias externas (`olefile`, `xlrd`, `openpyxl`), pero solo
hacen falta para regenerar el material de `reference/` a partir de los archivos
de Excel originales.
