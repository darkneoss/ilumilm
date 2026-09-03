# Procedencia del material de terceros

Este repositorio incluye material que **no es obra propia**. Va aquí acreditado,
separado de lo que sí lo es, y con lo que se sabe de sus términos —incluido lo
que no se sabe—.

El código propio (`engine/`, `tools/`, `tests/`, la documentación y la skill)
está bajo licencia MIT, ver [`LICENSE`](../LICENSE).

---

## 1. SEAD Street Lighting Tool — el VBA de `sead_vba_1.7.6/` y `sead_vba_1.8.1/`

**Qué es.** El código fuente Visual Basic de la *SEAD Street Lighting Tool*, la
herramienta Excel que este proyecto sustituye. Son 126 módulos `.bas` en dos
versiones: la 1.7.6, extraída del `.xls`, y la **1.8.1, que es la que generó
todas las corridas de referencia** contra las que se validó el motor.

**Quién lo hizo.** La herramienta es de la iniciativa **SEAD**
(*Super-efficient Equipment and Appliance Deployment*), del Clean Energy
Ministerial, y se distribuye de forma gratuita. Las macros llevan en su
encabezado la autoría explícita:

> `'This macro was developed by p2w2.  http://p2w2.com/`
> `'Please contact at CS@perceptive-analytics.com in case of any enquiry`

**Por qué está aquí.** El valor entero de este proyecto es reproducir los
números de esa herramienta: los estudios que ya se entregaron con ella tienen
que seguir defendiéndose. Sin el código a la vista, el motor sería una caja
negra que dice coincidir con otra caja negra. Con él, cualquiera puede auditar
el port línea por línea —de hecho `engine/` cita archivo y número de línea en
sus docstrings— y ver también las rarezas que se replicaron a propósito.

**Cómo se obtuvo.** Descomprimiendo los streams VBA del propio libro con
`tools/extract_vba.py`, que está en este repositorio. El proyecto VBA **no
tiene contraseña ni protección de visualización**: se comprobó descifrando los
campos `CMG`, `DPB` y `GC` del stream `PROJECT` según MS-OVBA 2.4.4, y los tres
salen en cero. Cualquiera que abra el archivo en Excel ve lo mismo. Quien
prefiera regenerarlo desde su propia copia:

```bash
python tools/extract_vba.py "SEAD Street Lighting TOOL.xlsm" reference/sead_vba_1.8.1
```

**Términos: no hay.** Se buscó licencia, copyright o condiciones de uso en el
propio libro (las 34 hojas), en el manual de referencia oficial del Clean
Energy Ministerial y en la ficha de OpenEI. **No existe ninguna declaración**:
ni una licencia abierta, ni un aviso de copyright, ni una restricción. La
herramienta se anuncia como gratuita y su código es legible sin obstáculo, pero
la ausencia de licencia significa, en rigor, que tampoco hay una cesión expresa
de derechos de redistribución.

Se incluye de todas formas, con esta nota y el crédito visible, por tres
razones: el material es de una iniciativa pública cuyo propósito declarado es
que estas herramientas se adopten; está distribuido gratuitamente y sin
protección alguna; y su presencia aquí sirve para verificar un trabajo, no para
competir con el original ni sustituirlo.

**Si eres de SEAD, de CLASP o de p2w2** y prefieres que esto no esté aquí, o
que esté de otra forma: abre un issue y sale del repositorio el mismo día. Las
especificaciones de `ESPEC_SEAD_*.md` y `tools/extract_vba.py` bastan para que
el proyecto siga siendo auditable sin el código incluido.

---

## 2. Las fotometrías de `catalogo/`

Seis archivos `.ies` de luminarios viales de **Construlita Lighting
International, S.A. de C.V.** Son los archivos fotométricos que el fabricante
publica para que se usen en software de diseño de iluminación, que es
exactamente el uso que se les da aquí.

Están en el repositorio porque la suite de regresión los necesita: sin ellos
esos casos se saltan y la validación no se puede reproducir. Son de su
fabricante y no de este proyecto.

---

## 3. La NOM-013-ENER-2013

La norma **no** está en el repositorio. Lo que hay en `engine/nom.py` son sus
tablas de valores —clasificaciones de vialidad, iluminancias mínimas,
uniformidades y DPEA máximos—, que son datos normativos de aplicación
obligatoria en México, publicados en el Diario Oficial de la Federación.
