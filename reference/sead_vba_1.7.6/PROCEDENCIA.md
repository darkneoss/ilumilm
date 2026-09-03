# VBA del SEAD Street Lighting Tool v1.7.6

Material de terceros. Ver [`reference/PROCEDENCIA.md`](../PROCEDENCIA.md) para
la atribución completa y los términos.

Extraído de `assets/SEAD street lighting tool.xls` con `tools/extract_vba.py`.
Solo están los 7 módulos del cálculo que citan las especificaciones; el resto
--interfaz, gráficas, traducción-- no se redistribuye. El árbol completo se
regenera con ese mismo script.
La herramienta es de la iniciativa SEAD (Clean Energy Ministerial) y sus macros
las desarrolló p2w2.

**Esta NO es la versión de referencia.** Los estudios de contraste los generó la
1.8.1, en `reference/sead_vba_1.8.1/`. Esta se conserva porque
`ESPEC_SEAD_IES.md` y `ESPEC_SEAD_ILUMINANCIA.md` citan sus números de línea, y
porque el contraste entre ambas explica varias decisiones del port —sobre todo
que aquí la inclinación del luminario está capturada en la interfaz pero muerta
en el cálculo—.
