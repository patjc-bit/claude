# Picos por hora y septiembre 2026

Actualización del análisis de afluencia y alertas de servicio (6 sucursales) con datos de entradas al 15-sep-2026 y de alertas al 21-sep-2026.

- `Picos_por_hora_y_septiembre_2026.docx`: documento principal (resumen, picos por hora, septiembre, calidad de datos, próximos pasos y fuentes).
- `figuras/`: gráficas incluidas en el documento.
- `analisis/`: scripts para reproducirlo.
  - `load.py` lee los bloques del Excel.
  - `charts.py` genera las figuras (matplotlib).
  - `build_doc.js` arma el Word (paquete `docx` de npm).

El Excel de origen (`Resumen_metricas.xlsx`) no está en el repo porque son datos internos. Para reproducir el análisis, cópialo como `analisis/metricas.xlsx` y corre los scripts desde esa carpeta. `build_doc.js` espera las figuras en el mismo directorio.
