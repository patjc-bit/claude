# Retiro del guardia nocturno: alternativas

Reporte de retiro del guardia nocturno en farmacias, presentado como tres alternativas (conservadora, moderada y moderada con ventanilla a las 22:00) en lugar de fases. Usa los valores finales de `Simulaciones.xlsx` (hojas Sim 1 Conservador y Sim 2 Moderado).

- `Retiro_guardia_nocturno_alternativas.docx`: reporte (resumen, situación actual, justificación, análisis financiero, alternativas, consideraciones y anexos con las farmacias de cada alternativa).
- `analisis/`: scripts para reproducirlo.
  - `extract.py` lee los valores calculados del Excel y escribe `datos.json`.
  - `build_doc.js` arma el Word (paquete `docx` de npm) con el formato del reporte anterior.

El Excel de origen no está en el repo porque son datos internos. Para reproducir:

```bash
python3 extract.py ruta/Simulaciones.xlsx datos.json
node build_doc.js datos.json ../Retiro_guardia_nocturno_alternativas.docx
```

Datos que no vienen del Excel (cobertura de 89%, 10,794 turnos sin guardia, 1,056 robos de enero a julio de 2026, 53% con guardia presente, reparto de robos por turno, 19% y 12% de Iztapalapa y GAM) se tomaron del contexto del proyecto. La cobertura del turno nocturno (80%) y los parámetros de seguimiento (tres meses, dos robos) están resaltados en amarillo porque falta confirmarlos.
