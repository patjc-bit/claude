const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, ImageRun, LevelFormat, ExternalHyperlink,
  BorderStyle, Footer, PageNumber,
} = require('docx');

const W = 9360; // Letter, 1" margins
const INK = '0B0B0B', INK2 = '52514E', MUTED = '898781', HEAD_FILL = 'E8EEF7', ALT_FILL = 'F6F6F4';

// ---------- helpers
const r = (text, o = {}) => new TextRun({ text, ...o });
const b = (text, o = {}) => new TextRun({ text, bold: true, ...o });
const P = (children, o = {}) => new Paragraph({ children: Array.isArray(children) ? children : [r(children)], spacing: { after: 120 }, ...o });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [r(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [r(t)] });
const bullet = (children, level = 0) => new Paragraph({ numbering: { reference: 'bul', level }, children: Array.isArray(children) ? children : [r(children)], spacing: { after: 80 } });
const num = (children) => new Paragraph({ numbering: { reference: 'num', level: 0 }, children: Array.isArray(children) ? children : [r(children)], spacing: { after: 80 } });
const note = (t) => new Paragraph({ children: [r(t, { italics: true, color: MUTED, size: 17 })], spacing: { after: 200 } });
const link = (text, url) => new ExternalHyperlink({ link: url, children: [r(text, { style: 'Hyperlink' })] });

function table(widths, rows, { header = true, fontSize = 17 } = {}) {
  const total = widths.reduce((a, c) => a + c, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'D9D8D2' };
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((row, i) => new TableRow({
      tableHeader: header && i === 0,
      children: row.map((cell, j) => new TableCell({
        width: { size: widths[j], type: WidthType.DXA },
        shading: header && i === 0 ? { type: ShadingType.CLEAR, fill: HEAD_FILL, color: 'auto' }
          : (i % 2 === 0 ? { type: ShadingType.CLEAR, fill: ALT_FILL, color: 'auto' } : undefined),
        borders: { top: border, bottom: border, left: border, right: border },
        margins: { top: 60, bottom: 60, left: 90, right: 90 },
        children: [new Paragraph({
          children: (Array.isArray(cell) ? cell : [cell]).map(c => typeof c === 'string'
            ? new TextRun({ text: c, bold: header && i === 0, size: fontSize, color: INK })
            : c),
        })],
      })),
    })),
  });
}

function figure(path, wPx, hPx, caption) {
  const widthPt = 468; // 6.5in
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
      children: [new ImageRun({ type: 'png', data: fs.readFileSync(path), transformation: { width: widthPt, height: Math.round(widthPt * hPx / wPx) } })] }),
    new Paragraph({ children: [r(caption, { italics: true, color: INK2, size: 17 })], spacing: { after: 200 } }),
  ];
}

// ---------- content
const children = [];

children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [r('Picos por hora y septiembre 2026')] }));
children.push(P([r('Actualización del análisis de afluencia y servicio · 6 sucursales · datos de entradas al 15-sep y de alertas al 21-sep-2026', { color: INK2 })]));
children.push(P([r('Continúa el documento “El 17 de julio de 2026, el primer viernes tras el inicio de vacaciones de verano SEP”.', { color: MUTED, italics: true, size: 18 })], { spacing: { after: 240 } }));

// ===== 1. Resumen
children.push(H1('1. Resumen'));
children.push(bullet([b('El “pico de las 15:00” es en realidad una meseta. '), r('En TAPO, 15:00 (51.4 entradas/h) y 18:00 (50.7) difieren 1.4 %; cinco horas seguidas (15–19 h) están a menos de 10 % del máximo. En ZIBATA el máximo ahora es 14:00 (65.7/h). Con los datos nuevos ya no se cumple lo que decía el documento anterior (“pico 15–16 h en AMBAS sucursales”).')]));
children.push(bullet([b('Las 3 sucursales nuevas tienen el pico al mediodía (11–13 h), '), r('2–3 horas antes que las anteriores. Es preliminar: solo hay 12 días completos por sucursal.')]));
children.push(bullet([b('El mostrador desatendido parece un problema de personal, no de demanda. '), r('El pico es a las 14–15 h en las 6 sucursales, aun donde el tráfico ya bajó. En XOCHIMILCO hay 27 % más alertas a las 15 h que en su hora de más tráfico, con 24 % menos entradas. También hay un repunte a las 07:00 en las 6: más alertas que a las 08:00 aunque entra menos gente. Es la huella típica de las comidas del personal, la apertura y los cambios de turno.')]));
children.push(bullet([b('La hipótesis del desfase tráfico→fila se cumple en ZIBATA, no en TAPO. '), r('En TAPO la fila (unifila) a las 15 h (1.78 alertas) es igual que a las 18 h (1.79): no hay desfase real. En ZIBATA sí lo hay: 29 % más alertas a las 19 h que a las 14 h. La teoría de colas dice que el pico de la fila llega unos minutos después del pico de llegadas (más o menos un tiempo de servicio), no horas después. Un desfase de 5 horas solo se explica por menos capacidad o por atenciones más largas en la tarde-noche.')]));
children.push(bullet([b('En septiembre, el regreso a clases (31-ago) casi no movió el tráfico de las anteriores: '), r('TAPO −0.3 % y ZIBATA +2.4 %, comparando los mismos días de la semana. El salto grande de ZIBATA fue en la 1.ª semana de agosto (+65 %, de ~540 a ~890 entradas/día). Coincide con el inicio del semestre universitario (la Anáhuac Querétaro está en Zibatá), pero falta descartar un cambio de sensor.')]));
children.push(bullet([b('La quincena no aparece en las entradas '), r('(índice ajustado 1.00 contra 1.00 en TAPO). Si tiene efecto, sería en el monto del ticket y no en cuántas personas entran. Eso solo se ve con datos del punto de venta.')]));
children.push(bullet([b('Hay tres alertas de calidad de datos: '), r('CHAPALITA probablemente cuenta de menos las entradas; las salidas no cuadran con las entradas (de −13 % a +16 %); y ZIBATA marca mostrador desatendido de madrugada sin ninguna entrada.')]));

// ===== 2. Qué cambió
children.push(H1('2. Qué cambió en los datos'));
children.push(P('El Excel pasó de 2 a 6 sucursales. Las nuevas (AVCONS, SANJACINTO, XOCHIMILCO) empiezan el 2-sep. La Hoja 2 agrega alertas de servicio: “mostrador desatendido” y “unifila”.'));
children.push(table([1350, 1150, 1150, 1250, 1400, 1020, 1020, 1020], [
  ['Sucursal', 'Grupo', 'Días con datos', 'Prom. entradas/día', 'Día más activo', 'Pico entradas', 'Pico mostrador', 'Pico unifila'],
  ['TAPO', 'Anterior', '69', '644', 'Viernes', '15 h*', '15 h', '18 h*'],
  ['ZIBATA', 'Anterior', '76', '713 (898 en sep)', 'Jueves', '14 h', '14 h', '19 h'],
  ['CHAPALITA', 'Anterior', '71', '121', 'Viernes', '22 h**', '14 h', '19 h'],
  ['AVCONS', 'Nueva', '14', '314', 'Lunes (n=2)', '11 h', '15 h', '12 h'],
  ['SANJACINTO', 'Nueva', '14', '193', 'Lunes (n=2)', '12 h', '14 h', '20 h'],
  ['XOCHIMILCO', 'Nueva', '14', '857', 'Viernes (n=2)', '12 h', '15 h', '12 h'],
]));
children.push(note('* En TAPO el pico de entradas (15 h) y el de unifila (18 h) son prácticamente empates con las horas vecinas. ** CHAPALITA: pico de entradas poco confiable (ver sección 5). “n=2” = solo dos observaciones de ese día de la semana.'));
children.push(P([r('Ubicaciones supuestas (confirmar): ', { bold: true }), r('TAPO – terminal de autobuses de oriente, CDMX; ZIBATA – El Marqués, Qro.; CHAPALITA – zona metropolitana de Guadalajara; AVCONS – Av. Constituyentes (probablemente Querétaro); SANJACINTO y XOCHIMILCO – CDMX. Importa porque el calendario escolar es nacional, pero la lluvia, los traslados y los polos de atracción (terminal, universidades, oficinas) son locales.')]));

// ===== 3. Picos por hora
children.push(H1('3. Picos por hora'));
children.push(H2('3.1 Tráfico: meseta en las anteriores, mediodía en las nuevas'));
children.push(...figure('fig1_perfil_horario.png', 1840, 1080, 'Figura 1. Entradas, mostrador desatendido y unifila por hora (7–22 h). Cada línea está en % de su propio máximo para comparar la forma de la curva; el punto marca la hora pico de cada serie.'));
children.push(P('Qué porcentaje del tráfico diario cae en cada franja:'));
children.push(table([1560, 1300, 1300, 1300, 1300, 1300, 1300], [
  ['Franja', 'TAPO', 'ZIBATA', 'CHAPALITA', 'AVCONS', 'SANJACINTO', 'XOCHIMILCO'],
  ['07–11 h', '18.9 %', '14.2 %', '18.6 %', '16.2 %', '22.0 %', '19.6 %'],
  ['11–14 h', '18.3 %', '21.8 %', '15.0 %', b('28.3 %', { size: 17 }), b('26.3 %', { size: 17 }), b('25.0 %', { size: 17 })],
  ['14–17 h', '22.1 %', b('24.9 %', { size: 17 }), '17.0 %', '25.5 %', '22.5 %', '20.2 %'],
  ['17–20 h', b('22.4 %', { size: 17 }), '24.7 %', '18.6 %', '18.6 %', '17.0 %', '17.7 %'],
  ['20–23 h', '14.2 %', '14.3 %', '19.8 %', '8.9 %', '10.6 %', '9.8 %'],
]));
children.push(note('Negritas = franja con mayor participación. Promedios de todo el periodo disponible de cada sucursal.'));
children.push(bullet([b('TAPO y ZIBATA reparten casi la mitad del día entre 14 y 20 h. '), r('La franja 14–17 h pesa lo mismo que 17–20 h. Para planear personal hay que pensar en un bloque de 6 horas, no en una hora pico.')]));
children.push(bullet([b('Las nuevas concentran 25–28 % del día entre 11 y 14 h '), r('y caen fuerte después de las 17 h. Encaja con zonas de oficinas o escuelas y con la salida del turno matutino escolar (≈12:30–13:30). Hay que ver si se mantiene cuando haya más semanas de datos.')]));

children.push(H2('3.2 Mostrador desatendido: la señal apunta al personal'));
children.push(P('Si las alertas de mostrador dependieran solo de cuánta gente entra, su pico coincidiría con el del tráfico. Eso pasa en TAPO y ZIBATA, donde el tráfico también es máximo a las 14–15 h, así que ahí no se pueden separar las dos causas. En las otras cuatro sucursales el pico de alertas cae a las 14–15 h, cuando ya entra menos gente:'));
children.push(table([1760, 1400, 1500, 2350, 2350], [
  ['Sucursal', 'Pico entradas', 'Pico mostrador', 'Alertas en su pico vs. en la hora pico de tráfico', 'Entradas en el pico de alertas vs. hora pico'],
  ['TAPO', '15 h', '15 h', '0 % (coinciden)', '0 %'],
  ['ZIBATA', '14 h', '14 h', '0 % (coinciden)', '0 %'],
  ['CHAPALITA', '22 h**', '14 h', '+23 %', '−22 %'],
  ['AVCONS', '11 h', '15 h', '+25 %', '−10 %'],
  ['SANJACINTO', '12 h', '14 h', '+28 %', '−11 %'],
  ['XOCHIMILCO', '12 h', '15 h', '+27 %', '−24 %'],
]));
children.push(P([b('Repunte de las 07:00. '), r('En las 6 sucursales hay más alertas de mostrador a las 07 h que a las 08 h (+5 % a +19 %). En 5 de las 6, a las 07 h entra menos gente que a las 08 h. Un repunte al abrir, sin más clientes, apunta a apertura o cambio de turno: acomodar mercancía, abrir caja, entregar el turno.')]));
children.push(P([b('Por qué 14–16 h. '), r('En México la comida principal es entre 14 y 16 h, incluso en oficinas (La Verdad Noticias, 2026). La Ley Federal del Trabajo (art. 63) obliga a dar al menos 30 minutos de descanso en jornada continua. Si el equipo de la sucursal también come en esa franja, el mostrador se queda sin gente justo cuando llegan los clientes que salen a comer. Esto se puede comprobar comparando el horario de comidas del personal con las alertas por hora.')]));

children.push(H2('3.3 Unifila: revisión de la hipótesis anterior'));
children.push(P([r('El documento anterior proponía: '), r('“a las 15:00 hay más personas con compras rápidas; a las 19:00, compras planeadas de mayor ticket que generan filas”.', { italics: true }), r(' Con los datos nuevos:')]));
children.push(table([1760, 1400, 1650, 1400, 1650, 1500], [
  ['Sucursal', 'Pico entradas', 'Unifila en esa hora', 'Pico unifila', 'Unifila en su pico', 'Diferencia'],
  ['TAPO', '15 h', '1.78', '18 h', '1.79', '≈0 % (empate)'],
  ['ZIBATA', '14 h', '1.00', '19 h', '1.29', '+29 %'],
  ['CHAPALITA', '22 h**', '0.74', '19 h', '1.05', 'n/c'],
  ['AVCONS', '11 h', '1.16', '12 h', '1.37', '+18 %'],
  ['SANJACINTO', '12 h', '0.99', '20 h', '1.07', '+8 %'],
  ['XOCHIMILCO', '12 h', '2.48', '12 h', '2.48', '0 % (coinciden)'],
]));
children.push(note('Unifila = alertas promedio por hora (Hoja 2).'));
children.push(...figure('fig2_unifila_por_cliente.png', 1840, 660, 'Figura 2. Alertas de unifila por cada cliente que entra, en índice (1.0 = promedio de 9–20 h de la sucursal). Rojo = más fila por cliente de lo normal para esa sucursal; azul = menos.'));
children.push(P([b('Lectura: '), r('en TAPO la fila crece al mismo ritmo que el tráfico todo el día (índice entre 0.90 y 1.17). En ZIBATA, AVCONS y SANJACINTO la fila por cliente sube en los extremos del día: al abrir (8–9 h, hasta 1.75) y en la noche (19–21 h, hasta 1.52). En esas horas hay menos gente atendiendo o cada atención tarda más.')]));
children.push(P([b('Qué dice la teoría de colas. '), r('En sistemas con llegadas que varían en el tiempo, el pico de congestión llega después del pico de llegadas, pero el retraso es de más o menos un tiempo de servicio: ≈1/μ con servicio exponencial y ≈1/(2μ) con servicio fijo (Eick, Massey y Whitt, 1993; Green y Kolesar, 1998). Si atender a un cliente toma unos minutos (supuesto por validar con el punto de venta), el retraso esperado es de minutos, no de 4–5 horas. Por eso el desfase de ZIBATA tiene que venir de la capacidad o del tiempo de servicio en la tarde-noche.')]));
children.push(P([b('Veredicto sobre la hipótesis anterior: '), r('se puede sostener, pero con estos datos no se puede comprobar, y hay explicaciones igual de sencillas. Queda como una de cuatro hipótesis que compiten:')]));
children.push(table([700, 3000, 2960, 2700], [
  ['#', 'Hipótesis', 'Contexto que la respalda', 'Dato que la confirma o descarta'],
  ['H1', 'Mezcla de compra: en la tarde-noche los tickets son más grandes y tardan más (hipótesis original).', 'Ninguno externo; es de comportamiento.', 'Punto de venta por hora: ticket promedio, artículos por ticket, tiempo de cobro.'],
  ['H2', 'Capacidad: menos gente en caja de 18 a 21 h (turno vespertino más corto, comidas tardías, tareas de cierre).', 'Repunte de alertas al abrir y al cerrar en 3 sucursales (Fig. 2).', 'Plantilla por hora: personas en mostrador y caja contra alertas.'],
  ['H3', 'Recetas después de consulta: más clientes con receta en la tarde, que tardan más en atenderse.', 'Las UMF del IMSS dan consulta de 8 a 14 h y de 14 a 20 h. El desabasto hace que muchas recetas se surtan en farmacias privadas. Los antibióticos exigen sellar la receta y guardar copia (COFEPRIS).', '% de tickets con receta o antibiótico por hora.'],
  ['H4', 'Llegadas en ráfaga: la gente llega concentrada en ciertos minutos dentro de la hora.', 'EOD 2017: ~2 millones de viajes empiezan entre 18:00 y 18:14 en la zona metropolitana del Valle de México. Lluvias vespertinas en CDMX.', 'Entradas cada 15 minutos (no por hora).'],
]));

children.push(H2('3.4 Contexto externo por franja horaria'));
children.push(table([1150, 4200, 4010], [
  ['Franja', 'Factor externo (fuente)', 'Qué predice en los datos'],
  ['07–08 h', 'Entrada escolar del turno matutino (≈7:00–8:00, SEP). Más de 4 millones de viajes empiezan entre 7:00 y 7:59 en la zona metropolitana del Valle de México (INEGI, EOD 2017).', 'Tráfico que sube rápido. El repunte de alertas de mostrador a las 07 h no lo explica la demanda: es operación interna (apertura o cambio de turno).'],
  ['11–14 h', 'Salida del turno matutino escolar (≈12:30–13:30). Consulta matutina del IMSS (8–14 h) seguida de surtir la receta.', 'Pico de mediodía en las sucursales nuevas. En septiembre podría haberse reforzado por el regreso a clases.'],
  ['14–16 h', 'Hora de la comida entre 14 y 16 h (La Verdad Noticias, 2026). Descanso mínimo de 30 min en jornada continua (LFT, art. 63).', 'Máximo de tráfico en TAPO y ZIBATA y pico de mostrador desatendido en las 6 sucursales.'],
  ['18–20 h', 'Salida laboral (jornada que suele terminar entre 18 y 19 h). ~2 millones de viajes empiezan entre 18:00 y 18:14 (EOD 2017). Consulta vespertina del IMSS hasta las 20 h. Tormentas de tarde-noche en CDMX (UNAM).', 'Segundo tramo alto en TAPO y ZIBATA y pico de unifila en ZIBATA, CHAPALITA y SANJACINTO.'],
  ['A futuro', 'Reforma de 40 horas publicada en el DOF (2026), con reducción gradual: 48 h en 2026, 46 h en 2027, hasta 40 h en 2030.', 'No afecta septiembre de 2026. Desde 2027 puede adelantar la salida laboral y mover el pico vespertino.'],
]));

// ===== 4. Septiembre
children.push(H1('4. Septiembre'));
children.push(H2('4.1 Lo que dicen los datos'));
children.push(...figure('fig3_diario.png', 1840, 780, 'Figura 3. Entradas diarias de las sucursales anteriores con media móvil de 7 días y fechas clave. Se excluyen los días parciales: primer día de cada sucursal y 15-sep.'));
children.push(table([2400, 2320, 2320, 2320], [
  ['Periodo (promedio entradas/día)', 'TAPO', 'ZIBATA', 'CHAPALITA'],
  ['1–15 jul (todavía clases SEP)', '628 (desde 9-jul)', '468', '171 (desde 7-jul)'],
  ['16 jul – 2 ago (vacaciones SEP)', '667', '536', '136'],
  ['3–30 ago (vacaciones SEP)', '648', '859', '106'],
  ['31 ago – 14 sep (clases SEP)', '645', '895', '110'],
  ['Cambio mismos días de la semana: 17–30 ago vs. 31 ago–13 sep', '−0.3 %', '+2.4 %', '−17 % (ruido alto)'],
]));
children.push(note('Solo días completos. TAPO y CHAPALITA empiezan a registrar el 8 y el 6 de julio respectivamente; ese primer día (parcial) se excluye.'));
children.push(bullet([b('TAPO (terminal de autobuses): '), r('lo que más la movió fueron las vacaciones, no septiembre. Entre semana pasó de 662 a 699 entradas (+6 %) al empezar las vacaciones SEP, y el 17-jul fue récord (827; +8 % ajustando por tendencia y día de la semana). En septiembre está estable (~645/día). El 16-sep cayó en miércoles y no hubo puente, así que no se esperaba una salida masiva. El sábado 12-sep (701) fue alto, pero dentro de lo normal.')]));
children.push(bullet([b('ZIBATA: '), r('el cambio fuerte fue la semana del 3 al 9 de agosto. Promedios semanales: 543 → 782 → 919 entradas/día. Entre semana subió 74 % y en fin de semana 30 %. Con el regreso a clases, entre semana subió otro 6 % (914 → 971) y el fin de semana bajó 5 % (723 → 687). La diferencia entre semana y fin de semana crece, como pasa en zonas con estudiantes o gente que va y viene a trabajar. Hipótesis: el semestre de la Anáhuac Querétaro (campus en Zibatá) empezó la 1.ª semana de agosto. Hay que descartarlo contra un cambio de sensor: el salto es brusco (lun 3-ago 660 → mar 4-ago 843) y también se ve en domingos (9-ago: 949).')]));
children.push(bullet([b('CHAPALITA: '), r('bajó en agosto (~65–100/día) y luego se recuperó; varía mucho. Con un contador dudoso (sección 5), no conviene sacar conclusiones de septiembre.')]));
children.push(bullet([b('Sucursales nuevas (3–14 sep): '), r('entre semana contra fin de semana, XOCHIMILCO 1,022 vs. 765, AVCONS 391 vs. 222 (−43 %) y SANJACINTO 224 vs. 198. En la semana 2 contra la semana 1, mismos días: XOCHIMILCO −14 %, AVCONS −9 %, SANJACINTO −1 %. Puede ser que el efecto de apertura se esté diluyendo; hacen falta 4–6 semanas para confirmarlo.')]));
children.push(bullet([b('15-sep (víspera del Grito y día de quincena; dato parcial hasta las 14 h): '), r('según el perfil horario, XOCHIMILCO iba para ~800–910 entradas, contra 1,121 de un martes promedio, y AVCONS para ~225–270 contra 341. SANJACINTO iba normal. Es una señal débil (un solo día, incompleto), pero vale la pena revisarla con el dato completo.')]));
children.push(bullet([b('Quincena: '), r('ajustando por tendencia y día de la semana, los días 14–16 y 29–1 no tienen más entradas que el resto (TAPO 1.00 vs. 1.00; ZIBATA 0.96 vs. 1.01). Si hay efecto, estaría en el monto de compra.')]));

children.push(H2('4.2 Contexto externo de septiembre 2026'));
children.push(table([1300, 3700, 4360], [
  ['Fecha', 'Evento (fuente)', 'Efecto esperado y qué revisar'],
  ['31-ago (lun)', 'Inicio del ciclo escolar 2026-2027: 185 días, calendario igual en todo el país, ~23 millones de alumnos (SEP).', 'Cambia la rutina de las mañanas y los mediodías. En las anteriores casi no se nota en el total diario (sección 4.1); puede notarse en la forma horaria, que aún no se puede medir.'],
  ['Septiembre', 'Suben las infecciones respiratorias agudas: en 2025, SINAVE registró ~216 mil casos por semana en jun–jul y más de 262 mil en septiembre (+21 %). La SEP alertó de un repunte al volver a clases.', 'Más demanda de antigripales, analgésicos y antibióticos. Si crecen las compras con receta, sube el tiempo de servicio (apoya H3). Revisar la mezcla de venta por categoría.'],
  ['15-sep (mar)', 'Quincena y víspera del Grito; día laboral y escolar normal.', 'Posible baja de tráfico en la tarde-noche por las celebraciones. Revisar las entradas por hora del 15-sep contra otros martes.'],
  ['16-sep (mié)', 'Descanso obligatorio (LFT, art. 74); no se recorre, así que no hay puente.', 'Día con patrón parecido a domingo. La Hoja 2 llega al 21-sep, así que ya contiene este día, pero sin desglose diario.'],
  ['19-sep (sáb), 12:00', 'Segundo Simulacro Nacional 2026, en todo el país, con alerta en celulares.', 'Posible caída puntual de 12 a 13 h y alertas atípicas en ese momento. Útil para confirmar que el sensor responde a eventos reales.'],
  ['Todo el mes', 'Temporada de lluvias: en CDMX las tormentas suelen caer en la tarde-noche (UNAM).', 'Menos gente caminando de 17 a 21 h en días de lluvia, y más llegadas juntas cuando para. Conviene cruzar con la lluvia diaria del SMN.'],
  ['Todo el mes', 'Dengue bajo en 2026: Jalisco con 163 casos a la semana 35 (−90 %); en Querétaro, 1 caso local y 1 defunción al 11-sep.', 'No se ve como motor de demanda este año, a diferencia de otros.'],
  ['Siguiente', 'Campaña invernal de vacunación (en 2025 empezó el 13-oct). Buen Fin del 13 al 17 de nov.', 'Próximos cambios de demanda: conviene tener listo el análisis por hora antes de octubre.'],
]));

// ===== 5. Calidad de datos
children.push(H1('5. Calidad de datos y límites'));
children.push(bullet([b('No se puede aislar el patrón horario de septiembre en las anteriores. '), r('La tabla por hora promedia todo el periodo (julio a septiembre). Para comparar las horas de septiembre contra las de julio y agosto hace falta el detalle hora por día.')]));
children.push(bullet([b('Muestra corta en las nuevas: '), r('12 días completos, o sea ~2 observaciones por día de la semana. Frases como “el lunes es el día más activo” todavía no son confiables.')]));
children.push(bullet([b('CHAPALITA probablemente cuenta de menos: '), r('121 entradas/día, pero con niveles de alertas parecidos a los demás. Eso da 10–14 alertas de unifila por cada 100 entradas, contra 2–4 en las otras sucursales. Además, su “pico” de entradas sale a las 22–23 h.')]));
children.push(bullet([b('Entradas y salidas no cuadran: '), r('salidas sobre entradas: TAPO +8 %, ZIBATA +9 %, CHAPALITA +16 %, AVCONS +6 %, SANJACINTO −10 %, XOCHIMILCO −13 %. Usar solo entradas y pedir que revisen los contadores.')]));
children.push(bullet([b('Madrugada: '), r('ZIBATA no registra entradas entre 0 y 7 h, pero sí ~0.6 alertas de mostrador por hora. ¿Atiende por ventanilla nocturna o son alertas falsas? XOCHIMILCO registra 27–32 entradas por hora entre 0 y 2 h en pocas noches; revisar si son reales.')]));
children.push(bullet([b('Días parciales excluidos: '), r('el primer día de cada sucursal y el 15-sep (datos hasta las 14 h). En el Excel, el 15-sep aparece con 0 para las anteriores: es un dato faltante, no un cero.')]));

// ===== 6. Próximos pasos
children.push(H1('6. Próximos pasos (en orden de prioridad)'));
children.push(num([b('Detalle hora por día de entradas y alertas. '), r('Permite comparar septiembre contra julio y agosto por hora y revisar el 15, 16 y 19 de septiembre. Es lo que más aporta con menos esfuerzo.')]));
children.push(num([b('Plantilla y horario de comidas del personal por sucursal y hora. '), r('Confirma o descarta H2 y la explicación del pico de mostrador a las 14–15 h. Posible piloto barato: escalonar comidas fuera de 14–16 h (por ejemplo 12:30–13:30 y 16:30–17:30) durante 2 semanas en XOCHIMILCO y SANJACINTO, y medir las alertas antes y después.')]));
children.push(num([b('Punto de venta por hora: '), r('tickets, ticket promedio, artículos, % con receta, % con antibiótico y tiempo de cobro. Sirve para probar H1 y H3 y para ver si la quincena mueve el gasto.')]));
children.push(num([b('Revisar los contadores: '), r('CHAPALITA, la diferencia entre entradas y salidas, y el salto de ZIBATA del 3–4 de agosto (¿cambio de sensor o puerta, o algo nuevo en la zona?).')]));
children.push(num([b('Entradas cada 15 minutos '), r('en TAPO y ZIBATA de 17 a 20 h, para probar H4 (llegadas en ráfaga a la salida del trabajo).')]));

// ===== 7. Qué se mantiene del documento anterior
children.push(H1('7. Qué se mantiene y qué se corrige del documento anterior'));
children.push(bullet([b('Se mantiene: '), r('el 17-jul sigue siendo el día récord de TAPO (827) y CHAPALITA (290). Ajustando por tendencia y día de la semana, fue +8 % en TAPO y +49 % en CHAPALITA.')]));
children.push(bullet([b('Se mantiene y se amplía: '), r('la hora de la comida (14–16 h) explica la franja alta de tráfico en TAPO y ZIBATA, y ahora también el pico de mostrador desatendido en las 6 sucursales.')]));
children.push(bullet([b('Se corrige: '), r('“pico de tráfico de 15:00–16:00 en AMBAS sucursales”. ZIBATA ahora tiene el pico a las 14 h, y en TAPO el pico es marginal dentro de una meseta de 15 a 19 h.')]));
children.push(bullet([b('Se replantea: '), r('la hipótesis de compra rápida contra compra planeada pasa a ser una de cuatro (sección 3.3). Además, el desfase entre tráfico y fila solo existe en ZIBATA, y en menor medida en SANJACINTO y AVCONS.')]));

// ===== Fuentes
children.push(H1('Fuentes'));
const src = [
  ['La Verdad Noticias (2026). Horarios de comida: ¿por qué en México se come tan tarde?', 'https://laverdadnoticias.com/pulso-viral/curiosidades-y-hacks/horarios-de-comida-mexico'],
  ['SEP (2026). Publica SEP Calendario Escolar 2026-2027 para Educación Básica y Normal.', 'https://educacionbasica.sep.gob.mx/publica-sep-calendario-escolar-2026-2027-para-educacion-basica-y-normal/'],
  ['El Financiero (30-ago-2026). Inicio del ciclo escolar 2026-2027: días feriados y puentes.', 'https://www.elfinanciero.com.mx/nacional/2026/08/30/inicio-del-ciclo-escolar-2026-2027-estos-son-los-dias-feriados-y-puentes-segun-el-calendario-de-la-sep/'],
  ['N+ (2026). Regreso a clases 2026: horarios de entrada y salida para alumnos de primaria SEP.', 'https://www.nmas.com.mx/sociedad/educacion/regreso-a-clases-2026-horarios-de-entrada-y-salida-para-alumnos-de-primaria-sep/'],
  ['INEGI (2018). Encuesta Origen-Destino en Hogares de la ZMVM 2017 – comunicado de resultados.', 'https://www.inegi.org.mx/contenidos/saladeprensa/boletines/2018/EstSociodemo/OrgenDest2018_02.pdf'],
  ['Gobierno de México. Atención médica en Unidades de Medicina Familiar del IMSS (horarios).', 'https://www.gob.mx/tramites/ficha/atencion-medica-en-unidades-de-medicina-familiar-del-imss/IMSS141'],
  ['adn40 (31-jul-2026). Desabasto de medicamentos en México.', 'https://www.adn40.mx/salud/2026-07-31/desabasto-de-medicamentos-en-mexico-provoca-avance-de-enfermedades-y-fallecimiento-de-pacientes/'],
  ['Cámara de Diputados (2011). COFEPRIS: necesario retener la receta médica de antibióticos en farmacias.', 'https://www3.diputados.gob.mx/camara/005_comunicacion/a_boletines/2011_2011/001_enero/15_15/2744_necesario_retener_la_receta_medica_de_antibioticos_en_farmacias_para_su_control_senala_cofepris_a_diputados_favor_de_utilizar_de_sabado_para_domingo'],
  ['Ley Federal del Trabajo, art. 63 (descanso en jornada continua).', 'https://sdv.com.mx/compendio/ley-federal-trabajo/articulo-63/'],
  ['Infobae (7-sep-2026). ¿Se recorre al lunes? Días festivos del 15 y 16 de septiembre.', 'https://www.infobae.com/mexico/2026/09/07/se-recorre-al-lunes-la-respuesta-de-la-sep-sobre-los-dias-festivos-por-el-15-y-16-de-septiembre/'],
  ['Expansión Política (19-sep-2026). ¿A qué hora será el simulacro nacional del 19 de septiembre de 2026?', 'https://politica.expansion.mx/cdmx/2026/09/19/a-que-hora-sera-simulacro-septiembre-2026'],
  ['Excélsior (2026). Más de 23 millones de estudiantes regresan a clases: medidas para prevenir enfermedades (datos SINAVE).', 'https://www.excelsior.com.mx/nacional/prevenir-enfermedades-y-contagios'],
  ['El Diario de Sonora (31-ago-2026). SEP alerta sobre repunte de infecciones al regreso a clases.', 'https://eldiariodesonora.com.mx/mexico/2026/08/31/sep-alerta-repunte-infecciones-prevencion-regreso-clases.html'],
  ['UNAM Global. ¿Está lloviendo más que antes en la CDMX? (lluvias vespertinas, isla de calor).', 'https://unamglobal.unam.mx/global_revista/lluvias-intensas-cdmx-causas-y-recomendaciones/'],
  ['Azteca Jalisco (2026). Dengue en Jalisco baja a 163 casos en 2026.', 'https://www.aztecajalisco.com/noticias-jalisco-local/dengue-en-jalisco-baja-a-163-casos-en-2026-permanece-un-caso-activo-de-sarampion/'],
  ['Rotativo (sep-2026). Querétaro confirma un caso y una muerte por dengue.', 'https://rotativo.com.mx/muerte-dengue-municipio-queretaro-septiembre-2026'],
  ['La Jornada (1-may-2026). Publican en DOF reforma para jornada laboral de 40 horas.', 'https://www.jornada.com.mx/noticia/2026/05/01/politica/publican-en-dof-reforma-para-jornada-laboral-de-40-horas'],
  ['Holland & Knight (feb-2026). Mexican Senate approves gradual reduction of the workweek to 40 hours.', 'https://www.hklaw.com/en/insights/publications/2026/02/mexican-senate-approves-gradual-reduction-of-the-workweek-to-40-hours'],
  ['Eick, Massey y Whitt (1993). Mt/G/∞ queues with sinusoidal arrival rates. Management Science 39(2).', 'https://pubsonline.informs.org/doi/10.1287/mnsc.39.2.241'],
  ['Green y Kolesar (1998). A note on approximating peak congestion in Mt/G/∞ queues with sinusoidal arrivals. Management Science 44(11).', 'https://pubsonline.informs.org/doi/10.1287/mnsc.44.11.S137'],
  ['Secretaría de Cultura – SIC. Universidad Anáhuac Querétaro (ubicación, El Marqués).', 'https://sic.cultura.gob.mx/ficha.php?table=universidad&table_id=802'],
  ['Universidad Anáhuac Querétaro. Costos periodo agosto–diciembre 2026.', 'https://www.anahuac.mx/queretaro/costos'],
  ['Merca20 (2026). Fiestas Patrias 2026: septiembre pone a prueba a los comercios antes del Buen Fin.', 'https://www.merca20.com/fiestas-patrias-2026-comercios-buen-fin/'],
  ['Publimetro (14-oct-2025). Arranca la campaña de vacunación invernal 2025–2026.', 'https://www.publimetro.com.mx/noticias/2025/10/14/arrancan-la-campana-de-vacunacion-invernal-20252026-contra-influenza-covid-19-y-neumococo/'],
];
src.forEach(([t, u]) => children.push(new Paragraph({ numbering: { reference: 'src', level: 0 }, spacing: { after: 60 },
  children: [new ExternalHyperlink({ link: u, children: [r(t, { style: 'Hyperlink', size: 17 })] }), r('  ' + new URL(u).hostname.replace(/^www\d?\./, ''), { size: 15, color: MUTED })] })));
children.push(note('Datos internos: Resumen_metricas.xlsx (Hoja 1: entradas y salidas, datos hasta 15-sep-2026 14:00; Hoja 2: alertas de mostrador desatendido y unifila, datos hasta 21-sep-2026 08:00).'));

const doc = new Document({
  creator: 'Patrick Joaquin Cruz',
  title: 'Picos por hora y septiembre 2026',
  styles: {
    default: { document: { run: { font: 'Arial', size: 20, color: INK } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { size: 40, bold: true, color: INK }, paragraph: { spacing: { after: 80 } } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, color: '1C5CAB' }, paragraph: { spacing: { before: 320, after: 140 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 23, bold: true, color: INK }, paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] },
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 300 } } } }] },
    { reference: 'src', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '[%1]', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 460 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [r('Picos por hora y septiembre 2026 · pág. ', { size: 16, color: MUTED }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: MUTED })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync('Picos_por_hora_y_septiembre_2026.docx', buf); console.log('written'); });
