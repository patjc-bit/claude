// Arma el reporte en Word a partir de datos.json (salida de extract.py).
// Uso: node build_doc.js [datos.json] [salida.docx]
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, BorderStyle, TabStopType,
} = require('docx');

const [, , DATA = 'datos.json', OUT = 'Retiro_guardia_nocturno_alternativas.docx'] = process.argv;
const D = JSON.parse(fs.readFileSync(DATA, 'utf8'));
const A = D.alts, B = D.base, G = D.grupos;

// Formato del reporte anterior: Times New Roman, títulos azul marino, tablas con encabezado 1B3A5C y filas alternas EEF3F9
const NAVY = '1B3A5C', BLUE = '0072CE', ALT_FILL = 'EEF3F9';
const W = 9360;

// ---------- formatos de número
const n = (x) => Math.round(x).toLocaleString('en-US');
const $ = (x) => '$' + n(x);
const mill = (x) => '$' + (x / 1e6).toFixed(1) + ' millones';
const pct = (x, d = 1) => (x * 100).toFixed(d) + '%';
const veces = (x) => Math.round(x) + ' veces';

// ---------- bloques
const r = (text, o = {}) => new TextRun({ text, ...o });
const mark = (text) => r(text, { highlight: 'yellow' });
const P = (children, o = {}) => new Paragraph({
  children: (Array.isArray(children) ? children : [children]).map((c) => (typeof c === 'string' ? r(c) : c)),
  spacing: { after: 140 }, ...o,
});
const H1 = (t, o = {}) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [r(t)], ...o });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [r(t)] });
const Pk = (children) => P(children, { keepNext: true });
const gap = () => new Paragraph({ spacing: { after: 60 }, children: [] });
const total = (left, right) => new Paragraph({
  spacing: { after: 140 }, tabStops: [{ type: TabStopType.LEFT, position: 2600 }],
  children: [r(left, { bold: true }), r('\t' + right, { bold: true })],
});

// Primera columna a la izquierda y en negritas; las demás centradas. boldRows: índices de filas de datos en negritas (totales).
function table(widths, rows, { boldRows = [], keep = rows.length <= 12 } = {}) {
  const sum = widths.reduce((a, c) => a + c, 0);
  if (sum !== W) throw new Error(`Anchos suman ${sum}, se esperaba ${W}`);
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'auto' };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: widths,
    borders: { ...borders, insideHorizontal: border, insideVertical: border },
    rows: rows.map((row, i) => {
      const head = i === 0;
      const fill = head ? NAVY : (i % 2 === 0 ? ALT_FILL : null);
      const bold = head || boldRows.includes(i - 1);
      return new TableRow({
        tableHeader: head,
        cantSplit: !head,
        children: row.map((cell, j) => new TableCell({
          width: { size: widths[j], type: WidthType.DXA },
          shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [new Paragraph({
            alignment: j === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
            keepNext: head || (keep && i < rows.length - 1),
            children: [r(String(cell), { bold: bold || j === 0, color: head ? 'FFFFFF' : '000000', size: 18 })],
          })],
        })),
      });
    }),
  });
}

const c = [];
const SV = D.sin_ventanilla;
const g2 = G.solo2;

// ===== 1. Resumen
c.push(H1('1. Resumen'));
c.push(P(`El turno nocturno concentra ${B.nocturnos} de los ${B.intramuros} guardias intramuros y cuesta ${mill(B.costo_noc)} al año. En él ocurrió el 10% de los robos de enero a julio de 2026, y 99 de sus 102 robos sucedieron entre las 22:00 y las 22:59. De las 23:00 a las 07:00 se registraron 3 robos en siete meses en toda la red.`));
c.push(P(`Se proponen dos alternativas para retirar el guardia nocturno. Las dos lo conservan en farmacias con robo organizado de noche y difieren en cómo identifican el riesgo de violencia en la CDMX: la alternativa 1 usa la alcaldía y la alternativa 2 la colonia. La alternativa 1 retira ${A[1].farmacias} guardias y ahorra ${mill(A[1].ahorro)} al año, el ${pct(A[1].pct)} del gasto actual. La alternativa 2 retira ${A[2].farmacias}, adelanta la ventanilla a las 22:00 en tres farmacias y ahorra ${mill(A[2].ahorro)} al año (${pct(A[2].pct)}). Los turnos matutino y vespertino no se modifican.`));

// ===== 2. Situación actual
c.push(H1('2. Situación actual'));
c.push(table([6360, 3000], [
  ['Concepto', 'Cantidad'],
  ['Farmacias con guardia', B.farmacias],
  ['Farmacias en la CDMX', B.cdmx],
  ['Farmacias foráneas', B.foraneas],
  ['Guardias totales', B.guardias],
  ['Guardias intramuros', B.intramuros],
  ['Guardias armados', B.armados],
  ['Guardias en turno nocturno', B.nocturnos],
  ['Costo anual', $(B.costo)],
  ['Costo anual del turno nocturno', $(B.costo_noc)],
]));
c.push(gap());
const [ht, sep] = B.empresas;
c.push(Pk(`Los guardias intramuros los prestan dos empresas: HT SICHERHEIT tiene ${ht.guardias} y SEPSISA ${sep.guardias}. Los ${B.armados} guardias armados se cobran a la tarifa de MACOTELA.`));
c.push(table([3360, 1800, 2200, 2000], [
  ['Empresa', 'Guardias', 'Costo mensual por guardia', 'Costo anual'],
  ...B.empresas.map((e) => [e.empresa, e.guardias, '$' + e.mensual.toLocaleString('en-US', { minimumFractionDigits: 2 }), $(e.anual)]),
  ['Total', B.guardias, '', $(B.costo)],
], { boldRows: [3] }));
c.push(gap());
c.push(P('De enero a julio de 2026 se cubrió el 89% de los turnos que pide el Semáforo y 10,794 turnos quedaron sin guardia. En el mismo periodo se registraron 1,056 robos por $6.89 millones, equivalentes a unos $11.9 millones al año. El gasto en guardias es 11.8 veces el monto robado.'));

// ===== 3. Justificación
c.push(H1('3. Justificación'));
c.push(H2('3.1 El gasto del turno nocturno no corresponde a su robo'));
c.push(table([3660, 1900, 1900, 1900], [
  ['Turno', 'Guardias intramuros', '% de guardias', '% de robos'],
  ['Matutino (07:00 a 14:30)', 121, '24%', '39%'],
  ['Vespertino (14:30 a 22:00)', 188, '38%', '52%'],
  ['Nocturno (22:00 a 07:00)', B.nocturnos, '38%', '10%'],
]));
c.push(gap());
c.push(P('El turno nocturno tiene prácticamente los mismos guardias que el vespertino, pero una quinta parte de sus robos. Además, el robo nocturno se concentra en la primera hora del turno: de los 102 robos nocturnos del periodo, 99 ocurrieron entre las 22:00 y las 22:59.'));
c.push(P('El robo organizado sigue el mismo patrón. De enero de 2025 a septiembre de 2026, 150 de los 1,361 robos organizados (11%) ocurrieron en el turno nocturno y 750 (55%) en el vespertino.'));

c.push(H2('3.2 Las farmacias con venta nocturna atienden por ventanilla'));
c.push(P('Después del cierre del piso de venta, las farmacias con venta nocturna atienden por ventanilla. El cliente no tiene acceso al anaquel, por lo que el robo que registra la base, que ocurre en piso de venta, no es posible en esas horas. Esto explica que de las 23:00 a las 07:00 solo hubo 3 robos en siete meses en toda la red.'));
c.push(P('El riesgo nocturno está en la hora previa al cambio a ventanilla. A esa hora se dirigen los criterios del punto 3.5 y el adelanto de ventanilla de la alternativa 2.'));

c.push(H2('3.3 El turno nocturno es el de menor cobertura'));
c.push(P(['La cobertura general de los turnos es de 89%. En el turno nocturno es cercana al ', mark('80%'), ', es decir, una de cada cinco noches el puesto queda sin guardia. Es el turno que el proveedor tiene más dificultad para cubrir.']));

c.push(H2('3.4 La presencia del guardia reduce poco el robo'));
c.push(P('El 53% de los robos del periodo ocurrió con un guardia presente y el 85% en horas en que el Semáforo ya pide guardia. Con guardia en piso, el robo por hora es entre 16% y 25% menor que sin guardia.'));
c.push(P('El guardia contiene una parte del robo, pero no lo evita. En un turno con poco robo, como el nocturno, ese efecto representa muy pocos eventos.'));

c.push(H2('3.5 Criterios para conservar el guardia nocturno'));
c.push(Pk('El retiro no es general. En cada farmacia se revisa si existe alguna condición que justifique mantener el guardia nocturno. Las dos alternativas aplican los mismos criterios en las farmacias foráneas y difieren en la CDMX.'));
c.push(table([1700, 3260, 2200, 2200], [
  ['Criterio', 'Qué indica', 'Alternativa 1', 'Alternativa 2'],
  ['Venta nocturna', 'La farmacia opera de noche por ventanilla', 'Conserva', 'Conserva solo en foráneas'],
  ['Iztapalapa o Gustavo A. Madero', 'Concentran el 19% y el 12% del robo a negocio con violencia de la CDMX (FGJ)', 'Conserva', 'Se sustituye por el criterio de colonia'],
  ['Colonia con robo con violencia alto', '21 farmacias de la CDMX en las colonias con más carpetas de robo a negocio con violencia (FGJ, 2021 a julio de 2024)', 'No se aplica', 'Conserva'],
  ['Robo organizado nocturno', 'Robos de bandas o de farderos individuales de monto alto en el turno nocturno, enero de 2025 a septiembre de 2026', 'Conserva con 2 o más', 'Conserva con 3 o más en la CDMX y 2 o más en foráneas'],
  ['Robo alto entre 22:00 y 22:59', 'Farmacias de la CDMX con venta nocturna y 3 o más robos a esa hora, enero de 2025 a septiembre de 2026', 'No se aplica', 'Se retira y se adelanta la ventanilla a las 22:00'],
  ['Horario distinto al Semáforo', `${D.revision_horario} farmacias cuyo horario implementado no coincide con el Semáforo`, 'Sin cambio hasta revisarse', 'Sin cambio hasta revisarse'],
]));
c.push(gap());
c.push(P('Una farmacia cumple el criterio de colonia si su colonia está en el 20% más alto de carpetas de robo a negocio con violencia entre las farmacias de la CDMX, en total o de 20:00 a 23:59 con un mínimo de 3 carpetas. Cumplen 21 farmacias. F119 Stim queda en revisión porque su colonia no tiene coincidencia en la base de la FGJ, y mientras tanto conserva el guardia en la alternativa 2.'));
c.push(P('Se considera organizado el robo de una banda o el de un fardero individual de $14,031 o más, monto que separa el 10% de robos más caros. De enero de 2025 a septiembre de 2026, el robo organizado representa el 51% de los robos y el 81% del monto. Un robo de banda promedia $8,876 y uno ocasional $1,891. Por eso pesa más en la decisión que el número total de robos.'));
c.push(P('Las farmacias foráneas tienen la misma regla en las dos alternativas, porque fuera de la CDMX no se cuenta con información de violencia por colonia.'));

// ===== 4. Análisis financiero
c.push(H1('4. Análisis financiero'));
c.push(P('Los montos están calculados a costo de contrato del Semáforo y no incluyen costos de transición. La alternativa 1 corresponde al escenario conservador de la simulación y la alternativa 2 al escenario moderado.'));
c.push(H2('4.1 Ahorro por alternativa'));
c.push(table([2760, 1300, 2000, 2000, 1300], [
  ['Escenario', 'Guardias', 'Costo anual', 'Ahorro anual', '%'],
  ['Semáforo actual', B.guardias, $(B.costo), 'N/a', 'N/a'],
  ...[1, 2].map((k) => [`Alternativa ${k}`, A[k].guardias, $(A[k].costo), $(A[k].ahorro), pct(A[k].pct)]),
]));
c.push(gap());
c.push(table([4160, 2600, 2600], [
  ['Concepto', 'Alternativa 1', 'Alternativa 2'],
  ['Guardias nocturnos retirados', A[1].farmacias, A[2].farmacias],
  ['En farmacias foráneas', A[1].for, A[2].for],
  ['En farmacias de la CDMX', A[1].cdmx, A[2].cdmx],
  ['De HT SICHERHEIT', A[1].ht, A[2].ht],
  ['De SEPSISA', A[1].sep, A[2].sep],
  ['Ahorro en foráneas', $(A[1].ahorro_for), $(A[2].ahorro_for)],
  ['Ahorro en la CDMX', $(A[1].ahorro_cdmx), $(A[2].ahorro_cdmx)],
  ['Guardias nocturnos que se conservan', A[1].noc_conservan, A[2].noc_conservan],
]));
c.push(gap());
c.push(P(`Las dos alternativas retiran los mismos ${A[1].for} guardias en farmacias foráneas, por ${mill(A[1].ahorro_for)}. La diferencia está en la CDMX, donde todos los guardias que se retiran son de HT SICHERHEIT.`));
const pv = D.previa;
c.push(P(`La estimación previa del proyecto era un ahorro de ${mill(pv.estimacion)} al quitar el nocturno salvo en farmacias con venta nocturna, en Iztapalapa y en Gustavo A. Madero. La alternativa 1 ahorra ${mill(pv.estimacion - A[1].ahorro)} menos porque además conserva el guardia por robo organizado (${mill(pv.bandas)}) y no modifica las farmacias en revisión de horario (${mill(pv.horario)}).`));

c.push(H2('4.2 Robo registrado en las farmacias que pierden el nocturno'));
const red = D.red;
c.push(table([5160, 2100, 2100], [
  ['Indicador', 'Alternativa 1', 'Alternativa 2'],
  ['Farmacias sin guardia nocturno', A[1].farmacias, A[2].farmacias],
  ['Robos en turno nocturno, enero a julio de 2026', A[1].rob_noc, A[2].rob_noc],
  ['Monto robado en turno nocturno, enero a julio de 2026', $(A[1].monto_noc), $(A[2].monto_noc)],
  [`% de los ${red.rob_noc} robos nocturnos de la red, enero a julio de 2026`, ...[1, 2].map((k) => pct(A[k].rob_noc / red.rob_noc, 0))],
  ['Robos entre 22:00 y 22:59, enero a julio de 2026', A[1].rob_22, A[2].rob_22],
  ['Robos entre 22:00 y 22:59, enero de 2025 a septiembre de 2026', A[1].rob_22_completo, A[2].rob_22_completo],
  [`% de los ${red.rob_22_completo} robos de esa hora en la red, enero de 2025 a septiembre de 2026`, ...[1, 2].map((k) => pct(A[k].rob_22_completo / red.rob_22_completo, 0))],
  ['Ahorro de enero a julio de 2026', $(A[1].ahorro_periodo), $(A[2].ahorro_periodo)],
]));
c.push(gap());
c.push(P(`En la alternativa 1, el ${pct(1 - A[1].rob_noc / red.rob_noc, 0)} del robo nocturno de la red queda en farmacias que conservan su guardia. Las ${A[1].farmacias} farmacias de esta alternativa registraron ${A[1].rob_noc} robos nocturnos en siete meses, todos entre las 22:00 y las 22:59. El ahorro del mismo periodo equivale a ${veces(A[1].veces)} el monto robado en esos turnos.`));
c.push(P(`La alternativa 2 retira además el guardia en ${g2.n} farmacias de la CDMX con más robo a esa hora. Esas farmacias registraron ${g2.rob_noc_completo} robos nocturnos de enero de 2025 a septiembre de 2026, por ${$(g2.monto_noc_completo)}, y ${g2.rob_22_completo} de ellos ocurrieron entre las 22:00 y las 22:59. Con la alternativa 2, el ${pct(A[2].rob_noc / red.rob_noc, 0)} del robo nocturno de la red queda en farmacias sin guardia nocturno y el ahorro del periodo equivale a ${veces(A[2].veces)} el monto robado.`));
c.push(P('Estas cifras corresponden a robos observados y no son una proyección.'));

c.push(H2('4.3 Sensibilidad de la alternativa 2'));
const S = D.sens;
c.push(Pk('El ahorro de la alternativa 2 depende de los umbrales que se aplican en la CDMX y del adelanto de ventanilla en las farmacias con robo alto a las 22:00.'));
c.push(table([6360, 3000], [
  ['Supuesto', 'Ahorro anual'],
  ['Criterios de este reporte', $(A[2].ahorro)],
  ['Robo organizado nocturno: 2 o más en vez de 3', $(S.bandas2)],
  ['Sin adelanto de ventanilla: las farmacias con robo alto a las 22:00 conservan el guardia', $(S.sin_ventanilla['3'])],
  ['Sin adelanto de ventanilla y robo alto desde 2 robos', $(S.sin_ventanilla['2'])],
  ['Sin adelanto de ventanilla y robo alto desde 4 robos', $(S.sin_ventanilla['4'])],
]));
c.push(gap());
c.push(P(`Con el adelanto de ventanilla, el umbral de robo a las 22:00 no cambia el ahorro, solo el número de farmacias en las que se adelanta: ${S.ventanillas['2']} con 2 o más robos, ${S.ventanillas['3']} con 3 o más y ${S.ventanillas['4']} con 4 o más.`));
c.push(P('La opción de reducir a un guardia los turnos de día con dos o más guardias no cambia el resultado. En la CDMX, solo F1 Centro tiene dos o más guardias en un turno de día, y los conserva por estar en una colonia con robo con violencia alto.'));

// ===== 5. Alternativas
c.push(H1('5. Alternativas de retiro'));
c.push(Pk(`Las dos alternativas conservan los turnos matutino y vespertino, los ${B.armados} guardias armados y el turno nocturno de las ${D.revision_horario} farmacias en revisión de horario. En las ${B.foraneas} farmacias foráneas aplican la misma regla y se diferencian en la CDMX.`));
c.push(table([2300, 4060, 1300, 1700], [
  ['Alternativa', 'Se conserva el nocturno en la CDMX por', 'Guardias retirados', 'Ahorro anual'],
  ['1. Conservadora', 'Venta nocturna, Iztapalapa o Gustavo A. Madero, o 2 o más robos organizados nocturnos', A[1].farmacias, $(A[1].ahorro)],
  ['2. Moderada', 'Colonia con robo con violencia alto o 3 o más robos organizados nocturnos', A[2].farmacias, $(A[2].ahorro)],
]));
c.push(gap());

// 5.1
const c1 = A[1].conservan;
c.push(H2('5.1 Alternativa 1. Conservadora'));
c.push(P('Se retira el guardia nocturno en todas las farmacias, salvo en las que operan con venta nocturna, están en Iztapalapa o Gustavo A. Madero o tienen dos o más robos organizados nocturnos de enero de 2025 a septiembre de 2026.'));
c.push(P(`Comprende ${A[1].farmacias} farmacias, ${A[1].for} foráneas y ${A[1].cdmx} de la CDMX. La lista está en los anexos 1 y 2. Conservan el guardia nocturno ${c1.total} farmacias: ${c1.venta} por venta nocturna, ${c1.alcaldia} por alcaldía, ${c1.bandas} por robo organizado y ${c1.horario} en revisión de horario.`));
const s1 = D.solo1.map((f) => `${f.clave} ${f.sucursal}`);
const v1 = D.variante1;
c.push(P(`Esta alternativa no aplica el criterio de colonia. Por eso cinco farmacias de la CDMX en colonias con robo con violencia alto pierden el guardia nocturno: ${s1.slice(0, -1).join(', ')} y ${s1[s1.length - 1]}. En la alternativa 2 lo conservan. Si se agregan como excepción, la alternativa 1 queda en ${v1.guardias} guardias y un ahorro de ${$(v1.ahorro)} al año (${pct(v1.pct)}).`));
c.push(total(`${A[1].farmacias} guardias retirados`, `Ahorro anual: ${$(A[1].ahorro)}`));

// 5.2
const c2 = A[2].conservan;
c.push(H2('5.2 Alternativa 2. Moderada'));
c.push(Pk('En las farmacias foráneas se aplica la regla de la alternativa 1. En la CDMX se revisan las condiciones en el orden de la tabla y se aplica la primera que se cumple.'));
c.push(table([900, 5260, 3200], [
  ['Orden', 'Condición en la CDMX', 'Resultado'],
  ['1', 'Colonia con robo con violencia alto o en revisión', 'Conserva'],
  ['2', '3 o más robos organizados nocturnos', 'Conserva'],
  ['3', 'Horario distinto al Semáforo', 'Sin cambio hasta revisarse'],
  ['4', 'Venta nocturna y 3 o más robos entre 22:00 y 22:59', 'Se retira y se adelanta la ventanilla a las 22:00'],
  ['5', 'Cualquier otro caso', 'Se retira'],
]));
c.push(gap());
c.push(P(`Comprende ${A[2].farmacias} farmacias, ${A[2].for} foráneas y ${A[2].cdmx} de la CDMX. Respecto a la alternativa 1, agrega ${g2.n} farmacias de la CDMX que conservaban el guardia por venta nocturna, alcaldía o dos robos organizados, y lo mantiene en las cinco farmacias en colonia con violencia. La lista está en los anexos 1 y 3.`));
c.push(P(`Conservan el guardia nocturno ${c2.total} farmacias. En la CDMX son ${c2.cdmx}: ${c2.colonia} por colonia con violencia, ${c2.colonia_rev} con colonia en revisión, ${c2.bandas_cdmx} por robo organizado y ${c2.horario_cdmx} en revisión de horario. En las foráneas son ${c2.for}: ${c2.venta_for} por venta nocturna, ${c2.bandas_for} por robo organizado y ${c2.horario_for} en revisión de horario.`));
c.push(Pk('En las tres farmacias con venta nocturna y tres o más robos entre las 22:00 y las 22:59 se retira el guardia y se adelanta el cambio a ventanilla a las 22:00, la hora en que se concentra el robo nocturno.'));
c.push(table([1000, 2000, 2200, 2160, 2000], [
  ['Clave', 'Sucursal', 'Alcaldía', 'Robos entre 22:00 y 22:59 (enero de 2025 a septiembre de 2026)', 'Robos entre 22:00 y 22:59 (enero a julio de 2026)'],
  ...D.ventanilla.map((f) => [f.clave, f.sucursal, f.alcaldia, f.rob22_completo, f.rob22_2026]),
]));
c.push(gap());
c.push(P(`Requiere confirmar con Operaciones que estas farmacias pueden atender por ventanilla desde las 22:00. Si no es posible, conservan el guardia y el ahorro de la alternativa 2 baja a ${$(SV.ahorro)} al año (${pct(SV.pct)}).`));
c.push(total(`${A[2].farmacias} guardias retirados`, `Ahorro anual: ${$(A[2].ahorro)}`));

// 5.3
const DIAS = 212 / 365;
const paso = (de, a) => {
  const ah = a.ahorro - (de ? de.ahorro : 0);
  const mo = a.monto_noc - (de ? de.monto_noc : 0);
  return { g: a.farmacias - (de ? de.farmacias : 0), ah, rob: a.rob_noc - (de ? de.rob_noc : 0), mo, v: ah * DIAS / mo };
};
const p01 = paso(null, A[1]), p12 = paso(A[1], A[2]);
c.push(H2('5.3 Diferencias entre alternativas'));
c.push(table([2060, 1300, 1700, 1500, 1500, 1300], [
  ['Cambio', 'Guardias retirados adicionales', 'Ahorro anual adicional', 'Robos nocturnos adicionales (enero a julio de 2026)', 'Monto robado adicional (enero a julio de 2026)', 'Ahorro del periodo frente al monto robado'],
  ['Del Semáforo a la alternativa 1', p01.g, $(p01.ah), p01.rob, $(p01.mo), veces(p01.v)],
  ['De la alternativa 1 a la 2', p12.g, $(p12.ah), p12.rob, $(p12.mo), veces(p12.v)],
]));
c.push(gap());
c.push(P(`El paso de la alternativa 1 a la 2 retira ${g2.n} guardias más y conserva los ${D.solo1.length} de las farmacias en colonia con violencia. Agrega ${mill(p12.ah)} al año con ${p12.rob} robos nocturnos más en siete meses. En los dos casos el ahorro supera al monto robado en los turnos que se retiran, pero la relación baja de ${Math.round(p01.v)} a ${Math.round(p12.v)} veces.`));

// 5.4
c.push(H2('5.4 Seguimiento y regla de regreso'));
c.push(P(['En cualquiera de las dos alternativas se propone un periodo de observación de ', mark('tres'), ' meses. El robo de las farmacias intervenidas se compara contra farmacias similares que conservan el guardia y no solo contra su propio histórico, porque el robo en la CDMX viene al alza desde mayo de 2026.']));
c.push(P(['El guardia regresa a una farmacia de inmediato ante cualquier incidente con violencia, o cuando registre ', mark('dos'), ' robos nocturnos en un mismo mes. En la alternativa 2 también se revisa que las tres farmacias con ventanilla a las 22:00 cumplan el nuevo horario.']));

// ===== 6. Consideraciones
c.push(H1('6. Consideraciones'));
c.push(P('Los ahorros son a costo de contrato y no incluyen costos de transición ni ajustes que pudieran aplicar con el proveedor.'));
c.push(P('La información de violencia por colonia proviene de la FGJ y llega a julio de 2024. 20 farmacias se asignaron a una colonia aproximada y están en revisión. Cinco de ellas cumplen el criterio de colonia: F11, F15, F81, F97 y F106. F11 y F15 se asignaron a Lindavista Sur y falta confirmar si corresponden a Lindavista Norte.'));
c.push(P(`Las ${D.revision_horario} farmacias con horario distinto al Semáforo no se modifican hasta revisarse. Septiembre de 2026 está incompleto en la base de robos y la asistencia de armados llega a julio de 2026.`));

// ===== Anexos
c.push(H1('Anexo 1. Farmacias sin guardia nocturno en las dos alternativas', { pageBreakBefore: true }));
c.push(table([1000, 2860, 2600, 1100, 1800], [
  ['Clave', 'Sucursal', 'Alcaldía o municipio', 'Formato', 'Empresa'],
  ...D.comunes.map((f) => [f.clave, f.sucursal, f.alcaldia, f.formato, f.empresa]),
]));
c.push(gap());

c.push(H1('Anexo 2. Farmacias sin guardia nocturno solo en la alternativa 1'));
c.push(table([800, 1700, 1800, 900, 1600, 2560], [
  ['Clave', 'Sucursal', 'Alcaldía', 'Formato', 'Empresa', 'Motivo para conservar en la alternativa 2'],
  ...D.solo1.map((f) => [f.clave, f.sucursal, f.alcaldia, f.formato, f.empresa, 'Colonia con robo con violencia alto']),
]));
c.push(gap());

c.push(H1('Anexo 3. Farmacias sin guardia nocturno solo en la alternativa 2'));
if (D.solo2.some((f) => f.empresa !== 'HT SICHERHEIT')) throw new Error('Anexo 3: hay farmacias que no son de HT SICHERHEIT');
c.push(Pk(`Las ${D.solo2.length} farmacias son de la CDMX y su guardia es de HT SICHERHEIT. La tabla indica por qué cada una conserva el guardia en la alternativa 1, sus robos entre las 22:00 y las 22:59 de enero de 2025 a septiembre de 2026 y si se propone adelantar su ventanilla a las 22:00.`));
c.push(table([750, 1550, 1700, 1000, 2060, 1050, 1250], [
  ['Clave', 'Sucursal', 'Alcaldía', 'Formato', 'Motivo en la alternativa 1', 'Robos 22:00 a 22:59', 'Ventanilla a las 22:00'],
  ...D.solo2.map((f) => [f.clave, f.sucursal, f.alcaldia, f.formato, f.motivo1, f.rob22_completo, f.ventanilla]),
]));
c.push(gap());

// ---------- documento
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Times New Roman', size: 20, language: { value: 'es-MX' } } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: 'Times New Roman', size: 28, bold: true, color: NAVY },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: 'Times New Roman', size: 23, bold: true, color: BLUE },
        paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    children: c,
  }],
});
// docx escribe w:highlightCs junto a w:highlight; no está en el esquema de Word y se quita.
const JSZip = require(require.resolve('jszip', { paths: [require.resolve('docx')] }));
Packer.toBuffer(doc)
  .then((buf) => JSZip.loadAsync(buf))
  .then(async (zip) => {
    const xml = await zip.file('word/document.xml').async('string');
    zip.file('word/document.xml', xml.replace(/<w:highlightCs [^>]*\/>/g, ''));
    return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  })
  .then((buf) => { fs.writeFileSync(OUT, buf); console.log('ok', OUT); });
