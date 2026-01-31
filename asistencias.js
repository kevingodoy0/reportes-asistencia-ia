const ExcelJS = require('exceljs');
const axios = require('axios');

const ESTILOS = {
    header: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203764' } }, font: { color: { argb: 'FFFFFFFF' }, bold: true }, alignment: { horizontal: 'center' } },
    borde: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
};

// Datos del Horario Real UT Nayarit (IA-51)
const MATERIAS_IA51 = [
    { n: 'Ecuaciones Diferenciales', h: '07:00 - 09:30', a: 18, f: 2 },
    { n: 'Inglés V', h: '07:50 - 09:30', a: 15, f: 5 },
    { n: 'Aprendizaje de Máquinas', h: '07:00 - 09:30', a: 19, f: 1 },
    { n: 'Minería de Datos', h: '11:40 - 14:10', a: 17, f: 3 },
    { n: 'Fund. de Visión', h: '10:00 - 13:20', a: 16, f: 4 },
    { n: 'Proyecto Integrador II', h: '08:40 - 11:40', a: 20, f: 0 }
];

async function getChartBuffer(labels, data, titulo) {
    const config = {
        type: 'bar',
        data: { labels, datasets: [{ label: titulo, data, backgroundColor: '#28a745' }] }
    };
    const url = `https://quickchart.io/chart?c=${encodeURIComponent(JSON.stringify(config))}`;
    try {
        const res = await axios.get(url, { responseType: 'arraybuffer' });
        return res.data; // Buffer en RAM
    } catch (e) { return null; }
}

// 1. ARCHIVO INDIVIDUAL
async function generarIndividual() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Ficha Kevin');
    sheet.mergeCells('A1:E1');
    sheet.getCell('A1').value = 'CONTROL INDIVIDUAL - UT NAYARIT';
    sheet.getCell('A1').style = ESTILOS.header;

    sheet.addRow(['Alumno:', 'KEVIN GODOY', 'Matrícula:', 'TSU-IA-2024']);
    sheet.addRow(['Grupo:', 'IA-51', 'Dispositivo:', 'pba7D3']); //
    sheet.addRow([]);

    const hRow = sheet.addRow(['Materia', 'Horario', 'Asistencias', 'Faltas', '%']);
    hRow.eachCell(c => c.style = ESTILOS.header);

    MATERIAS_IA51.forEach(m => {
        const p = ((m.a / (m.a + m.f)) * 100).toFixed(1) + '%';
        const r = sheet.addRow([m.n, m.h, m.a, m.f, p]);
        r.eachCell(c => c.border = ESTILOS.borde);
    });

    const buffer = await getChartBuffer(MATERIAS_IA51.map(m => m.n), MATERIAS_IA51.map(m => m.a), 'Mis Asistencias');
    if (buffer) {
        const img = workbook.addImage({ buffer, extension: 'png' });
        sheet.addImage(img, { tl: { col: 0, row: 13 }, ext: { width: 450, height: 250 } });
    }
    await workbook.xlsx.writeFile('./1_Individual_Kevin.xlsx');
    console.log("✅ 1/3 Ficha Individual generada.");
}

// 2. ARCHIVO GRUPAL (Una tabla por materia en el mismo archivo)
async function generarGrupal() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('IA-51 Materias');
    let fila = 1;

    const alumnos = [{ n: 'Kevin G.', a: 18 }, { n: 'Maria G.', a: 15 }, { n: 'Carlos R.', a: 10 }, { n: 'Ana B.', a: 20 }];

    for (const mat of MATERIAS_IA51) {
        sheet.mergeCells(`A${fila}:D${fila}`);
        sheet.getCell(`A${fila}`).value = `MATERIA: ${mat.n} (${mat.h})`;
        sheet.getCell(`A${fila}`).style = ESTILOS.header;
        fila++;

        const h = sheet.addRow(['Alumno', 'Asistencias', 'Faltas', 'Estatus']);
        h.eachCell(c => c.style = { ...ESTILOS.header, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } } });
        fila++;

        alumnos.forEach(al => {
            const r = sheet.addRow([al.n, mat.a, mat.f, mat.a >= 15 ? 'OK' : 'ALERTA']);
            r.eachCell(c => c.border = ESTILOS.borde);
            fila++;
        });
        fila += 2; // Espacio
    }
    await workbook.xlsx.writeFile('./2_Grupal_IA51_Completo.xlsx');
    console.log("✅ 2/3 Reporte Grupal generado.");
}

// 3. ARCHIVO GENERAL
async function generarGeneral() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Carreras');
    sheet.mergeCells('A1:C1');
    sheet.getCell('A1').value = 'REPORTE GENERAL POR CARRERA';
    sheet.getCell('A1').style = ESTILOS.header;

    const carreras = [{ n: 'IA', a: 1200 }, { n: 'Meca', a: 980 }];
    sheet.getRow(3).values = ['Carrera', 'Asistencias', 'Estatus'];
    carreras.forEach(c => sheet.addRow([c.n, c.a, 'ÓPTIMO']));

    const buffer = await getChartBuffer(carreras.map(c => c.n), carreras.map(c => c.a), 'Total');
    if (buffer) {
        const img = workbook.addImage({ buffer, extension: 'png' });
        sheet.addImage(img, { tl: { col: 0, row: 6 }, ext: { width: 350, height: 200 } });
    }
    await workbook.xlsx.writeFile('./3_General_Carreras.xlsx');
    console.log("✅ 3/3 Reporte General generado.");
}

async function main() {
    console.log("🚀 Generando los 3 archivos finales...");
    await generarIndividual();
    await generarGrupal();
    await generarGeneral();
    console.log("\nProceso terminado. Revisa tu carpeta.");
}

main();