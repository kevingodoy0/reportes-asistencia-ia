const ExcelJS = require('exceljs');

// --- FUNCIÓN 1: PARA FICHAS INDIVIDUALES DETALLADAS ---
async function generarFichaAlumno(datosAlumno, clases) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Ficha Individual');
    sheet.mergeCells('A1:E1');
    const titulo = sheet.getCell('A1');
    titulo.value = 'FICHA DETALLADA DEL ALUMNO';
    titulo.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    titulo.alignment = { horizontal: 'center' };
    titulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203764' } };

    sheet.getCell('A3').value = 'Nombre:';
    sheet.getCell('B3').value = datosAlumno.nombre;
    sheet.getCell('A4').value = 'Matrícula:';
    sheet.getCell('B4').value = datosAlumno.matricula;
    sheet.getCell('D3').value = 'Grupo:';
    sheet.getCell('E3').value = datosAlumno.grupo;
    sheet.getCell('D4').value = 'Dispositivo:';
    sheet.getCell('E4').value = datosAlumno.dispositivo;

    sheet.getRow(6).values = ['Clase', 'Horario', 'Asistencias', 'Faltas', '% Asistencia', 'Gráfica'];
    sheet.columns = [{ key: 'clase', width: 20 }, { key: 'hor', width: 15 }, { key: 'asis', width: 10 }, { key: 'fal', width: 10 }, { key: 'por', width: 12 }, { key: 'gra', width: 20 }];
    sheet.getRow(6).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };

    clases.forEach(c => {
        const total = c.asistencias + c.faltas;
        const porc = ((c.asistencias / total) * 100).toFixed(1) + '%';
        const barra = '█'.repeat(Math.round((c.asistencias / total) * 10)) + '░'.repeat(Math.round((c.faltas / total) * 10));
        const row = sheet.addRow([c.nombre, c.horario, c.asistencias, c.faltas, porc, barra]);
        row.getCell(6).font = { color: { argb: 'FF00B050' }, size: 14 };
    });

    await workbook.xlsx.writeFile(`./Reporte_1_Ficha_${datosAlumno.nombre.replace(' ', '_')}.xlsx`);
}

// --- FUNCIÓN 2: PARA REPORTES DE TABLAS (GRUPAL Y GENERAL) ---
async function crearReporteTabla(nombreArchivo, titulo, datos) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Asistencia');
    sheet.mergeCells('A1:G1');
    const header = sheet.getCell('A1');
    header.value = titulo.toUpperCase();
    header.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    header.alignment = { horizontal: 'center' };
    header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203764' } };

    sheet.getRow(3).values = ['Nombre', 'Grupo', 'Carrera', 'Asistencias', 'Faltas', '%', 'Gráfica'];
    sheet.columns = [{ width: 20 }, { width: 10 }, { width: 20 }, { width: 12 }, { width: 10 }, { width: 10 }, { width: 20 }];
    sheet.getRow(3).font = { bold: true };

    datos.forEach(d => {
        const total = d.asistencias + d.faltas;
        const porc = ((d.asistencias / total) * 100).toFixed(1) + '%';
        const barra = '█'.repeat(Math.round((d.asistencias / total) * 10)) + '░'.repeat(Math.round((d.faltas / total) * 10));
        const row = sheet.addRow([d.nombre, d.grupo, d.carrera, d.asistencias, d.faltas, porc, barra]);
        row.getCell(7).font = { color: { argb: 'FF00B050' }, size: 14 };
    });

    await workbook.xlsx.writeFile(`./${nombreArchivo}.xlsx`);
}

// --- EJECUCIÓN DE LOS 3 REPORTES AL MISMO TIEMPO ---
async function generarTodo() {
    console.log("Generando los 3 archivos solicitados...");

    // 1. FICHA INDIVIDUAL
    await generarFichaAlumno(
        { nombre: 'Kevin Godoy', matricula: 'TSU-IA-2024', grupo: 'IA-51', dispositivo: 'pba7D3' },
        [{ nombre: 'IA Aplicada', horario: '8-10am', asistencias: 19, faltas: 1 }]
    );

    // 2. REPORTE GRUPAL (5 Alumnos)
    const alumnosIA = [
        { nombre: 'Kevin Godoy', grupo: 'IA-51', carrera: 'IA', asistencias: 19, faltas: 1 },
        { nombre: 'Maria G', grupo: 'IA-51', carrera: 'IA', asistencias: 15, faltas: 5 },
        { nombre: 'Carlos R', grupo: 'IA-51', carrera: 'IA', asistencias: 10, faltas: 10 },
        { nombre: 'Ana B', grupo: 'IA-51', carrera: 'IA', asistencias: 20, faltas: 0 },
        { nombre: 'Luis T', grupo: 'IA-51', carrera: 'IA', asistencias: 12, faltas: 8 }
    ];
    await crearReporteTabla("Reporte_2_Grupal_IA51", "Reporte de Asistencia Grupo IA-51", alumnosIA);

    // 3. REPORTE GENERAL
    const carreras = [
        { nombre: 'Carrera IA', grupo: 'Varios', carrera: 'IA', asistencias: 500, faltas: 50 },
        { nombre: 'Carrera Mecatrónica', grupo: 'Varios', carrera: 'Meca', asistencias: 400, faltas: 100 }
    ];
    await crearReporteTabla("Reporte_3_General_Carreras", "Consolidado General de Carreras", carreras);

    console.log("✅ ¡Los 3 archivos han sido creados en tu carpeta!");
}

generarTodo();