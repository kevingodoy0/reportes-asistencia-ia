// Usamos require porque tu compañero pidió CommonJS
const ExcelJS = require('exceljs');

async function crearPrueba() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Mi Primer Reporte');

    // Creamos las columnas
    sheet.columns = [
        { header: 'ID', key: 'id' },
        { header: 'Nombre', key: 'nom' },
        { header: 'Carrera', key: 'carr' }
    ];

    // Agregamos una fila de datos
    sheet.addRow({ id: 1, nom: 'Tu Nombre', carr: 'Inteligencia Artificial' });

    // Guardamos el archivo
    await workbook.xlsx.writeFile('Prueba.xlsx');
    console.log('¡Archivo Excel creado con éxito!');
}

crearPrueba();