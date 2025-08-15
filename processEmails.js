const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const chalk = require('chalk');

// Configuración
const INPUT_DIR = path.join(__dirname, 'input');
const OUTPUT_DIR = path.join(__dirname, 'output');
const INPUT_FILE_NAME = 'correos.xlsx';
const OUTPUT_FILE_NAME = 'processed_emails.json';

const COLUMN_NAMES = {
    CURRENT_EMAIL: 'current_email', // o 'old_email', 'email_actual', etc.
    NEW_EMAIL: 'new_email'     // o 'nuevo_email', 'email_nuevo', etc.
};

function processExcelFile() {
    try {
        console.log(chalk.blue('🔍 Buscando archivo Excel...'));
        
        // Verificar directorio de input
        if (!fs.existsSync(INPUT_DIR)) {
            fs.mkdirSync(INPUT_DIR, { recursive: true });
            console.log(chalk.yellow(`Se creó el directorio 'input'. Por favor coloca tu archivo ${INPUT_FILE_NAME} allí`));
            return;
        }

        const inputPath = path.join(INPUT_DIR, INPUT_FILE_NAME);
        
        // Verificar archivo Excel
        if (!fs.existsSync(inputPath)) {
            console.log(chalk.red(`No se encontró el archivo ${INPUT_FILE_NAME} en el directorio 'input'`));
            console.log(chalk.yellow(`Por favor coloca tu archivo Excel con los correos en /input/${INPUT_FILE_NAME}`));
            return;
        }

        console.log(chalk.green(`✔ Archivo encontrado: ${inputPath}`));
        
        // Leer el archivo Excel
        const workbook = xlsx.readFile(inputPath);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 'A' });
        
        console.log(chalk.blue('🔄 Procesando datos...'));
        
        const headerRow = getHeaderRow(worksheet);
        const columnIndices = findColumnIndices(headerRow);
        
        if (!columnIndices.current || !columnIndices.new) {
            console.log(chalk.red('No se encontraron las columnas requeridas en el archivo Excel'));
            console.log(chalk.yellow(`Asegúrate que existan columnas llamadas '${COLUMN_NAMES.CURRENT_EMAIL}' y '${COLUMN_NAMES.NEW_EMAIL}'`));
            return;
        }

        // Procesar los datos
        const users = [];
        let skippedRows = 0;
        
        jsonData.forEach((row, index) => {
            // Saltar fila de encabezado
            if (index === 0) return;
            
            const current_email = row[columnIndices.current] ? String(row[columnIndices.current]).trim() : null;
            const new_email = row[columnIndices.new] ? String(row[columnIndices.new]).trim() : null;
            
            if (!current_email || !new_email) {
                console.log(chalk.yellow(` Fila ${index + 1} ignorada - datos faltantes`));
                skippedRows++;
                return;
            }
            
            if (!isValidEmail(current_email)) {
                console.log(chalk.yellow(`Fila ${index + 1} ignorada - correo actual inválido: ${current_email}`));
                skippedRows++;
                return;
            }
            
            if (!isValidEmail(new_email)) {
                console.log(chalk.yellow(`Fila ${index + 1} ignorada - nuevo correo inválido: ${new_email}`));
                skippedRows++;
                return;
            }
            
            users.push({ current_email, new_email });
        });
        
        const result = { users };
        
        if (!fs.existsSync(OUTPUT_DIR)) {
            fs.mkdirSync(OUTPUT_DIR, { recursive: true });
        }
        
        // Guardar el resultado
        const outputPath = path.join(OUTPUT_DIR, OUTPUT_FILE_NAME);
        fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
        
        // Mostrar resultados
        console.log(chalk.green('\nProceso completado con éxito!'));
        console.log(chalk.cyan(`Total de pares procesados: ${users.length}`));
        console.log(chalk.yellow(`Total de filas omitidas: ${skippedRows}`));
        console.log(chalk.green(`Resultado guardado en: ${outputPath}`));
        
    } catch (error) {
        console.log(chalk.red('Error durante el procesamiento:'));
        console.log(chalk.red(error.message));
        process.exit(1);
    }
}

// Obtener la fila de encabezados
function getHeaderRow(worksheet) {
    const headers = {};
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    
    for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = xlsx.utils.encode_cell({ r: range.s.r, c: col });
        const cell = worksheet[cellAddress];
        headers[col] = cell ? cell.v : null;
    }
    
    return headers;
}

// Buscar índices de columnas por nombre
function findColumnIndices(headerRow) {
    const indices = { current: null, new: null };
    
    Object.entries(headerRow).forEach(([colIndex, headerValue]) => {
        const header = String(headerValue).toLowerCase().trim();
        
        if (header.includes(COLUMN_NAMES.CURRENT_EMAIL.toLowerCase())) {
            indices.current = colIndex;
        }
        
        if (header.includes(COLUMN_NAMES.NEW_EMAIL.toLowerCase())) {
            indices.new = colIndex;
        }
    });
    
    return indices;
}

function isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// Ejecutar el proceso
processExcelFile();