const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const chalk = require('chalk');

const COLUMN_NAMES = {
    CURRENT_EMAIL: 'current_email',
    NEW_EMAIL: 'new_email'
};

const INPUT_DIR = path.join(__dirname, 'input');
const OUTPUT_DIR = path.join(__dirname, 'output');
const INPUT_FILE_NAME = 'correos.xlsx';
const OUTPUT_FILE_NAME = 'processed_emails.json';

function processExcelFile() {
    try {
        console.log(chalk.blue('🔍 Buscando archivo Excel...'));
        
        const inputPath = path.join(INPUT_DIR, INPUT_FILE_NAME);
        const workbook = xlsx.readFile(inputPath);
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(firstSheet);
        
        console.log(chalk.green(`✔ Archivo encontrado: ${inputPath}`));
        console.log(chalk.blue('🔄 Procesando datos...'));

        const users = [];
        let skipped = 0;

        data.forEach(row => {
            const current = row[COLUMN_NAMES.CURRENT_EMAIL]?.toString().trim();
            const nuevo = row[COLUMN_NAMES.NEW_EMAIL]?.toString().trim();

            if (!current || !nuevo) {
                skipped++;
                return;
            }

            if (!isValidEmail(current) || !isValidEmail(nuevo)) {
                console.log(chalk.yellow(`⚠️  Fila ignorada: ${current} → ${nuevo}`));
                skipped++;
                return;
            }

            users.push({ current_email: current, new_email: nuevo });
        });

        // Crear directorio output si no existe
        if (!fs.existsSync(OUTPUT_DIR)) {
            fs.mkdirSync(OUTPUT_DIR, { recursive: true });
        }

        const outputPath = path.join(OUTPUT_DIR, OUTPUT_FILE_NAME);
        fs.writeFileSync(outputPath, JSON.stringify({ users }, null, 2));

        console.log(chalk.green('\n✅ Proceso completado!'));
        console.log(chalk.cyan(`📊 Pares procesados: ${users.length}`));
        console.log(chalk.yellow(`⏩ Filas omitidas: ${skipped}`));
        console.log(chalk.green(`💾 Guardado en: ${outputPath}`));

    } catch (error) {
        console.log(chalk.red('❌ Error:'), error.message);
        process.exit(1);
    }
}

function isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

processExcelFile();