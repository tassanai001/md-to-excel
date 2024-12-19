const { convertMarkdownToExcel } = require('./src/converter');
const { ExcelToMarkdownConverter } = require('./src/reverseConverter');

const args = process.argv.slice(2);
if (args.length !== 3) {
    console.log('Usage: npm start <conversion-type> <input-file> <output-file>');
    console.log('Example: npm start md-to-xlsx routes.md routes.xlsx');
    console.log('Example: npm start xlsx-to-md routes.xlsx routes.md');
    process.exit(1);
}

const [conversionType, inputFile, outputFile] = args;

if (conversionType === 'md-to-xlsx') {
    convertMarkdownToExcel(inputFile, outputFile);
} else if (conversionType === 'xlsx-to-md') {
    ExcelToMarkdownConverter.convertExcelToMarkdown(inputFile, outputFile);
} else {
    console.log('Invalid conversion type. Use "md-to-xlsx" or "xlsx-to-md".');
    process.exit(1);
}
