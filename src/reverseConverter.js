const fs = require('fs');
const XLSX = require('xlsx');

class ExcelToMarkdownConverter {
    static convertExcelToMarkdown(inputFile, outputFile) {
        // Read the Excel file
        const workbook = XLSX.readFile(inputFile);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        // Convert JSON data to Markdown format
        let markdownContent = '';
        let currentSection = '';

        jsonData.forEach(row => {
            if (row.Section && row.Section !== currentSection) {
                currentSection = row.Section;
                markdownContent += `# ${currentSection}\n`;
            }
            if (row.Path) {
                markdownContent += `path: "${row.Path}"\n`;
                if (row.RouteAuthorize) {
                    markdownContent += `routeAuthorize: ["${row.RouteAuthorize.split(',').join('", "')}"]\n`;
                } else {
                    markdownContent += `routeAuthorize: ""\n`;
                }
            }
        });

        // Write the Markdown content to the output file
        fs.writeFileSync(outputFile, markdownContent.trim());
    }
}

module.exports = { ExcelToMarkdownConverter };
