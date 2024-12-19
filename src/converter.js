const fs = require('fs');
const XLSX = require('xlsx');

class MarkdownConverter {
    static parseMarkdown(content) {
        const lines = content.split('\n').filter(line => line.trim() !== '');
        const data = [];
        let currentSection = '';
        let previousSection = '';

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();

            if (line.startsWith('#')) {
                previousSection = currentSection;
                currentSection = line.replace('#', '').trim();

                // Add empty row between sections (except for the first section)
                if (previousSection !== '' && data.length > 0) {
                    data.push({
                        Section: '',
                        Path: '',
                        RouteAuthorize: ''
                    });
                }
                continue;
            }

            if (line.startsWith('path:')) {
                const path = line.replace('path:', '').trim().replace(/"/g, '');
                let routeAuthorize = '';

                if (i + 1 < lines.length && lines[i + 1].trim().startsWith('routeAuthorize:')) {
                    routeAuthorize = lines[i + 1]
                        .replace('routeAuthorize:', '')
                        .trim()
                        .replace(/[\[\]"]/g, '');
                    i++;
                }

                data.push({
                    Section: currentSection,
                    Path: path,
                    RouteAuthorize: routeAuthorize
                });
            }
        }

        return data;
    }

    static createExcel(data, outputFile) {
        // Create worksheet
        const ws = XLSX.utils.json_to_sheet(data);

        // Set column widths
        const colWidths = {
            'A': 30,  // Section
            'B': 60,  // Path
            'C': 40   // RouteAuthorize
        };
        ws['!cols'] = Object.values(colWidths).map(width => ({ width }));

        // Apply styles to header row
        const headerStyle = {
            font: { bold: true },
            fill: { fgColor: { rgb: "CCCCCC" } }
        };

        // Create workbook and add worksheet
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Routes');

        // Write file
        XLSX.writeFile(wb, outputFile);
    }

    static formatData(data) {
        // Group data by section
        const groupedData = [];
        let currentSection = '';

        data.forEach(row => {
            if (row.Section !== currentSection && row.Section !== '') {
                // Add a blank row before new section (except for first section)
                if (currentSection !== '') {
                    groupedData.push({
                        Section: '',
                        Path: '',
                        RouteAuthorize: ''
                    });
                }
                currentSection = row.Section;

                // Add section header
                groupedData.push({
                    Section: row.Section,
                    Path: '',
                    RouteAuthorize: ''
                });
            }

            // Add the regular row
            if (row.Path !== '') {
                groupedData.push({
                    Section: '',  // Clear section name from repeated rows
                    Path: row.Path,
                    RouteAuthorize: row.RouteAuthorize
                });
            }
        });

        return groupedData;
    }
}

function convertMarkdownToExcel(inputFile, outputFile) {
    try {
        // Validate input file exists
        if (!fs.existsSync(inputFile)) {
            throw new Error(`Input file '${inputFile}' does not exist`);
        }

        // Validate input file is markdown
        if (!inputFile.toLowerCase().endsWith('.md')) {
            throw new Error('Input file must be a markdown file (.md)');
        }

        // Validate output file is xlsx
        if (!outputFile.toLowerCase().endsWith('.xlsx')) {
            throw new Error('Output file must be an Excel file (.xlsx)');
        }

        // Read and parse markdown
        console.log(`Reading markdown file: ${inputFile}`);
        const markdownContent = fs.readFileSync(inputFile, 'utf8');
        const rawData = MarkdownConverter.parseMarkdown(markdownContent);

        // Format data with proper sections and spacing
        console.log('Formatting data with sections...');
        const formattedData = MarkdownConverter.formatData(rawData);

        // Create Excel file
        console.log(`Creating Excel file: ${outputFile}`);
        MarkdownConverter.createExcel(formattedData, outputFile);

        // Log statistics
        const totalRoutes = rawData.filter(row => row.Path !== '').length;
        const totalSections = new Set(rawData.filter(row => row.Section !== '').map(row => row.Section)).size;

        console.log('\nConversion Summary:');
        console.log(`✓ Total routes processed: ${totalRoutes}`);
        console.log(`✓ Total sections: ${totalSections}`);
        console.log(`✓ Output file created: ${outputFile}`);

    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

// If running directly (not imported as module)
if (require.main === module) {
    const args = process.argv.slice(2);
    if (args.length !== 2) {
        console.log('Usage: node converter.js <input-markdown-file> <output-excel-file>');
        console.log('Example: node converter.js routes.md routes.xlsx');
        process.exit(1);
    }

    const [inputFile, outputFile] = args;
    convertMarkdownToExcel(inputFile, outputFile);
}

module.exports = { convertMarkdownToExcel };