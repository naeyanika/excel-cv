const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

class ExcelService {
  static convert(inputFile, outputFormat) {
    const workbook = XLSX.readFile(inputFile.path);
    const outputFileName = `converted-${Date.now()}.${outputFormat}`;
    const outputPath = path.join('uploads', outputFileName);
    
    XLSX.writeFile(workbook, outputPath);
    
    return { outputPath, outputFileName };
  }

  static cleanup(files) {
    files.forEach(file => {
      if (fs.existsSync(file)) {
        fs.unlinkSync(file);
      }
    });
  }
}

module.exports = ExcelService;