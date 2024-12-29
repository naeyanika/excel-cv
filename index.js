const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Create uploads directory if it doesn't exist
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

app.use(express.static('public'));

app.post('/convert', upload.single('file'), (req, res) => {
  try {
    const inputFile = req.file;
    const outputFormat = req.body.format;
    
    if (!inputFile || !outputFormat) {
      return res.status(400).json({ error: 'File and output format are required' });
    }

    // Read the workbook
    const workbook = XLSX.readFile(inputFile.path);
    
    // Generate output filename
    const outputFileName = `converted-${Date.now()}.${outputFormat}`;
    const outputPath = path.join('uploads', outputFileName);

    // Write the file in the requested format
    XLSX.writeFile(workbook, outputPath);

    // Send the file
    res.download(outputPath, outputFileName, (err) => {
      // Clean up files after sending
      fs.unlinkSync(inputFile.path);
      fs.unlinkSync(outputPath);
    });
  } catch (error) {
    res.status(500).json({ error: 'Conversion failed' });
  }
});

app.listen(3000, () => {
  console.log('Excel converter running on http://localhost:3000');
});