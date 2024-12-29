const express = require('express');
const router = express.Router();
const upload = require('../config/multer');
const ExcelService = require('../services/excelService');

router.post('/convert', upload.single('file'), (req, res) => {
  try {
    const inputFile = req.file;
    const outputFormat = req.body.format;
    
    if (!inputFile || !outputFormat) {
      return res.status(400).json({ error: 'File and output format are required' });
    }

    const { outputPath, outputFileName } = ExcelService.convert(inputFile, outputFormat);

    res.download(outputPath, outputFileName, (err) => {
      ExcelService.cleanup([inputFile.path, outputPath]);
    });
  } catch (error) {
    res.status(500).json({ error: 'Conversion failed' });
  }
});

module.exports = router;