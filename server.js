const express = require('express');
const multer = require('multer');
const fs = require('fs');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = 3000;


app.get('/favicon.ico', (req, res) => res.status(204).end());

// Set up file storage
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public')); // Serve frontend files
app.use(express.json());

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// API: Upload JSON and Excel format files
app.post('/upload', upload.fields([{ name: 'jsonFile' }, { name: 'excelFile' }]), async (req, res) => {
    try {
        console.log("🚀 Files uploaded, processing started...");

        // Get file paths
        const jsonFilePath = req.files['jsonFile'][0].path;
        const excelFilePath = req.files['excelFile'][0].path;
        const outputFilePath = path.join(__dirname, 'public', 'applicants.xlsx');

        // Read Excel format file
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const worksheet = workbook.worksheets[0];

        // Extract output headers and JSON property names
        let outputKeys = worksheet.getRow(1).values.slice(1).map(val => val?.toString().trim()).filter(Boolean);
        let jsonKeys = worksheet.getRow(2).values.slice(1).map(val => val?.toString().trim().replace(/\[|\]/g, '')).filter(Boolean);

        if (outputKeys.length !== jsonKeys.length) {
            throw new Error("Mismatch between header keys and JSON property names.");
        }

        console.log("✅ Extracted Headers & JSON Mapping...");

        // Read JSON file
        const jsonData = fs.readFileSync(jsonFilePath, 'utf-8');
        const parsedData = JSON.parse(jsonData);

        if (!Array.isArray(parsedData) || parsedData.length === 0) {
            throw new Error("Invalid or empty JSON data.");
        }

        const startTime = Date.now();

        // Create output workbook
        const outputWorkbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: outputFilePath });
        const outputWorksheet = outputWorkbook.addWorksheet('Applicants');

        // Write headers
        outputWorksheet.addRow(outputKeys).commit();

        // Write data rows in batches for performance
        parsedData.forEach((applicant, index) => {
            let row = outputKeys.map((_, i) => applicant[jsonKeys[i]]?.toString() ?? '');
            outputWorksheet.addRow(row).commit();
            if (index % 100 === 0) console.log(`✅ Processed ${index + 1} rows...`);
        });

        await outputWorkbook.commit();
        console.log(`🎉 File "applicants.xlsx" generated successfully!`);
        console.log(`⏱️ Time taken: ${(Date.now() - startTime) / 1000} seconds`);

        // Delete uploaded files (JSON & Excel Format) after processing
        fs.unlinkSync(jsonFilePath);
        fs.unlinkSync(excelFilePath);
        console.log("🗑️ Deleted temporary JSON & Excel format files!");

        res.json({ success: true, downloadUrl: '/download' });

    } catch (err) {
        console.error("❌ Error:", err.message);
        res.status(500).json({ error: err.message });
    }
});

// API: Download the processed file and delete it afterward
app.get('/download', (req, res) => {
    const filePath = path.join(__dirname, 'public', 'applicants.xlsx');

    if (fs.existsSync(filePath)) {
        res.download(filePath, (err) => {
            if (!err) {
                // Delete applicants.xlsx after download
                fs.unlinkSync(filePath);
                console.log("🗑️ Deleted applicants.xlsx after download!");
            } else {
                console.error("❌ Error downloading file:", err);
            }
        });
    } else {
        res.status(404).json({ error: "❌ File not found. Please upload and process data first." });
    }
});

app.listen(port, () => {
    console.log(`🚀 Server running at http://localhost:${port}`);
});
