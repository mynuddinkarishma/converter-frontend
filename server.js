const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors()); // Allows your frontend to connect to this backend

// Set up temporary storage for uploaded files
const upload = multer({ dest: 'uploads/' });

app.post('/upload', upload.single('file'), (req, res) => {
    const file = req.file;
    const conversionType = req.body.conversionType;

    if (!file) {
        return res.status(400).send('No file uploaded.');
    }

    console.log(`Received ${file.originalname} for ${conversionType} conversion.`);

    // ==========================================
    // TODO: Add your actual conversion logic here
    // Example libraries you would use:
    // - PDF to Word: 'pdf-parse' or cloud APIs
    // - Word to PDF: 'libreoffice-convert'
    // ==========================================

    // SIMULATED RESPONSE (For testing):
    // Just sending the exact same file back as a placeholder
    res.download(file.path, `converted_${file.originalname}`, (err) => {
        // Clean up the temporary file after sending
        fs.unlinkSync(file.path);
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Backend server running on port ${PORT}`);
});