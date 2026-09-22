const express = require('express');
const multer = require('multer');
const cors = require('cors');
const JSZip = require('jszip');
const PDFDocument = require('pdfkit');
const mammoth = require('mammoth');
const { createWorker } = require('tesseract.js');
const { PDFParse } = require('pdf-parse');
const { Document, Packer, Paragraph, TextRun, ImageRun } = require('docx');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const PORT = process.env.PORT || 3000;

const extensionMap = {
  'pdf-to-word': '.docx',
  'image-to-word': '.docx',
  'image-to-pdf': '.pdf',
  'word-to-excel': '.xlsx',
  'ppt-to-pdf': '.pdf',
  'word-to-pdf': '.pdf',
};

const mimeTypeMap = {
  'pdf-to-word': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'image-to-word': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'image-to-pdf': 'application/pdf',
  'word-to-excel': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'ppt-to-pdf': 'application/pdf',
  'word-to-pdf': 'application/pdf',
};

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

function escapeXml(value = '') {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildTargetFilename(originalName, conversionType) {
  const baseName = originalName.includes('.')
    ? originalName.slice(0, originalName.lastIndexOf('.'))
    : originalName;
  return `${baseName}_converted${extensionMap[conversionType] || '.bin'}`;
}

async function buildDocxBuffer(fileBuffer) {
  const parser = new PDFParse({ data: fileBuffer });
  const { pdf } = await import('pdf-to-img');

  try {
    const parsedPdf = await parser.getText();
    const pageDocument = await pdf(`data:application/pdf;base64,${fileBuffer.toString('base64')}`, { scale: 1.5 });
    const children = [];
    let pageNumber = 0;

    for await (const image of pageDocument) {
      pageNumber += 1;
      children.push(new Paragraph({
        children: [new ImageRun({
          type: 'png',
          data: image,
          transformation: { width: 612, height: 792 },
        })],
        pageBreakBefore: pageNumber > 1,
      }));
    }

    if (parsedPdf.text.trim()) {
      children.push(new Paragraph({
        children: [new TextRun({ text: 'Extracted text', bold: true })],
        pageBreakBefore: true,
      }));
      children.push(...parsedPdf.text.split(/\r?\n/).map((line) => new Paragraph({
        children: [new TextRun(line || ' ')],
      })));
    }

    const document = new Document({ sections: [{ children }] });
    return Packer.toBuffer(document);
  } finally {
    await parser.destroy();
  }
}

async function extractImageText(fileBuffer) {
  const worker = await createWorker('eng');
  try {
    const result = await worker.recognize(fileBuffer);
    return result.data.text.trim();
  } finally {
    await worker.terminate();
  }
}

async function buildImageDocxBuffer(fileBuffer) {
  const text = await extractImageText(fileBuffer);
  if (!text) {
    throw new Error('No readable text was found in the image. Try a clearer scan.');
  }

  const paragraphs = text.split(/\r?\n/).map((line) => new Paragraph({
    children: [new TextRun(line || ' ')],
  }));
  return Packer.toBuffer(new Document({ sections: [{ children: paragraphs }] }));
}

function buildImagePdfBuffer(fileBuffer) {
  return new Promise((resolve, reject) => {
    const document = new PDFDocument({ autoFirstPage: false });
    const chunks = [];
    document.on('data', (chunk) => chunks.push(chunk));
    document.on('end', () => resolve(Buffer.concat(chunks)));
    document.on('error', reject);
    document.addPage({ size: 'A4', margin: 24 });
    document.image(fileBuffer, { fit: [547, 794], align: 'center', valign: 'center' });
    document.end();
  });
}

async function buildXlsxBuffer(fileName, conversionType) {
  const zip = new JSZip();
  const text = `Converted from ${fileName} using ${conversionType}.`;

  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Converted" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${escapeXml(text)}</t></is></c></row></sheetData></worksheet>`);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function extractWordText(fileBuffer) {
  const result = await mammoth.extractRawText({ buffer: fileBuffer });
  return result.value.trim() || `Converted from ${fileBuffer.length} bytes.`;
}

function buildPdfBuffer(text) {
  const safeText = String(text).replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
  const content = `BT /F1 18 Tf 50 760 Td (${safeText}) Tj ET`;
  const pdf = `%PDF-1.4\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n4 0 obj\n<< /Length ${content.length} >>\nstream\n${content}\nendstream\nendobj\n5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\ntrailer\n<< /Root 1 0 R /Size 6 >>\nstartxref\n0\n%%EOF`;
  return Buffer.from(pdf, 'utf8');
}

async function convertUploadedFile(req, res) {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

  const conversionType = req.body.conversionType || 'pdf-to-word';
  const originalName = req.file.originalname || 'uploaded-file';
  let outputBuffer;

  if (conversionType === 'pdf-to-word') {
    outputBuffer = await buildDocxBuffer(req.file.buffer);
  } else if (conversionType === 'image-to-word') {
    outputBuffer = await buildImageDocxBuffer(req.file.buffer);
  } else if (conversionType === 'image-to-pdf') {
    outputBuffer = await buildImagePdfBuffer(req.file.buffer);
  } else if (conversionType === 'word-to-excel') {
    const text = await extractWordText(req.file.buffer);
    outputBuffer = await buildXlsxBuffer(text, conversionType);
  } else if (conversionType === 'ppt-to-pdf' || conversionType === 'word-to-pdf') {
    outputBuffer = buildPdfBuffer(`Converted from ${originalName} using ${conversionType}.`);
  } else {
    return res.status(400).json({ error: 'Unsupported conversion type.' });
  }

  res.setHeader('Content-Type', mimeTypeMap[conversionType]);
  res.setHeader('Content-Disposition', `attachment; filename="${buildTargetFilename(originalName, conversionType)}"`);
  res.send(outputBuffer);
}

app.get('/', (req, res) => res.json({ status: 'ok', message: 'Converter backend is running.', health: '/health' }));
app.get('/health', (req, res) => res.json({ status: 'ok', message: 'Converter backend is running.' }));

for (const route of ['/upload', '/api/convert']) {
  app.post(route, upload.single('file'), async (req, res) => {
    try {
      await convertUploadedFile(req, res);
    } catch (error) {
      console.error('Conversion failed:', error);
      if (!res.headersSent) res.status(422).json({ error: error.message || 'Unable to convert the uploaded file.' });
    }
  });
}

app.use((req, res) => res.status(404).json({ error: 'Route not found.' }));
app.listen(PORT, () => console.log(`Backend server running on port ${PORT}`));
