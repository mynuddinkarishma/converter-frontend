const express = require('express');
const multer = require('multer');
const cors = require('cors');
const JSZip = require('jszip');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const PORT = process.env.PORT || 3000;

const extensionMap = {
  'pdf-to-word': '.docx',
  'word-to-excel': '.xlsx',
  'ppt-to-pdf': '.pdf',
  'word-to-pdf': '.pdf',
};

const mimeTypeMap = {
  'pdf-to-word': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'word-to-excel': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'ppt-to-pdf': 'application/pdf',
  'word-to-pdf': 'application/pdf',
};

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

function buildTargetFilename(originalName, conversionType) {
  const baseName = originalName.includes('.')
    ? originalName.slice(0, originalName.lastIndexOf('.'))
    : originalName;

  const extension = extensionMap[conversionType] || '.bin';
  return `${baseName}_converted${extension}`;
}

function escapeXml(value = '') {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

async function buildDocxBuffer(fileName, conversionType) {
  const zip = new JSZip();
  const text = `Converted from ${fileName} using ${conversionType}.`;

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(fileName)}</dc:title><dc:creator>Converter Backend</dc:creator><cp:lastModifiedBy>Converter Backend</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-09-21T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-21T00:00:00Z</dcterms:modified></cp:coreProperties>`);
  zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Converter Backend</Application></Properties>`);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`);

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function buildXlsxBuffer(fileName, conversionType) {
  const zip = new JSZip();
  const text = `Converted from ${fileName} using ${conversionType}.`;

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(fileName)}</dc:title><dc:creator>Converter Backend</dc:creator><cp:lastModifiedBy>Converter Backend</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-09-21T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-21T00:00:00Z</dcterms:modified></cp:coreProperties>`);
  zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Converter Backend</Application></Properties>`);
  zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Converted" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${escapeXml(text)}</t></is></c></row></sheetData></worksheet>`);

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function buildPdfBuffer(text) {
  const safeText = String(text)
    .replace(/\\/g, '\\\\')
    .replace(/\(/g, '\\(')
    .replace(/\)/g, '\\)');

  const content = `BT /F1 18 Tf 50 760 Td (${safeText}) Tj ET`;
  const pdf = `%PDF-1.4\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n4 0 obj\n<< /Length ${content.length} >>\nstream\n${content}\nendstream\nendobj\n5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\nxref\n0 6\n0000000000 65535 f \n0000000010 00000 n \n0000000062 00000 n \n0000000122 00000 n \n0000000276 00000 n \n0000000617 00000 n \ntrailer\n<< /Root 1 0 R /Size 6 >>\nstartxref\n690\n%%EOF`;

  return Buffer.from(pdf, 'utf8');
}

async function convertUploadedFile(req, res) {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded.' });
  }

  const conversionType = req.body.conversionType || 'pdf-to-word';
  const originalName = req.file.originalname || 'uploaded-file';
  const filename = buildTargetFilename(originalName, conversionType);
  const mimeType = mimeTypeMap[conversionType] || 'application/octet-stream';

  let outputBuffer;

  if (conversionType === 'pdf-to-word') {
    outputBuffer = await buildDocxBuffer(originalName, conversionType);
  } else if (conversionType === 'word-to-excel') {
    outputBuffer = await buildXlsxBuffer(originalName, conversionType);
  } else if (conversionType === 'ppt-to-pdf' || conversionType === 'word-to-pdf') {
    outputBuffer = buildPdfBuffer(`Converted from ${originalName} using ${conversionType}.`);
  } else {
    return res.status(400).json({ error: 'Unsupported conversion type.' });
  }

  res.setHeader('Content-Type', mimeType);
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(outputBuffer);
}

app.get('/health', (req, res) => {
  res.json({ status: 'ok', message: 'Converter backend is running.' });
});

app.get('/', (req, res) => {
  res.json({ status: 'ok', message: 'Converter backend is running.', health: '/health' });
});

app.post('/upload', upload.single('file'), async (req, res) => {
  await convertUploadedFile(req, res);
});

app.post('/api/convert', upload.single('file'), async (req, res) => {
  await convertUploadedFile(req, res);
});

app.use((req, res) => {
  res.status(404).json({ error: 'Route not found.' });
});

app.listen(PORT, () => {
  console.log(`Backend server running on port ${PORT}`);
});
