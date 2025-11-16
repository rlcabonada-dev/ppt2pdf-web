/**
 * server.js
 * Minimal Express server that:
 * - serves static frontend from /public
 * - accepts file uploads at POST /convert
 * - calls LibreOffice (soffice) headless to convert to PDF
 * - returns the generated PDF as a download
 *
 * Note: LibreOffice (soffice) must be installed and in PATH.
 */

const express = require('express');
const multer = require('multer');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { spawn } = require('child_process');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({
  dest: path.join(os.tmpdir(), 'ppt_uploads'),
  limits: {
    // no file size limit in code (server or nginx may still limit)
  },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.ppt' || ext === '.pptx') cb(null, true);
    else cb(new Error('Only .ppt and .pptx files allowed'), false);
  }
});

// Serve frontend
app.use(express.static(path.join(__dirname, 'public')));

// Temporary store for generated files (id -> { path, downloadName }) with TTL
const generatedFiles = {};
// Temporary store for preview images (id -> path) with TTL
const previewFiles = {};

// Detect soffice path. You can override by setting environment variable `SOFFICE_PATH`.
const SOFFICE_CANDIDATES = [
  process.env.SOFFICE_PATH,
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  '/usr/bin/soffice',
  '/usr/bin/libreoffice',
  '/Applications/LibreOffice.app/Contents/MacOS/soffice',
  'soffice'
].filter(Boolean);

function findSoffice() {
  for (const p of SOFFICE_CANDIDATES) {
    try {
      if (p === 'soffice') return 'soffice'; // assume in PATH
      if (fs.existsSync(p)) return p;
    } catch (e) {
      // ignore
    }
  }
  return 'soffice';
}

const SOFFICE_BIN = findSoffice();
console.log('Using soffice binary:', SOFFICE_BIN);

// helper: run soffice to convert a single input file to PDF in workdir
function convertWithSoffice(inPath, workdir, timeoutMs = 150000) {
  return new Promise((resolve, reject) => {
    const args = ['--headless', '--invisible', '--convert-to', 'pdf', '--outdir', workdir, inPath];
    const soffice = spawn(SOFFICE_BIN, args, { stdio: 'ignore' });

    const timeout = setTimeout(() => {
      try { soffice.kill(); } catch (e) {}
      reject(new Error('Conversion timeout'));
    }, timeoutMs);

    soffice.on('error', (err) => {
      clearTimeout(timeout);
      reject(err);
    });

    soffice.on('close', (code) => {
      clearTimeout(timeout);
      // accept close and let caller check for produced PDFs
      resolve(code);
    });
  });
}

// Convert endpoint: accepts multiple files (field name 'file')
app.post('/convert', upload.array('file', 50), async (req, res) => {
  const files = req.files;
  if (!files || files.length === 0) return res.status(400).json({ error: 'No files uploaded.' });

  const workdir = path.join(os.tmpdir(), 'ppt2pdf_' + uuidv4());
  try {
    fs.mkdirSync(workdir, { recursive: true });

    // move uploaded files to workdir
    for (const f of files) {
      const inPath = path.join(workdir, f.originalname);
      fs.renameSync(f.path, inPath);
    }

    // convert each file (sequentially to avoid overload)
    for (const f of files) {
      const inPath = path.join(workdir, f.originalname);
      try {
        await convertWithSoffice(inPath, workdir);
      } catch (err) {
        cleanup(workdir);
        return res.status(500).json({ error: 'Conversion failed for: ' + f.originalname + '. ' + (err.message || '') });
      }
    }

    // collect produced PDFs
    const pdfFiles = fs.readdirSync(workdir).filter(fn => fn.toLowerCase().endsWith('.pdf'));
    if (pdfFiles.length === 0) {
      cleanup(workdir);
      return res.status(500).json({ error: 'Conversion failed: no PDF produced.' });
    }

    // prepare storage dir for final artifacts
    const filesDir = path.join(os.tmpdir(), 'ppt2pdf_files');
    fs.mkdirSync(filesDir, { recursive: true });
    const id = uuidv4();

    // If only one PDF -> copy and return direct download URL
    if (pdfFiles.length === 1) {
      const pdfPath = path.join(workdir, pdfFiles[0]);
      const destName = id + path.extname(pdfPath);
      const destPath = path.join(filesDir, destName);
      fs.copyFileSync(pdfPath, destPath);
      generatedFiles[id] = { path: destPath, downloadName: pdfFiles[0] };
      // schedule cleanup
      setTimeout(() => { try { fs.unlinkSync(destPath); } catch (e) {} delete generatedFiles[id]; }, 10 * 60 * 1000);
      cleanup(workdir);
      return res.json({ success: true, downloadUrl: `/download/${id}`, filename: pdfFiles[0] });
    }

    // Multiple PDFs -> create a ZIP
    const archiver = require('archiver');
    const zipName = id + '.zip';
    const zipPath = path.join(filesDir, zipName);
    const output = fs.createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => {
      generatedFiles[id] = { path: zipPath, downloadName: 'converted-pdfs.zip' };
      setTimeout(() => { try { fs.unlinkSync(zipPath); } catch (e) {} delete generatedFiles[id]; }, 10 * 60 * 1000);
      cleanup(workdir);
      return res.json({ success: true, downloadUrl: `/download/${id}`, filename: 'converted-pdfs.zip' });
    });

    archive.on('error', (err) => {
      cleanup(workdir);
      return res.status(500).json({ error: 'Failed creating ZIP: ' + err.message });
    });

    archive.pipe(output);
    for (const pdf of pdfFiles) {
      const p = path.join(workdir, pdf);
      archive.file(p, { name: pdf });
    }
    archive.finalize();

  } catch (err) {
    // cleanup uploaded temp files if any
    try { if (files) files.forEach(f => { try { fs.unlinkSync(f.path); } catch (e) {} }); } catch (e) {}
    cleanup(workdir);
    return res.status(500).json({ error: 'Server error.' });
  }
});

// Preview endpoint: accepts a single PPT/PPTX and returns a URL to the first-slide PNG
app.post('/preview', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });
  const uploadedPath = req.file.path;
  const originalName = req.file.originalname;
  const workdir = path.join(os.tmpdir(), 'ppt_preview_' + uuidv4());
  try {
    fs.mkdirSync(workdir, { recursive: true });
    const inPath = path.join(workdir, originalName);
    fs.renameSync(uploadedPath, inPath);

    // convert to PNG (LibreOffice will export slides as PNG files)
    const args = ['--headless', '--invisible', '--convert-to', 'png', '--outdir', workdir, inPath];
    const soffice = spawn(SOFFICE_BIN, args);

    const timeout = setTimeout(() => { try { soffice.kill(); } catch (e) {} }, 120000);

    let stderr = '';
    soffice.stderr && soffice.stderr.on('data', d => { stderr += d.toString(); });

    soffice.on('error', (err) => {
      clearTimeout(timeout);
      cleanup(workdir);
      return res.status(500).json({ error: 'Failed to start LibreOffice (soffice).', detail: err.message });
    });

    soffice.on('close', (code) => {
      clearTimeout(timeout);
      try {
        const images = fs.readdirSync(workdir).filter(f => /\.png$/i.test(f));
        if (images.length === 0) {
          cleanup(workdir);
          return res.status(500).json({ error: 'Preview generation failed: no image produced.', detail: stderr });
        }
        // pick the first image (LibreOffice typically names slides as <basename>.png or <basename>-0001.png)
        images.sort();
        const imgPath = path.join(workdir, images[0]);

        // copy to filesDir for serving
        const filesDir = path.join(os.tmpdir(), 'ppt2pdf_preview_files');
        fs.mkdirSync(filesDir, { recursive: true });
        const id = uuidv4();
        const destName = id + path.extname(imgPath);
        const destPath = path.join(filesDir, destName);
        fs.copyFileSync(imgPath, destPath);

        previewFiles[id] = destPath;
        // schedule cleanup
        setTimeout(() => { try { fs.unlinkSync(destPath); } catch (e) {} delete previewFiles[id]; }, 10 * 60 * 1000);

        cleanup(workdir);
        return res.json({ success: true, previewUrl: `/preview/${id}` });
      } catch (err) {
        cleanup(workdir);
        return res.status(500).json({ error: 'Server error during preview generation.' });
      }
    });
  } catch (err) {
    try { fs.unlinkSync(uploadedPath); } catch (e) {}
    cleanup(workdir);
    return res.status(500).json({ error: 'Server error.' });
  }
});

// Serve preview image
app.get('/preview/:id', (req, res) => {
  const id = req.params.id;
  const p = previewFiles[id];
  if (!p || !fs.existsSync(p)) return res.status(404).send('Preview not found or expired.');
  res.sendFile(p);
});

app.get('/health', (req, res) => res.json({ status: 'ok' }));
app.get('/download/:id', (req, res) => {
  const id = req.params.id;
  const info = generatedFiles[id];
  if (!info || !fs.existsSync(info.path)) return res.status(404).send('File not found or expired.');
  const downloadName = info.downloadName || path.basename(info.path);
  res.download(info.path, downloadName, (err) => {
    try { fs.unlinkSync(info.path); } catch (e) {}
    delete generatedFiles[id];
  });
});

app.listen(PORT, () => console.log(`Server listening on http://localhost:${PORT}`));

function cleanup(dir) {
  try {
    if (fs.existsSync(dir)) {
      fs.readdirSync(dir).forEach(file => {
        try { fs.unlinkSync(path.join(dir, file)); } catch (e) {}
      });
      fs.rmdirSync(dir);
    }
  } catch (e) {
    // ignore cleanup errors
  }
}
