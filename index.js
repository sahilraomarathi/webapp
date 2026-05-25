const express = require('express');
const { execFile } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

function runPythonScript(script, data, outputFormat) {
  return new Promise((resolve, reject) => {
    const ext = outputFormat === 'pdf' ? '.pdf' : '.docx';
    const tmpOut = path.join(os.tmpdir(), `doc_${Date.now()}_${Math.random().toString(36).slice(2)}${ext}`);
    const jsonArg = JSON.stringify(data);
    execFile('python3', [script, jsonArg, tmpOut, outputFormat || 'docx'], { timeout: 60000 }, (err, stdout, stderr) => {
      if (err) return reject(new Error(stderr || err.message));
      if (!fs.existsSync(tmpOut)) return reject(new Error('Output file not created'));
      resolve(tmpOut);
    });
  });
}

function handleGenerate(scriptName, requiredFields, filenamePrefix) {
  return async (req, res) => {
    try {
      const format = req.query.format === 'pdf' ? 'pdf' : 'docx';
      for (const k of requiredFields) {
        if (!req.body[k]) return res.status(400).json({ error: `Missing field: ${k}` });
      }
      const script = path.join(__dirname, scriptName);
      const outPath = await runPythonScript(script, req.body, format);
      const safeName = req.body.applicantFullName.replace(/\s+/g, '_');
      const ext = format === 'pdf' ? 'pdf' : 'docx';
      const filename = `${filenamePrefix}_${safeName}.${ext}`;
      const contentType = format === 'pdf'
        ? 'application/pdf'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      res.setHeader('Content-Type', contentType);
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      const stream = fs.createReadStream(outPath);
      stream.pipe(res);
      stream.on('close', () => fs.unlink(outPath, () => {}));
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: e.message });
    }
  };
}

app.post('/generate/no-ration', handleGenerate(
  'fill_no_ration.py',
  ['applicantTitle','applicantFullName','fatherHusbandName','age','dob','address','todayDate'],
  'No_RationCard'
));

app.post('/generate/affidavit', handleGenerate(
  'fill_affidavit.py',
  ['magistrateDate','applicantTitle','applicantFullName','applicantAge','occupation',
   'applicantAddress','deceasedRelation','deceasedTitle','deceasedFullName',
   'svNumber','svDate','expiredDate','expiredPlace','todayDate'],
  'Affidavit'
));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
