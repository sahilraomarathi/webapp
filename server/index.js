const express = require('express');
const { execFile } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

function runPythonScript(script, data) {
  return new Promise((resolve, reject) => {
    const tmpDocx = path.join(os.tmpdir(), `doc_${Date.now()}_${Math.random().toString(36).slice(2)}.docx`);
    const jsonArg = JSON.stringify(data);
    execFile('python3', [script, jsonArg, tmpDocx], { timeout: 30000 }, (err, stdout, stderr) => {
      if (err) return reject(new Error(stderr || err.message));
      if (!fs.existsSync(tmpDocx)) return reject(new Error('DOCX not created'));
      resolve(tmpDocx);
    });
  });
}

app.post('/generate/no-ration', async (req, res) => {
  try {
    const required = ['applicantTitle','applicantFullName','fatherHusbandName','age','dob','address','todayDate'];
    for (const k of required) {
      if (!req.body[k]) return res.status(400).json({ error: `Missing field: ${k}` });
    }
    const docxPath = await runPythonScript(path.join(__dirname, 'fill_no_ration.py'), req.body);
    const filename = `No_RationCard_${req.body.applicantFullName.replace(/\s+/g,'_')}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    const stream = fs.createReadStream(docxPath);
    stream.pipe(res);
    stream.on('close', () => fs.unlink(docxPath, () => {}));
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/generate/affidavit', async (req, res) => {
  try {
    const required = ['magistrateDate','applicantTitle','applicantFullName','applicantAge','occupation',
                      'applicantAddress','deceasedRelation','deceasedTitle','deceasedFullName',
                      'svNumber','svDate','expiredDate','expiredPlace','todayDate'];
    for (const k of required) {
      if (!req.body[k]) return res.status(400).json({ error: `Missing field: ${k}` });
    }
    const docxPath = await runPythonScript(path.join(__dirname, 'fill_affidavit.py'), req.body);
    const filename = `Affidavit_${req.body.applicantFullName.replace(/\s+/g,'_')}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    const stream = fs.createReadStream(docxPath);
    stream.pipe(res);
    stream.on('close', () => fs.unlink(docxPath, () => {}));
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
