const express = require('express');
const path = require('path');
const { Pool } = require('pg');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const STAFF_PIN = process.env.STAFF_PIN || '1234';

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL && process.env.DATABASE_URL.includes('render.com')
    ? { rejectUnauthorized: false }
    : false
});

const DAY_NAMES = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

async function initDb() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS entries (
      id SERIAL PRIMARY KEY,
      entry_date DATE NOT NULL,
      period INT NOT NULL,
      teacher TEXT NOT NULL,
      done TEXT NOT NULL,
      evidence TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);
}

function dayNameFor(dateStr) {
  const d = new Date(dateStr + 'T00:00:00Z');
  return DAY_NAMES[d.getUTCDay()];
}

// ---- Teacher-facing API ----
app.post('/api/submit', async (req, res) => {
  try {
    const { entry_date, period, teacher, done, evidence } = req.body;
    if (!entry_date || !period || !teacher || !done || !evidence) {
      return res.status(400).json({ error: 'جميع الحقول مطلوبة' });
    }
    if (String(evidence).length > 80) {
      return res.status(400).json({ error: 'الشواهد يجب ألا تتجاوز 80 حرفًا' });
    }
    const result = await pool.query(
      `INSERT INTO entries (entry_date, period, teacher, done, evidence)
       VALUES ($1,$2,$3,$4,$5) RETURNING id`,
      [entry_date, period, teacher, done, evidence]
    );
    res.json({ ok: true, id: result.rows[0].id });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'تعذر حفظ البيانات، حاولي مرة أخرى' });
  }
});

// ---- Staff-facing API (protected by a simple PIN header) ----
function checkPin(req, res, next) {
  const pin = req.headers['x-staff-pin'] || req.query.pin;
  if (pin !== STAFF_PIN) return res.status(401).json({ error: 'رمز الدخول غير صحيح' });
  next();
}

app.get('/api/entries', checkPin, async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT id, entry_date, period, teacher, done, evidence, created_at
       FROM entries ORDER BY entry_date DESC, period ASC, id DESC LIMIT 500`
    );
    const rows = result.rows.map(r => ({
      ...r,
      entry_date: r.entry_date.toISOString().slice(0, 10),
      day: dayNameFor(r.entry_date.toISOString().slice(0, 10))
    }));
    res.json(rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'تعذر جلب البيانات' });
  }
});

app.get('/api/summary', checkPin, async (req, res) => {
  try {
    const totalToday = await pool.query(
      `SELECT COUNT(*)::int AS c FROM entries WHERE entry_date = CURRENT_DATE`
    );
    const byTeacherMonth = await pool.query(`
      SELECT teacher, COUNT(*)::int AS c
      FROM entries
      WHERE date_trunc('month', entry_date) = date_trunc('month', CURRENT_DATE)
      GROUP BY teacher ORDER BY c DESC
    `);
    res.json({
      today_total: totalToday.rows[0].c,
      by_teacher_month: byTeacherMonth.rows
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'تعذر جلب الملخص' });
  }
});

app.get('/api/export', checkPin, async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT entry_date, period, teacher, done, evidence, created_at
       FROM entries ORDER BY entry_date DESC, period ASC`
    );
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('حصص الاحتياط', { views: [{ rightToLeft: true }] });
    ws.columns = [
      { header: 'التاريخ', key: 'entry_date', width: 14 },
      { header: 'اليوم', key: 'day', width: 12 },
      { header: 'الحصة', key: 'period', width: 8 },
      { header: 'اسم المعلمة', key: 'teacher', width: 22 },
      { header: 'ما تم تنفيذه', key: 'done', width: 40 },
      { header: 'الشواهد', key: 'evidence', width: 32 },
      { header: 'وقت التسجيل', key: 'created_at', width: 20 }
    ];
    ws.getRow(1).font = { bold: true };
    result.rows.forEach(r => {
      const dateStr = r.entry_date.toISOString().slice(0, 10);
      ws.addRow({
        entry_date: dateStr,
        day: dayNameFor(dateStr),
        period: r.period,
        teacher: r.teacher,
        done: r.done,
        evidence: r.evidence,
        created_at: r.created_at.toISOString().slice(0, 16).replace('T', ' ')
      });
    });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=توثيق_حصص_الاحتياط.xlsx');
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'تعذر تصدير الملف' });
  }
});

app.get('/health', (req, res) => res.send('ok'));

const PORT = process.env.PORT || 3000;
initDb()
  .then(() => app.listen(PORT, () => console.log('Server running on ' + PORT)))
  .catch(err => {
    console.error('DB init failed', err);
    process.exit(1);
  });
