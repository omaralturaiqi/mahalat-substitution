const express = require('express');
const path = require('path');
const { Pool } = require('pg');
const ExcelJS = require('exceljs');
const multer = require('multer');

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 3 * 1024 * 1024 }, // 3MB max per image
  fileFilter: (req, file, cb) => {
    if (!file.mimetype.startsWith('image/')) return cb(new Error('يجب أن يكون الملف صورة'));
    cb(null, true);
  }
});

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
  // Added later: optional evidence photo, stored directly in the database.
  await pool.query(`ALTER TABLE entries ADD COLUMN IF NOT EXISTS evidence_image BYTEA;`);
  await pool.query(`ALTER TABLE entries ADD COLUMN IF NOT EXISTS evidence_image_mime TEXT;`);

  // ---- School timetable builder (separate feature; does not touch entries above) ----
  await pool.query(`
    CREATE TABLE IF NOT EXISTS tt_subjects (
      id SERIAL PRIMARY KEY,
      name TEXT UNIQUE NOT NULL
    );
  `);
  await pool.query(`
    CREATE TABLE IF NOT EXISTS tt_classes (
      id SERIAL PRIMARY KEY,
      grade TEXT NOT NULL,
      section TEXT NOT NULL,
      UNIQUE(grade, section)
    );
  `);
  await pool.query(`
    CREATE TABLE IF NOT EXISTS tt_requirements (
      id SERIAL PRIMARY KEY,
      class_id INT NOT NULL REFERENCES tt_classes(id) ON DELETE CASCADE,
      subject_id INT NOT NULL REFERENCES tt_subjects(id) ON DELETE CASCADE,
      teacher TEXT NOT NULL,
      weekly_periods INT NOT NULL CHECK (weekly_periods > 0),
      UNIQUE(class_id, subject_id)
    );
  `);
  await pool.query(`
    CREATE TABLE IF NOT EXISTS tt_slots (
      id SERIAL PRIMARY KEY,
      class_id INT NOT NULL REFERENCES tt_classes(id) ON DELETE CASCADE,
      day INT NOT NULL,
      period INT NOT NULL,
      subject_id INT REFERENCES tt_subjects(id) ON DELETE SET NULL,
      teacher TEXT,
      UNIQUE(class_id, day, period)
    );
  `);
}

function dayNameFor(dateStr) {
  const d = new Date(dateStr + 'T00:00:00Z');
  return DAY_NAMES[d.getUTCDay()];
}

// ---- Teacher-facing API ----
// Accepts multipart/form-data so an optional evidence photo can ride along
// with the same fields. The image field name is "image".
app.post('/api/submit', (req, res) => {
  upload.single('image')(req, res, async (uploadErr) => {
    if (uploadErr) {
      return res.status(400).json({ error: uploadErr.message === 'يجب أن يكون الملف صورة'
        ? uploadErr.message
        : 'تعذر رفع الصورة (الحد الأقصى 3 ميجابايت)' });
    }
    try {
      const { entry_date, period, teacher, done, evidence } = req.body;
      if (!entry_date || !period || !teacher || !done || !evidence) {
        return res.status(400).json({ error: 'جميع الحقول مطلوبة' });
      }
      if (String(evidence).length > 80) {
        return res.status(400).json({ error: 'الشواهد يجب ألا تتجاوز 80 حرفًا' });
      }
      const image = req.file ? req.file.buffer : null;
      const imageMime = req.file ? req.file.mimetype : null;
      const result = await pool.query(
        `INSERT INTO entries (entry_date, period, teacher, done, evidence, evidence_image, evidence_image_mime)
         VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING id`,
        [entry_date, period, teacher, done, evidence, image, imageMime]
      );
      res.json({ ok: true, id: result.rows[0].id });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'تعذر حفظ البيانات، حاولي مرة أخرى' });
    }
  });
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
      `SELECT id, entry_date, period, teacher, done, evidence, created_at,
              (evidence_image IS NOT NULL) AS has_image
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

// Serves one entry's evidence photo (protected by the same staff PIN).
app.get('/api/image/:id', checkPin, async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT evidence_image, evidence_image_mime FROM entries WHERE id = $1`,
      [req.params.id]
    );
    const row = result.rows[0];
    if (!row || !row.evidence_image) return res.status(404).send('لا توجد صورة');
    res.setHeader('Content-Type', row.evidence_image_mime || 'image/jpeg');
    res.send(row.evidence_image);
  } catch (err) {
    console.error(err);
    res.status(500).send('تعذر تحميل الصورة');
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
    const { from, to } = req.query;
    const conditions = [];
    const params = [];
    if (from) { params.push(from); conditions.push(`entry_date >= $${params.length}`); }
    if (to) { params.push(to); conditions.push(`entry_date <= $${params.length}`); }
    const whereClause = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';

    const result = await pool.query(
      `SELECT id, entry_date, period, teacher, done, evidence, created_at,
              evidence_image, evidence_image_mime
       FROM entries ${whereClause} ORDER BY entry_date DESC, period ASC`,
      params
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
      { header: 'صورة الشاهد', key: 'photo', width: 16 },
      { header: 'وقت التسجيل', key: 'created_at', width: 20 }
    ];
    ws.getRow(1).font = { bold: true };

    result.rows.forEach((r, idx) => {
      const rowNum = idx + 2; // header is row 1
      const dateStr = r.entry_date.toISOString().slice(0, 10);
      ws.addRow({
        entry_date: dateStr,
        day: dayNameFor(dateStr),
        period: r.period,
        teacher: r.teacher,
        done: r.done,
        evidence: r.evidence,
        photo: '',
        created_at: r.created_at.toISOString().slice(0, 16).replace('T', ' ')
      });
      if (r.evidence_image) {
        const ext = (r.evidence_image_mime || '').includes('png') ? 'png' : 'jpeg';
        const imgId = wb.addImage({ buffer: r.evidence_image, extension: ext });
        ws.addImage(imgId, {
          tl: { col: 6, row: rowNum - 1 },
          ext: { width: 90, height: 90 }
        });
        ws.getRow(rowNum).height = 70;
      }
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const utf8Name = encodeURIComponent('توثيق_حصص_الاحتياط.xlsx');
    res.setHeader('Content-Disposition', `attachment; filename="substitution-report.xlsx"; filename*=UTF-8''${utf8Name}`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'تعذر تصدير الملف' });
  }
});

// ============================================================
// School Timetable Builder (separate feature — reads/writes its own
// tt_* tables only; never touches the `entries` substitution-tracker data)
// ============================================================
const TT_DAY_NAMES = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'];
// Sunday has 7 periods; the rest of the week has 6.
const TT_PERIODS_PER_DAY = [7, 6, 6, 6, 6];
const TT_MAX_PERIODS = Math.max(...TT_PERIODS_PER_DAY);

// -- Subjects --
app.get('/api/tt/subjects', checkPin, async (req, res) => {
  const r = await pool.query('SELECT * FROM tt_subjects ORDER BY name');
  res.json(r.rows);
});
app.post('/api/tt/subjects', checkPin, async (req, res) => {
  try {
    const { name } = req.body;
    if (!name) return res.status(400).json({ error: 'اسم المادة مطلوب' });
    const r = await pool.query('INSERT INTO tt_subjects (name) VALUES ($1) ON CONFLICT (name) DO NOTHING RETURNING *', [name]);
    res.json(r.rows[0] || { ok: true, existed: true });
  } catch (err) { console.error(err); res.status(500).json({ error: 'تعذر إضافة المادة' }); }
});
app.delete('/api/tt/subjects/:id', checkPin, async (req, res) => {
  await pool.query('DELETE FROM tt_subjects WHERE id=$1', [req.params.id]);
  res.json({ ok: true });
});

// -- Classes --
app.get('/api/tt/classes', checkPin, async (req, res) => {
  const r = await pool.query('SELECT * FROM tt_classes ORDER BY grade, section');
  res.json(r.rows);
});
app.post('/api/tt/classes', checkPin, async (req, res) => {
  try {
    const { grade, section } = req.body;
    if (!grade || !section) return res.status(400).json({ error: 'الصف والفصل مطلوبان' });
    const r = await pool.query(
      'INSERT INTO tt_classes (grade, section) VALUES ($1,$2) ON CONFLICT (grade, section) DO NOTHING RETURNING *',
      [grade, section]
    );
    res.json(r.rows[0] || { ok: true, existed: true });
  } catch (err) { console.error(err); res.status(500).json({ error: 'تعذر إضافة الفصل' }); }
});
app.delete('/api/tt/classes/:id', checkPin, async (req, res) => {
  await pool.query('DELETE FROM tt_classes WHERE id=$1', [req.params.id]);
  res.json({ ok: true });
});

// -- Requirements (weekly periods + teacher per class/subject) --
app.get('/api/tt/requirements', checkPin, async (req, res) => {
  const r = await pool.query(`
    SELECT r.*, s.name AS subject_name, c.grade, c.section
    FROM tt_requirements r
    JOIN tt_subjects s ON s.id = r.subject_id
    JOIN tt_classes c ON c.id = r.class_id
    ORDER BY c.grade, c.section, s.name
  `);
  res.json(r.rows);
});
app.post('/api/tt/requirements', checkPin, async (req, res) => {
  try {
    const { class_id, subject_id, teacher, weekly_periods } = req.body;
    if (!class_id || !subject_id || !teacher || !weekly_periods) {
      return res.status(400).json({ error: 'جميع الحقول مطلوبة' });
    }
    const r = await pool.query(
      `INSERT INTO tt_requirements (class_id, subject_id, teacher, weekly_periods) VALUES ($1,$2,$3,$4)
       ON CONFLICT (class_id, subject_id) DO UPDATE SET teacher=$3, weekly_periods=$4 RETURNING *`,
      [class_id, subject_id, teacher, weekly_periods]
    );
    res.json(r.rows[0]);
  } catch (err) { console.error(err); res.status(500).json({ error: 'تعذر حفظ المتطلب' }); }
});
app.delete('/api/tt/requirements/:id', checkPin, async (req, res) => {
  await pool.query('DELETE FROM tt_requirements WHERE id=$1', [req.params.id]);
  res.json({ ok: true });
});

// -- Slots (the generated/edited timetable grid) --
app.get('/api/tt/slots', checkPin, async (req, res) => {
  const { class_id, teacher } = req.query;
  let query = `
    SELECT sl.*, s.name AS subject_name, c.grade, c.section
    FROM tt_slots sl
    LEFT JOIN tt_subjects s ON s.id = sl.subject_id
    JOIN tt_classes c ON c.id = sl.class_id
  `;
  const params = [];
  if (class_id) { params.push(class_id); query += ` WHERE sl.class_id = $${params.length}`; }
  else if (teacher) { params.push(teacher); query += ` WHERE sl.teacher = $${params.length}`; }
  query += ' ORDER BY sl.day, sl.period';
  const r = await pool.query(query, params);
  res.json(r.rows);
});

app.put('/api/tt/slots', checkPin, async (req, res) => {
  try {
    const { class_id, day, period, subject_id, teacher } = req.body;
    if (class_id == null || day == null || period == null) {
      return res.status(400).json({ error: 'بيانات الخانة ناقصة' });
    }
    if (day < 0 || day >= TT_PERIODS_PER_DAY.length || period < 1 || period > TT_PERIODS_PER_DAY[day]) {
      return res.status(400).json({ error: 'هذه الخانة غير موجودة في جدول هذا اليوم' });
    }
    if (!subject_id) {
      await pool.query('DELETE FROM tt_slots WHERE class_id=$1 AND day=$2 AND period=$3', [class_id, day, period]);
      return res.json({ ok: true, cleared: true });
    }
    const r = await pool.query(
      `INSERT INTO tt_slots (class_id, day, period, subject_id, teacher) VALUES ($1,$2,$3,$4,$5)
       ON CONFLICT (class_id, day, period) DO UPDATE SET subject_id=$4, teacher=$5 RETURNING *`,
      [class_id, day, period, subject_id, teacher || null]
    );
    res.json(r.rows[0]);
  } catch (err) { console.error(err); res.status(500).json({ error: 'تعذر حفظ الخانة' }); }
});

// -- Automatic generation (greedy + randomized restarts) --
function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function runGenerationAttempt(activities, periodsPerDay) {
  const acts = shuffle(activities);
  const classBusy = new Set();
  const teacherBusy = new Set();
  const classDaySubjectCount = {};
  const placements = [];
  const unplaced = [];

  for (const act of acts) {
    const candidates = [];
    for (let d = 0; d < periodsPerDay.length; d++) {
      for (let p = 1; p <= periodsPerDay[d]; p++) candidates.push([d, p]);
    }
    // Prefer days where this subject hasn't already been placed for this class today
    candidates.sort((a, b) => {
      const ca = ((classDaySubjectCount[act.class_id] || {})[a[0]] || {})[act.subject_id] || 0;
      const cb = ((classDaySubjectCount[act.class_id] || {})[b[0]] || {})[act.subject_id] || 0;
      if (ca !== cb) return ca - cb;
      return Math.random() - 0.5;
    });

    let placed = false;
    for (const [d, p] of candidates) {
      const ck = `c${act.class_id}-${d}-${p}`;
      const tk = `t${act.teacher}-${d}-${p}`;
      if (classBusy.has(ck) || teacherBusy.has(tk)) continue;
      classBusy.add(ck);
      teacherBusy.add(tk);
      classDaySubjectCount[act.class_id] = classDaySubjectCount[act.class_id] || {};
      classDaySubjectCount[act.class_id][d] = classDaySubjectCount[act.class_id][d] || {};
      classDaySubjectCount[act.class_id][d][act.subject_id] =
        (classDaySubjectCount[act.class_id][d][act.subject_id] || 0) + 1;
      placements.push({ class_id: act.class_id, day: d, period: p, subject_id: act.subject_id, teacher: act.teacher });
      placed = true;
      break;
    }
    if (!placed) unplaced.push(act);
  }
  return { placements, unplaced };
}

app.post('/api/tt/generate', checkPin, async (req, res) => {
  try {
    const reqs = (await pool.query('SELECT * FROM tt_requirements')).rows;
    if (!reqs.length) return res.status(400).json({ error: 'لا توجد متطلبات مواد محفوظة بعد' });

    const activities = [];
    reqs.forEach(r => {
      for (let i = 0; i < r.weekly_periods; i++) {
        activities.push({ class_id: r.class_id, subject_id: r.subject_id, teacher: r.teacher });
      }
    });

    let best = null;
    const ATTEMPTS = 120;
    for (let i = 0; i < ATTEMPTS; i++) {
      const result = runGenerationAttempt(activities, TT_PERIODS_PER_DAY);
      if (!best || result.unplaced.length < best.unplaced.length) {
        best = result;
        if (best.unplaced.length === 0) break;
      }
    }

    await pool.query('DELETE FROM tt_slots');
    for (const p of best.placements) {
      await pool.query(
        `INSERT INTO tt_slots (class_id, day, period, subject_id, teacher) VALUES ($1,$2,$3,$4,$5)`,
        [p.class_id, p.day, p.period, p.subject_id, p.teacher]
      );
    }

    const unplacedDetails = await Promise.all(best.unplaced.map(async u => {
      const c = await pool.query('SELECT grade, section FROM tt_classes WHERE id=$1', [u.class_id]);
      const s = await pool.query('SELECT name FROM tt_subjects WHERE id=$1', [u.subject_id]);
      return {
        class: c.rows[0] ? `${c.rows[0].grade} - ${c.rows[0].section}` : u.class_id,
        subject: s.rows[0] ? s.rows[0].name : u.subject_id,
        teacher: u.teacher
      };
    }));

    res.json({ placed: best.placements.length, unplaced: unplacedDetails });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'تعذر توليد الجدول' });
  }
});

app.get('/api/tt/export', checkPin, async (req, res) => {
  try {
    const classes = (await pool.query('SELECT * FROM tt_classes ORDER BY grade, section')).rows;
    const slots = (await pool.query(`
      SELECT sl.*, s.name AS subject_name FROM tt_slots sl
      LEFT JOIN tt_subjects s ON s.id = sl.subject_id
    `)).rows;

    const wb = new ExcelJS.Workbook();
    for (const cls of classes) {
      const ws = wb.addWorksheet(`${cls.grade}-${cls.section}`.slice(0, 28), { views: [{ rightToLeft: true }] });
      ws.getCell(1, 1).value = 'الحصة \\ اليوم';
      TT_DAY_NAMES.forEach((d, i) => { ws.getCell(1, i + 2).value = d; ws.getCell(1, i + 2).font = { bold: true }; });
      ws.getColumn(1).width = 12;
      for (let i = 2; i <= TT_DAY_NAMES.length + 1; i++) ws.getColumn(i).width = 20;
      for (let p = 1; p <= TT_MAX_PERIODS; p++) {
        ws.getCell(p + 1, 1).value = 'حصة ' + p;
        ws.getCell(p + 1, 1).font = { bold: true };
      }
      const bySlot = {};
      slots.filter(s => s.class_id === cls.id).forEach(s => { bySlot[`${s.day}-${s.period}`] = s; });
      for (let d = 0; d < TT_DAY_NAMES.length; d++) {
        for (let p = 1; p <= TT_PERIODS_PER_DAY[d]; p++) {
          const s = bySlot[`${d}-${p}`];
          if (s) ws.getCell(p + 1, d + 2).value = `${s.subject_name || ''}\n${s.teacher || ''}`;
        }
        for (let p = TT_PERIODS_PER_DAY[d] + 1; p <= TT_MAX_PERIODS; p++) {
          const cell = ws.getCell(p + 1, d + 2);
          cell.value = '—';
          cell.font = { color: { argb: 'FFBBBBBB' } };
        }
      }
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const utf8Name = encodeURIComponent('الجدول_المدرسي.xlsx');
    res.setHeader('Content-Disposition', `attachment; filename="school-timetable.xlsx"; filename*=UTF-8''${utf8Name}`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'تعذر تصدير الجدول' });
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
