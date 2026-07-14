const express = require('express');
const { Pool } = require('pg');
const nodemailer = require('nodemailer');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const multer = require('multer');
const { CloudinaryStorage } = require('multer-storage-cloudinary');
const cloudinary = require('cloudinary').v2;
const ExcelJS = require('exceljs');

const router = express.Router();

// ---------------------------------------------------------------------------
// Single shared PostgreSQL pool for the whole app (server.js reuses this —
// no more duplicate Pool instances).
// ---------------------------------------------------------------------------
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL && process.env.DATABASE_URL.includes('localhost')
    ? false
    : { rejectUnauthorized: false }
});

// ---------------------------------------------------------------------------
// Cloudinary + multer (certificate uploads)
// ---------------------------------------------------------------------------
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

const storage = new CloudinaryStorage({
  cloudinary,
  params: {
    folder: 'nextstep-applications',
    resource_type: 'auto', // allows PDFs as well as images
    allowed_formats: ['jpg', 'jpeg', 'png', 'pdf'],
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 }, // 5MB, matches the frontend hint
  fileFilter: (req, file, cb) => {
    const ok = ['image/jpeg', 'image/png', 'application/pdf'].includes(file.mimetype);
    cb(ok ? null : new Error('Only JPG, PNG, or PDF files are allowed.'), ok);
  },
});

const certificateUpload = upload.fields([
  { name: 'ssc_certificate', maxCount: 1 },
  { name: 'hsc_certificate', maxCount: 1 },
  { name: 'honours_certificate', maxCount: 1 },
  { name: 'masters_certificate', maxCount: 1 },
  { name: 'ielts_certificate', maxCount: 1 },
]);

// Separate storage/upload config for event banner images (own Cloudinary folder).
const eventImageStorage = new CloudinaryStorage({
  cloudinary,
  params: {
    folder: 'nextstep-events',
    resource_type: 'image',
    allowed_formats: ['jpg', 'jpeg', 'png', 'webp'],
  },
});

const eventImageUpload = multer({
  storage: eventImageStorage,
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = ['image/jpeg', 'image/png', 'image/webp'].includes(file.mimetype);
    cb(ok ? null : new Error('Only JPG, PNG, or WEBP images are allowed.'), ok);
  },
}).single('image');

// Combined storage/upload config for destinations: a photo (image) AND an
// optional process document (pdf). Cloudinary folder/resource_type is picked
// per-file based on which field it came in on.
const destinationStorage = new CloudinaryStorage({
  cloudinary,
  params: async (req, file) => {
    if (file.fieldname === 'pdf') {
      // Raw resources on Cloudinary do NOT auto-append a file extension to
      // the delivery URL the way images do — the "format" option here is
      // ignored for naming. Without an explicit ".pdf" in the public_id,
      // downloaded files have no extension, so phones/apps can't tell what
      // they are and refuse to open them (shows as "unrecognized format").
      // Building the public_id ourselves guarantees the URL, and therefore
      // the downloaded file name, always ends in .pdf.
      const base = (file.originalname || 'document')
        .replace(/\.pdf$/i, '')
        .replace(/[^a-zA-Z0-9-_]+/g, '-')
        .slice(0, 60);
      return {
        folder: 'nextstep-destinations-pdfs',
        resource_type: 'raw',
        public_id: `${base}-${Date.now()}.pdf`,
        use_filename: false,
        unique_filename: false,
      };
    }
    return {
      folder: 'nextstep-destinations',
      resource_type: 'image',
      allowed_formats: ['jpg', 'jpeg', 'png', 'webp'],
    };
  },
});

const destinationUpload = multer({
  storage: destinationStorage,
  limits: { fileSize: 8 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.fieldname === 'pdf') {
      const ok = file.mimetype === 'application/pdf';
      return cb(ok ? null : new Error('The process document must be a PDF file.'), ok);
    }
    const ok = ['image/jpeg', 'image/png', 'image/webp'].includes(file.mimetype);
    cb(ok ? null : new Error('Only JPG, PNG, or WEBP images are allowed.'), ok);
  },
}).fields([
  { name: 'image', maxCount: 1 },
  { name: 'pdf', maxCount: 1 },
]);

// ---------------------------------------------------------------------------
// Table setup
// ---------------------------------------------------------------------------
(async () => {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS contacts (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL,
        email TEXT NOT NULL,
        phone TEXT,
        message TEXT,
        created_at TIMESTAMP DEFAULT NOW()
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS student_applications (
        id SERIAL PRIMARY KEY,
        full_name TEXT NOT NULL,
        email TEXT NOT NULL,
        phone TEXT,
        ssc_board TEXT,
        ssc_year INTEGER,
        ssc_gpa NUMERIC(3,2),
        ssc_certificate_url TEXT,
        hsc_board TEXT,
        hsc_year INTEGER,
        hsc_gpa NUMERIC(3,2),
        hsc_certificate_url TEXT,
        honours_subject TEXT,
        honours_institution TEXT,
        honours_result TEXT,
        honours_year INTEGER,
        honours_certificate_url TEXT,
        masters_subject TEXT,
        masters_institution TEXT,
        masters_result TEXT,
        masters_year INTEGER,
        masters_certificate_url TEXT,
        ielts_score TEXT,
        ielts_test_date DATE,
        ielts_certificate_url TEXT,
        status TEXT NOT NULL DEFAULT 'pending',
        created_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Safe migration for databases that already had student_applications
    // before the IELTS fields were introduced.
    await pool.query(`
      ALTER TABLE student_applications
        ADD COLUMN IF NOT EXISTS ielts_score TEXT,
        ADD COLUMN IF NOT EXISTS ielts_test_date DATE,
        ADD COLUMN IF NOT EXISTS ielts_certificate_url TEXT;
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS events (
        id SERIAL PRIMARY KEY,
        title TEXT NOT NULL,
        subtitle TEXT,
        description TEXT,
        location TEXT,
        event_date DATE,
        event_time TEXT,
        image_url TEXT,
        badge_label TEXT,
        cta_label TEXT DEFAULT 'Learn more',
        cta_link TEXT,
        is_active BOOLEAN NOT NULL DEFAULT true,
        display_order INTEGER NOT NULL DEFAULT 0,
        created_at TIMESTAMP DEFAULT NOW()
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS destinations (
        id SERIAL PRIMARY KEY,
        country TEXT NOT NULL,
        flag_emoji TEXT,
        title TEXT NOT NULL,
        description TEXT,
        rating NUMERIC(2,1),
        image_url TEXT,
        pdf_url TEXT,
        cta_label TEXT DEFAULT 'View tour',
        cta_link TEXT,
        is_active BOOLEAN NOT NULL DEFAULT true,
        display_order INTEGER NOT NULL DEFAULT 0,
        created_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Safe migration for databases that already had destinations before the
    // per-destination process-document (PDF) field was introduced.
    await pool.query(`
      ALTER TABLE destinations
        ADD COLUMN IF NOT EXISTS pdf_url TEXT;
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS site_settings (
        key TEXT PRIMARY KEY,
        value TEXT
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS invoices (
        id SERIAL PRIMARY KEY,
        invoice_number TEXT NOT NULL UNIQUE,
        invoice_date DATE NOT NULL DEFAULT CURRENT_DATE,
        due_date DATE,
        client_name TEXT NOT NULL,
        client_email TEXT,
        client_phone TEXT,
        client_address TEXT,
        items JSONB NOT NULL DEFAULT '[]',
        discount NUMERIC(12,2) NOT NULL DEFAULT 0,
        tax_percent NUMERIC(5,2) NOT NULL DEFAULT 0,
        currency TEXT NOT NULL DEFAULT 'BDT',
        notes TEXT,
        status TEXT NOT NULL DEFAULT 'unpaid',
        created_at TIMESTAMP DEFAULT NOW()
      );
    `);

    console.log('Connected to PostgreSQL — contacts, student_applications, events, destinations, site_settings & invoices tables ready.');
  } catch (err) {
    console.error('Database connection/setup error:', err.message);
  }
})();

// ---------------------------------------------------------------------------
// Email transporter (used for contact form + optional application notices)
// ---------------------------------------------------------------------------
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
  }
});

// ---------------------------------------------------------------------------
// Admin auth
//
// Set these in .env:
//   JWT_SECRET=<long random string>
//   ADMIN_USERNAME=youradmin
//   ADMIN_PASSWORD_HASH=<bcrypt hash — generate with the snippet below>
//
// To generate a hash, run this once locally:
//   node -e "console.log(require('bcryptjs').hashSync('your-password', 10))"
// and paste the output into ADMIN_PASSWORD_HASH.
// ---------------------------------------------------------------------------
function requireAdmin(req, res, next) {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;

  if (!token) {
    return res.status(401).json({ error: 'Missing or invalid authorization header.' });
  }
  try {
    req.admin = jwt.verify(token, process.env.JWT_SECRET);
    next();
  } catch (err) {
    return res.status(401).json({ error: 'Invalid or expired session. Please log in again.' });
  }
}

router.post('/admin/login', async (req, res) => {
  const { username, password } = req.body;

  if (!username || !password) {
    return res.status(400).json({ error: 'Username and password are required.' });
  }
  if (!process.env.JWT_SECRET || !process.env.ADMIN_USERNAME || !process.env.ADMIN_PASSWORD_HASH) {
    console.error('Admin auth is not configured — missing JWT_SECRET / ADMIN_USERNAME / ADMIN_PASSWORD_HASH.');
    return res.status(500).json({ error: 'Admin login is not configured on the server.' });
  }

  try {
    const validUsername = username === process.env.ADMIN_USERNAME;
    const validPassword = validUsername
      ? await bcrypt.compare(password, process.env.ADMIN_PASSWORD_HASH)
      : false;

    if (!validUsername || !validPassword) {
      return res.status(401).json({ error: 'Invalid username or password.' });
    }

    const token = jwt.sign({ username }, process.env.JWT_SECRET, { expiresIn: '12h' });
    res.json({ token, expiresIn: '12h' });
  } catch (err) {
    console.error('Admin login error:', err.message);
    res.status(500).json({ error: 'Login failed. Please try again.' });
  }
});

// Lets the admin panel verify a stored token is still valid on page load.
router.get('/admin/me', requireAdmin, (req, res) => {
  res.json({ username: req.admin.username });
});

// ---------------------------------------------------------------------------
// Public: contact form
// ---------------------------------------------------------------------------
router.post('/contact', async (req, res) => {
  const { name, email, phone, message } = req.body;

  if (!name || !email || !message) {
    return res.status(400).json({ error: 'Name, email, and message are required.' });
  }

  try {
    const result = await pool.query(
      `INSERT INTO contacts (name, email, phone, message) VALUES ($1, $2, $3, $4) RETURNING id`,
      [name, email, phone || null, message]
    );

    transporter.sendMail({
      from: process.env.EMAIL_USER,
      to: process.env.EMAIL_USER,
      subject: `New Contact Form Submission from ${name}`,
      text: `Name: ${name}\nEmail: ${email}\nPhone: ${phone || 'N/A'}\n\nMessage:\n${message}`
    }).catch(err => console.error('Contact notification email failed:', err.message));

    res.status(200).json({ success: true, id: result.rows[0].id });
  } catch (err) {
    console.error('Contact form error:', err.message);
    res.status(500).json({ error: 'Something went wrong. Please try again later.' });
  }
});

// ---------------------------------------------------------------------------
// Public: upcoming events (only active ones, admin-controlled)
// ---------------------------------------------------------------------------
router.get('/events', async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT id, title, subtitle, description, location, event_date, event_time,
              image_url, badge_label, cta_label, cta_link
       FROM events
       WHERE is_active = true
       ORDER BY display_order ASC, event_date ASC NULLS LAST, created_at ASC`
    );
    res.json(result.rows);
  } catch (err) {
    console.error('Fetch public events error:', err.message);
    res.status(500).json({ error: 'Could not load events.' });
  }
});

// ---------------------------------------------------------------------------
// Public: popular destinations
// ---------------------------------------------------------------------------
router.get('/destinations', async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT id, country, flag_emoji, title, description, rating,
              image_url, pdf_url, cta_label, cta_link
       FROM destinations
       WHERE is_active = true
       ORDER BY display_order ASC, created_at ASC`
    );
    res.json(result.rows);
  } catch (err) {
    console.error('Fetch public destinations error:', err.message);
    res.status(500).json({ error: 'Could not load destinations.' });
  }
});

// ---------------------------------------------------------------------------
// Public: site settings (currently just WhatsApp chat config)
// ---------------------------------------------------------------------------
router.get('/settings', async (req, res) => {
  try {
    const result = await pool.query('SELECT key, value FROM site_settings');
    const settings = result.rows.reduce((acc, r) => ({ ...acc, [r.key]: r.value }), {});
    res.json({
      whatsapp_enabled: settings.whatsapp_enabled === 'true',
      whatsapp_number: settings.whatsapp_number || '',
      whatsapp_message: settings.whatsapp_message || 'Hi! I would like to know more about your services.',
    });
  } catch (err) {
    console.error('Fetch public settings error:', err.message);
    res.status(500).json({ error: 'Could not load settings.' });
  }
});

// ---------------------------------------------------------------------------
// Public: student application (with certificate uploads to Cloudinary)
// ---------------------------------------------------------------------------
router.post('/student-application', (req, res) => {
  certificateUpload(req, res, async (uploadErr) => {
    if (uploadErr) {
      console.error('Upload error:', uploadErr.message);
      const message = uploadErr.code === 'LIMIT_FILE_SIZE'
        ? 'One of your files is larger than 5MB.'
        : uploadErr.message || 'File upload failed.';
      return res.status(400).json({ error: message });
    }

    const b = req.body;
    if (!b.full_name || !b.email) {
      return res.status(400).json({ error: 'Full name and email are required.' });
    }

    const files = req.files || {};
    const urlOf = (field) => (files[field] && files[field][0] ? files[field][0].path : null);
    const toIntOrNull = (v) => (v === undefined || v === null || v === '' ? null : parseInt(v, 10));
    const toNumOrNull = (v) => (v === undefined || v === null || v === '' ? null : parseFloat(v));

    try {
      const result = await pool.query(
        `INSERT INTO student_applications
          (full_name, email, phone,
           ssc_board, ssc_year, ssc_gpa, ssc_certificate_url,
           hsc_board, hsc_year, hsc_gpa, hsc_certificate_url,
           honours_subject, honours_institution, honours_result, honours_year, honours_certificate_url,
           masters_subject, masters_institution, masters_result, masters_year, masters_certificate_url,
           ielts_score, ielts_test_date, ielts_certificate_url)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24)
         RETURNING id`,
        [
          b.full_name, b.email, b.phone || null,
          b.ssc_board || null, toIntOrNull(b.ssc_year), toNumOrNull(b.ssc_gpa), urlOf('ssc_certificate'),
          b.hsc_board || null, toIntOrNull(b.hsc_year), toNumOrNull(b.hsc_gpa), urlOf('hsc_certificate'),
          b.honours_subject || null, b.honours_institution || null, b.honours_result || null, toIntOrNull(b.honours_year), urlOf('honours_certificate'),
          b.masters_subject || null, b.masters_institution || null, b.masters_result || null, toIntOrNull(b.masters_year), urlOf('masters_certificate'),
          b.ielts_score || null, b.ielts_test_date || null, urlOf('ielts_certificate'),
        ]
      );

      transporter.sendMail({
        from: process.env.EMAIL_USER,
        to: process.env.EMAIL_USER,
        subject: `New Student Application — ${b.full_name}`,
        text: `New application #${result.rows[0].id}\nName: ${b.full_name}\nEmail: ${b.email}\nPhone: ${b.phone || 'N/A'}`
      }).catch(err => console.error('Application notification email failed:', err.message));

      res.status(200).json({ success: true, id: result.rows[0].id });
    } catch (err) {
      console.error('Student application error:', err.message);
      res.status(500).json({ error: 'Something went wrong. Please try again later.' });
    }
  });
});

// ---------------------------------------------------------------------------
// Admin (protected): contacts
// ---------------------------------------------------------------------------
router.get('/admin/contacts', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM contacts ORDER BY created_at DESC');
    res.json(result.rows);
  } catch (err) {
    console.error('Fetch contacts error:', err.message);
    res.status(500).json({ error: 'Could not fetch contacts.' });
  }
});

router.delete('/admin/contacts/:id', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('DELETE FROM contacts WHERE id = $1 RETURNING id', [req.params.id]);
    if (result.rowCount === 0) return res.status(404).json({ error: 'Contact not found.' });
    res.json({ success: true });
  } catch (err) {
    console.error('Delete contact error:', err.message);
    res.status(500).json({ error: 'Could not delete contact.' });
  }
});

// ---------------------------------------------------------------------------
// Admin (protected): student applications
// ---------------------------------------------------------------------------
const VALID_STATUSES = ['pending', 'under_review', 'approved', 'rejected'];

router.get('/admin/applications', requireAdmin, async (req, res) => {
  try {
    const { status } = req.query;
    const params = [];
    let query = 'SELECT * FROM student_applications';
    if (status && VALID_STATUSES.includes(status)) {
      params.push(status);
      query += ' WHERE status = $1';
    }
    query += ' ORDER BY created_at DESC';

    const result = await pool.query(query, params);
    res.json(result.rows);
  } catch (err) {
    console.error('Fetch applications error:', err.message);
    res.status(500).json({ error: 'Could not fetch applications.' });
  }
});

router.get('/admin/applications/:id', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM student_applications WHERE id = $1', [req.params.id]);
    if (result.rows.length === 0) return res.status(404).json({ error: 'Application not found.' });
    res.json(result.rows[0]);
  } catch (err) {
    console.error('Fetch application error:', err.message);
    res.status(500).json({ error: 'Could not fetch application.' });
  }
});

router.patch('/admin/applications/:id/status', requireAdmin, async (req, res) => {
  const { status } = req.body;
  if (!VALID_STATUSES.includes(status)) {
    return res.status(400).json({ error: `Status must be one of: ${VALID_STATUSES.join(', ')}` });
  }
  try {
    const result = await pool.query(
      'UPDATE student_applications SET status = $1 WHERE id = $2 RETURNING *',
      [status, req.params.id]
    );
    if (result.rows.length === 0) return res.status(404).json({ error: 'Application not found.' });
    res.json(result.rows[0]);
  } catch (err) {
    console.error('Update application status error:', err.message);
    res.status(500).json({ error: 'Could not update application status.' });
  }
});

router.delete('/admin/applications/:id', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('DELETE FROM student_applications WHERE id = $1 RETURNING id', [req.params.id]);
    if (result.rowCount === 0) return res.status(404).json({ error: 'Application not found.' });
    res.json({ success: true });
  } catch (err) {
    console.error('Delete application error:', err.message);
    res.status(500).json({ error: 'Could not delete application.' });
  }
});

// ---------------------------------------------------------------------------
// Admin (protected): dashboard summary stats
// ---------------------------------------------------------------------------
router.get('/admin/stats', requireAdmin, async (req, res) => {
  try {
    const [apps, byStatus, contacts] = await Promise.all([
      pool.query('SELECT COUNT(*)::int AS total FROM student_applications'),
      pool.query('SELECT status, COUNT(*)::int AS count FROM student_applications GROUP BY status'),
      pool.query('SELECT COUNT(*)::int AS total FROM contacts'),
    ]);
    res.json({
      totalApplications: apps.rows[0].total,
      totalContacts: contacts.rows[0].total,
      byStatus: byStatus.rows.reduce((acc, r) => ({ ...acc, [r.status]: r.count }), {}),
    });
  } catch (err) {
    console.error('Fetch stats error:', err.message);
    res.status(500).json({ error: 'Could not fetch stats.' });
  }
});

// ---------------------------------------------------------------------------
// Admin (protected): Excel export of all student applications
// ---------------------------------------------------------------------------
router.get('/admin/export-applications', requireAdmin, async (req, res) => {
  try {
    const { rows } = await pool.query(
      `SELECT id, full_name, email, phone,
              ssc_board, ssc_year, ssc_gpa, ssc_certificate_url,
              hsc_board, hsc_year, hsc_gpa, hsc_certificate_url,
              honours_subject, honours_institution, honours_result, honours_year, honours_certificate_url,
              masters_subject, masters_institution, masters_result, masters_year, masters_certificate_url,
              ielts_score, ielts_test_date, ielts_certificate_url,
              status, created_at
       FROM student_applications
       ORDER BY created_at DESC`
    );

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Nextstep Immigration';
    wb.created = new Date();

    const sheet = wb.addWorksheet('Applications', { views: [{ state: 'frozen', ySplit: 1 }] });

    sheet.columns = [
      { header: 'ID', key: 'id', width: 6 },
      { header: 'Full Name', key: 'full_name', width: 24 },
      { header: 'Email', key: 'email', width: 26 },
      { header: 'Phone', key: 'phone', width: 16 },
      { header: 'SSC Board', key: 'ssc_board', width: 14 },
      { header: 'SSC Year', key: 'ssc_year', width: 10 },
      { header: 'SSC GPA', key: 'ssc_gpa', width: 10 },
      { header: 'SSC Certificate', key: 'ssc_certificate_url', width: 30 },
      { header: 'HSC Board', key: 'hsc_board', width: 14 },
      { header: 'HSC Year', key: 'hsc_year', width: 10 },
      { header: 'HSC GPA', key: 'hsc_gpa', width: 10 },
      { header: 'HSC Certificate', key: 'hsc_certificate_url', width: 30 },
      { header: 'Honours Subject', key: 'honours_subject', width: 20 },
      { header: 'Honours Institution', key: 'honours_institution', width: 24 },
      { header: 'Honours Result', key: 'honours_result', width: 16 },
      { header: 'Honours Year', key: 'honours_year', width: 12 },
      { header: 'Honours Certificate', key: 'honours_certificate_url', width: 30 },
      { header: "Master's Subject", key: 'masters_subject', width: 20 },
      { header: "Master's Institution", key: 'masters_institution', width: 24 },
      { header: "Master's Result", key: 'masters_result', width: 16 },
      { header: "Master's Year", key: 'masters_year', width: 12 },
      { header: "Master's Certificate", key: 'masters_certificate_url', width: 30 },
      { header: 'IELTS Score', key: 'ielts_score', width: 12 },
      { header: 'IELTS Test Date', key: 'ielts_test_date', width: 16 },
      { header: 'IELTS Certificate', key: 'ielts_certificate_url', width: 30 },
      { header: 'Status', key: 'status', width: 12 },
      { header: 'Submitted At', key: 'created_at', width: 20 },
    ];

    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF12213B' } };
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFF6F4EE' } };

    rows.forEach((r) => {
      sheet.addRow({
        ...r,
        ielts_test_date: r.ielts_test_date ? new Date(r.ielts_test_date).toLocaleDateString() : '',
        created_at: r.created_at ? new Date(r.created_at).toLocaleString() : '',
      });
    });

    sheet.autoFilter = { from: 'A1', to: `${sheet.getColumn(sheet.columns.length).letter}1` };

    const filename = `student-applications-${new Date().toISOString().slice(0, 10)}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Export failed:', err);
    res.status(500).json({ error: 'Failed to generate Excel export.' });
  }
});

// ---------------------------------------------------------------------------
// Admin (protected): upcoming events management
// ---------------------------------------------------------------------------
router.get('/admin/events', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM events ORDER BY display_order ASC, event_date ASC NULLS LAST, created_at DESC');
    res.json(result.rows);
  } catch (err) {
    console.error('Fetch admin events error:', err.message);
    res.status(500).json({ error: 'Could not fetch events.' });
  }
});

router.post('/admin/events', requireAdmin, (req, res) => {
  eventImageUpload(req, res, async (uploadErr) => {
    if (uploadErr) {
      console.error('Event image upload error:', uploadErr.message);
      const message = uploadErr.code === 'LIMIT_FILE_SIZE'
        ? 'Image is larger than 5MB.'
        : uploadErr.message || 'Image upload failed.';
      return res.status(400).json({ error: message });
    }

    const b = req.body;
    if (!b.title) {
      return res.status(400).json({ error: 'Title is required.' });
    }

    try {
      const imageUrl = req.file ? req.file.path : (b.image_url || null);
      const result = await pool.query(
        `INSERT INTO events
          (title, subtitle, description, location, event_date, event_time, image_url,
           badge_label, cta_label, cta_link, is_active, display_order)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)
         RETURNING *`,
        [
          b.title,
          b.subtitle || null,
          b.description || null,
          b.location || null,
          b.event_date || null,
          b.event_time || null,
          imageUrl,
          b.badge_label || null,
          b.cta_label || 'Learn more',
          b.cta_link || null,
          b.is_active === undefined ? true : b.is_active === 'true' || b.is_active === true,
          Number.isFinite(Number(b.display_order)) ? Number(b.display_order) : 0,
        ]
      );
      res.status(201).json(result.rows[0]);
    } catch (err) {
      console.error('Create event error:', err.message);
      res.status(500).json({ error: 'Could not create event.' });
    }
  });
});

router.put('/admin/events/:id', requireAdmin, (req, res) => {
  eventImageUpload(req, res, async (uploadErr) => {
    if (uploadErr) {
      console.error('Event image upload error:', uploadErr.message);
      const message = uploadErr.code === 'LIMIT_FILE_SIZE'
        ? 'Image is larger than 5MB.'
        : uploadErr.message || 'Image upload failed.';
      return res.status(400).json({ error: message });
    }

    const b = req.body;
    if (!b.title) {
      return res.status(400).json({ error: 'Title is required.' });
    }

    try {
      const existing = await pool.query('SELECT image_url FROM events WHERE id = $1', [req.params.id]);
      if (existing.rows.length === 0) return res.status(404).json({ error: 'Event not found.' });

      const imageUrl = req.file ? req.file.path : (b.image_url || existing.rows[0].image_url);

      const result = await pool.query(
        `UPDATE events SET
          title=$1, subtitle=$2, description=$3, location=$4, event_date=$5, event_time=$6,
          image_url=$7, badge_label=$8, cta_label=$9, cta_link=$10, is_active=$11, display_order=$12
         WHERE id=$13
         RETURNING *`,
        [
          b.title,
          b.subtitle || null,
          b.description || null,
          b.location || null,
          b.event_date || null,
          b.event_time || null,
          imageUrl,
          b.badge_label || null,
          b.cta_label || 'Learn more',
          b.cta_link || null,
          b.is_active === undefined ? true : b.is_active === 'true' || b.is_active === true,
          Number.isFinite(Number(b.display_order)) ? Number(b.display_order) : 0,
          req.params.id,
        ]
      );
      res.json(result.rows[0]);
    } catch (err) {
      console.error('Update event error:', err.message);
      res.status(500).json({ error: 'Could not update event.' });
    }
  });
});

router.patch('/admin/events/:id/toggle', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query(
      'UPDATE events SET is_active = NOT is_active WHERE id = $1 RETURNING *',
      [req.params.id]
    );
    if (result.rows.length === 0) return res.status(404).json({ error: 'Event not found.' });
    res.json(result.rows[0]);
  } catch (err) {
    console.error('Toggle event error:', err.message);
    res.status(500).json({ error: 'Could not update event.' });
  }
});

router.delete('/admin/events/:id', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('DELETE FROM events WHERE id = $1 RETURNING id', [req.params.id]);
    if (result.rowCount === 0) return res.status(404).json({ error: 'Event not found.' });
    res.json({ success: true });
  } catch (err) {
    console.error('Delete event error:', err.message);
    res.status(500).json({ error: 'Could not delete event.' });
  }
});

// ---------------------------------------------------------------------------
// Admin (protected): popular destinations
// ---------------------------------------------------------------------------
router.get('/admin/destinations', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM destinations ORDER BY display_order ASC, created_at DESC');
    res.json(result.rows);
  } catch (err) {
    console.error('Fetch admin destinations error:', err.message);
    res.status(500).json({ error: 'Could not fetch destinations.' });
  }
});

router.post('/admin/destinations', requireAdmin, (req, res) => {
  destinationUpload(req, res, async (uploadErr) => {
    if (uploadErr) {
      console.error('Destination upload error:', uploadErr.message);
      const message = uploadErr.code === 'LIMIT_FILE_SIZE'
        ? 'File is larger than 8MB.'
        : uploadErr.message || 'Upload failed.';
      return res.status(400).json({ error: message });
    }

    const b = req.body;
    if (!b.country || !b.title) {
      return res.status(400).json({ error: 'Country and title are required.' });
    }

    try {
      const imageFile = req.files && req.files.image ? req.files.image[0] : null;
      const pdfFile = req.files && req.files.pdf ? req.files.pdf[0] : null;
      const imageUrl = imageFile ? imageFile.path : (b.image_url || null);
      const pdfUrl = pdfFile ? pdfFile.path : (b.pdf_url || null);
      const result = await pool.query(
        `INSERT INTO destinations
          (country, flag_emoji, title, description, rating, image_url, pdf_url,
           cta_label, cta_link, is_active, display_order)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
         RETURNING *`,
        [
          b.country,
          b.flag_emoji || null,
          b.title,
          b.description || null,
          b.rating === '' || b.rating === undefined ? null : Number(b.rating),
          imageUrl,
          pdfUrl,
          b.cta_label || 'View tour',
          b.cta_link || null,
          b.is_active === undefined ? true : b.is_active === 'true' || b.is_active === true,
          Number.isFinite(Number(b.display_order)) ? Number(b.display_order) : 0,
        ]
      );
      res.status(201).json(result.rows[0]);
    } catch (err) {
      console.error('Create destination error:', err.message);
      res.status(500).json({ error: 'Could not create destination.' });
    }
  });
});

router.put('/admin/destinations/:id', requireAdmin, (req, res) => {
  destinationUpload(req, res, async (uploadErr) => {
    if (uploadErr) {
      console.error('Destination upload error:', uploadErr.message);
      const message = uploadErr.code === 'LIMIT_FILE_SIZE'
        ? 'File is larger than 8MB.'
        : uploadErr.message || 'Upload failed.';
      return res.status(400).json({ error: message });
    }

    const b = req.body;
    if (!b.country || !b.title) {
      return res.status(400).json({ error: 'Country and title are required.' });
    }

    try {
      const existing = await pool.query('SELECT image_url, pdf_url FROM destinations WHERE id = $1', [req.params.id]);
      if (existing.rows.length === 0) return res.status(404).json({ error: 'Destination not found.' });

      const imageFile = req.files && req.files.image ? req.files.image[0] : null;
      const pdfFile = req.files && req.files.pdf ? req.files.pdf[0] : null;
      const imageUrl = imageFile ? imageFile.path : (b.image_url || existing.rows[0].image_url);
      const pdfUrl = pdfFile ? pdfFile.path : (b.pdf_url || existing.rows[0].pdf_url);

      const result = await pool.query(
        `UPDATE destinations SET
          country=$1, flag_emoji=$2, title=$3, description=$4, rating=$5,
          image_url=$6, pdf_url=$7, cta_label=$8, cta_link=$9, is_active=$10, display_order=$11
         WHERE id=$12
         RETURNING *`,
        [
          b.country,
          b.flag_emoji || null,
          b.title,
          b.description || null,
          b.rating === '' || b.rating === undefined ? null : Number(b.rating),
          imageUrl,
          pdfUrl,
          b.cta_label || 'View tour',
          b.cta_link || null,
          b.is_active === undefined ? true : b.is_active === 'true' || b.is_active === true,
          Number.isFinite(Number(b.display_order)) ? Number(b.display_order) : 0,
          req.params.id,
        ]
      );
      res.json(result.rows[0]);
    } catch (err) {
      console.error('Update destination error:', err.message);
      res.status(500).json({ error: 'Could not update destination.' });
    }
  });
});

router.patch('/admin/destinations/:id/toggle', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query(
      'UPDATE destinations SET is_active = NOT is_active WHERE id = $1 RETURNING *',
      [req.params.id]
    );
    if (result.rows.length === 0) return res.status(404).json({ error: 'Destination not found.' });
    res.json(result.rows[0]);
  } catch (err) {
    console.error('Toggle destination error:', err.message);
    res.status(500).json({ error: 'Could not update destination.' });
  }
});

router.delete('/admin/destinations/:id', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('DELETE FROM destinations WHERE id = $1 RETURNING id', [req.params.id]);
    if (result.rowCount === 0) return res.status(404).json({ error: 'Destination not found.' });
    res.json({ success: true });
  } catch (err) {
    console.error('Delete destination error:', err.message);
    res.status(500).json({ error: 'Could not delete destination.' });
  }
});

// ---------------------------------------------------------------------------
// Admin (protected): site settings (WhatsApp chat config)
// ---------------------------------------------------------------------------
router.get('/admin/settings', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('SELECT key, value FROM site_settings');
    const settings = result.rows.reduce((acc, r) => ({ ...acc, [r.key]: r.value }), {});
    res.json({
      whatsapp_enabled: settings.whatsapp_enabled === 'true',
      whatsapp_number: settings.whatsapp_number || '',
      whatsapp_message: settings.whatsapp_message || 'Hi! I would like to know more about your services.',
    });
  } catch (err) {
    console.error('Fetch admin settings error:', err.message);
    res.status(500).json({ error: 'Could not fetch settings.' });
  }
});

router.put('/admin/settings', requireAdmin, async (req, res) => {
  const { whatsapp_enabled, whatsapp_number, whatsapp_message } = req.body;
  try {
    const entries = [
      ['whatsapp_enabled', String(!!whatsapp_enabled)],
      ['whatsapp_number', whatsapp_number || ''],
      ['whatsapp_message', whatsapp_message || ''],
    ];
    for (const [key, value] of entries) {
      await pool.query(
        `INSERT INTO site_settings (key, value) VALUES ($1, $2)
         ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
        [key, value]
      );
    }
    res.json({ success: true });
  } catch (err) {
    console.error('Update settings error:', err.message);
    res.status(500).json({ error: 'Could not update settings.' });
  }
});

// ---------------------------------------------------------------------------
// Admin (protected): invoices
// ---------------------------------------------------------------------------
const VALID_INVOICE_STATUSES = ['unpaid', 'paid', 'overdue', 'draft'];

function computeInvoiceTotal(items, discount, taxPercent) {
  const subtotal = (items || []).reduce(
    (sum, it) => sum + (Number(it.qty) || 0) * (Number(it.unit_price) || 0),
    0
  );
  const afterDiscount = Math.max(subtotal - (Number(discount) || 0), 0);
  const total = afterDiscount + (afterDiscount * (Number(taxPercent) || 0)) / 100;
  return Math.round(total * 100) / 100;
}

router.get('/admin/invoices', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM invoices ORDER BY invoice_date DESC, id DESC');
    res.json(result.rows);
  } catch (err) {
    console.error('Fetch admin invoices error:', err.message);
    res.status(500).json({ error: 'Could not fetch invoices.' });
  }
});

// Suggests the next sequential invoice number, e.g. INV-0001 -> INV-0002.
router.get('/admin/invoices/next-number', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT invoice_number FROM invoices ORDER BY id DESC LIMIT 1`
    );
    let nextNum = 1;
    if (result.rows.length) {
      const match = result.rows[0].invoice_number.match(/(\d+)$/);
      if (match) nextNum = parseInt(match[1], 10) + 1;
    }
    res.json({ invoice_number: `INV-${String(nextNum).padStart(4, '0')}` });
  } catch (err) {
    console.error('Next invoice number error:', err.message);
    res.status(500).json({ error: 'Could not generate invoice number.' });
  }
});

router.post('/admin/invoices', requireAdmin, async (req, res) => {
  const b = req.body;
  if (!b.invoice_number || !b.client_name || !Array.isArray(b.items) || b.items.length === 0) {
    return res.status(400).json({ error: 'Invoice number, client name, and at least one item are required.' });
  }
  const status = VALID_INVOICE_STATUSES.includes(b.status) ? b.status : 'unpaid';

  try {
    const result = await pool.query(
      `INSERT INTO invoices
        (invoice_number, invoice_date, due_date, client_name, client_email, client_phone,
         client_address, items, discount, tax_percent, currency, notes, status)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13)
       RETURNING *`,
      [
        b.invoice_number,
        b.invoice_date || new Date().toISOString().slice(0, 10),
        b.due_date || null,
        b.client_name,
        b.client_email || null,
        b.client_phone || null,
        b.client_address || null,
        JSON.stringify(b.items),
        Number(b.discount) || 0,
        Number(b.tax_percent) || 0,
        b.currency || 'BDT',
        b.notes || null,
        status,
      ]
    );
    res.status(201).json(result.rows[0]);
  } catch (err) {
    console.error('Create invoice error:', err.message);
    if (err.code === '23505') {
      return res.status(409).json({ error: 'An invoice with that number already exists.' });
    }
    res.status(500).json({ error: 'Could not create invoice.' });
  }
});

router.put('/admin/invoices/:id', requireAdmin, async (req, res) => {
  const b = req.body;
  if (!b.invoice_number || !b.client_name || !Array.isArray(b.items) || b.items.length === 0) {
    return res.status(400).json({ error: 'Invoice number, client name, and at least one item are required.' });
  }
  const status = VALID_INVOICE_STATUSES.includes(b.status) ? b.status : 'unpaid';

  try {
    const result = await pool.query(
      `UPDATE invoices SET
        invoice_number=$1, invoice_date=$2, due_date=$3, client_name=$4, client_email=$5,
        client_phone=$6, client_address=$7, items=$8, discount=$9, tax_percent=$10,
        currency=$11, notes=$12, status=$13
       WHERE id=$14
       RETURNING *`,
      [
        b.invoice_number,
        b.invoice_date || new Date().toISOString().slice(0, 10),
        b.due_date || null,
        b.client_name,
        b.client_email || null,
        b.client_phone || null,
        b.client_address || null,
        JSON.stringify(b.items),
        Number(b.discount) || 0,
        Number(b.tax_percent) || 0,
        b.currency || 'BDT',
        b.notes || null,
        status,
        req.params.id,
      ]
    );
    if (result.rows.length === 0) return res.status(404).json({ error: 'Invoice not found.' });
    res.json(result.rows[0]);
  } catch (err) {
    console.error('Update invoice error:', err.message);
    if (err.code === '23505') {
      return res.status(409).json({ error: 'An invoice with that number already exists.' });
    }
    res.status(500).json({ error: 'Could not update invoice.' });
  }
});

router.delete('/admin/invoices/:id', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('DELETE FROM invoices WHERE id = $1 RETURNING id', [req.params.id]);
    if (result.rowCount === 0) return res.status(404).json({ error: 'Invoice not found.' });
    res.json({ success: true });
  } catch (err) {
    console.error('Delete invoice error:', err.message);
    res.status(500).json({ error: 'Could not delete invoice.' });
  }
});

module.exports = router;
module.exports.pool = pool;
