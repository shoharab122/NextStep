CREATE TABLE IF NOT EXISTS student_applications (
  id SERIAL PRIMARY KEY,

  -- Personal info
  full_name TEXT NOT NULL,
  email TEXT NOT NULL,
  phone TEXT,

  -- SSC (Secondary School Certificate)
  ssc_board TEXT,
  ssc_year INTEGER,
  ssc_gpa NUMERIC(3,2),
  ssc_certificate_url TEXT,

  -- HSC (Higher Secondary Certificate)
  hsc_board TEXT,
  hsc_year INTEGER,
  hsc_gpa NUMERIC(3,2),
  hsc_certificate_url TEXT,

  -- Honours (Bachelor's)
  honours_subject TEXT,
  honours_institution TEXT,
  honours_result TEXT,
  honours_year INTEGER,
  honours_certificate_url TEXT,

  -- Master's
  masters_subject TEXT,
  masters_institution TEXT,
  masters_result TEXT,
  masters_year INTEGER,
  masters_certificate_url TEXT,

  status TEXT DEFAULT 'pending',
  created_at TIMESTAMP DEFAULT NOW()
);