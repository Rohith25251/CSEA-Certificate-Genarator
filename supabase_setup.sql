-- ========================================================
-- CSEA CERTIFICATE GENERATOR - FULL RLS & STORAGE BUCKET SQL
-- Run this query in your Supabase SQL Editor:
-- Dashboard -> SQL Editor -> New Query -> Run
-- ========================================================

-- 1. Grant schema usage permissions
GRANT USAGE ON SCHEMA public TO anon, authenticated, service_role;

-- 2. Grant table access to anon, authenticated and service_role
GRANT ALL ON ALL TABLES IN SCHEMA public TO anon, authenticated, service_role;
GRANT ALL ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated, service_role;
GRANT ALL ON ALL ROUTINES IN SCHEMA public TO anon, authenticated, service_role;

-- 3. Set default privileges for future tables
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT ALL ON TABLES TO anon, authenticated, service_role;
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT ALL ON SEQUENCES TO anon, authenticated, service_role;
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT ALL ON ROUTINES TO anon, authenticated, service_role;

-- 4. Create Tables if they don't exist yet
CREATE TABLE IF NOT EXISTS public.events (
  event_id TEXT PRIMARY KEY,
  event_name TEXT NOT NULL,
  event_category TEXT,
  event_date DATE DEFAULT CURRENT_DATE,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS public.students (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  register_no TEXT UNIQUE NOT NULL,
  name TEXT NOT NULL,
  email TEXT NOT NULL,
  phone TEXT,
  department TEXT DEFAULT 'Computer Science and Engineering',
  year_of_study TEXT ,
  section TEXT,
  college_name TEXT DEFAULT 'Kongu Engineering College',
  created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS public.certificates (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  student_id UUID REFERENCES public.students(id) ON DELETE CASCADE,
  event_id TEXT REFERENCES public.events(event_id) ON DELETE CASCADE,
  student_name TEXT NOT NULL,
  student_email TEXT NOT NULL,
  issue_date DATE DEFAULT CURRENT_DATE,
  email_status TEXT DEFAULT 'pending',
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. ENABLE Row Level Security (RLS) on all 3 tables
ALTER TABLE public.events ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.students ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.certificates ENABLE ROW LEVEL SECURITY;

-- 6. Drop existing policies if any to avoid conflicts
DROP POLICY IF EXISTS "Enable full access for events" ON public.events;
DROP POLICY IF EXISTS "Enable full access for students" ON public.students;
DROP POLICY IF EXISTS "Enable full access for certificates" ON public.certificates;

-- 7. Create Permissive RLS Policies for events (SELECT, INSERT, UPDATE, DELETE)
CREATE POLICY "Enable full access for events"
ON public.events
FOR ALL
TO public, anon, authenticated, service_role
USING (true)
WITH CHECK (true);

-- 8. Create Permissive RLS Policies for students (SELECT, INSERT, UPDATE, DELETE)
CREATE POLICY "Enable full access for students"
ON public.students
FOR ALL
TO public, anon, authenticated, service_role
USING (true)
WITH CHECK (true);

-- 9. Create Permissive RLS Policies for certificates (SELECT, INSERT, UPDATE, DELETE)
CREATE POLICY "Enable full access for certificates"
ON public.certificates
FOR ALL
TO public, anon, authenticated, service_role
USING (true)
WITH CHECK (true);

-- ========================================================
-- 10. SUPABASE STORAGE BUCKET FOR TEMPLATES & POLICIES
-- ========================================================
INSERT INTO storage.buckets (id, name, public)
VALUES ('templates', 'templates', true)
ON CONFLICT (id) DO NOTHING;

-- Drop existing storage policies if any
DROP POLICY IF EXISTS "Allow public uploads to templates" ON storage.objects;
DROP POLICY IF EXISTS "Allow public updates to templates" ON storage.objects;
DROP POLICY IF EXISTS "Allow public downloads from templates" ON storage.objects;

-- Create permissive storage policies for templates bucket (INSERT, UPDATE, SELECT)
CREATE POLICY "Allow public uploads to templates"
ON storage.objects
FOR INSERT
TO public, anon, authenticated, service_role
WITH CHECK (bucket_id = 'templates');

CREATE POLICY "Allow public updates to templates"
ON storage.objects
FOR UPDATE
TO public, anon, authenticated, service_role
USING (bucket_id = 'templates')
WITH CHECK (bucket_id = 'templates');

CREATE POLICY "Allow public downloads from templates"
ON storage.objects
FOR SELECT
TO public, anon, authenticated, service_role
USING (bucket_id = 'templates');

-- ========================================================
-- 11. APP SETTINGS TABLE FOR EMAIL TEMPLATE CUSTOMIZATIONS
-- ========================================================
CREATE TABLE IF NOT EXISTS public.app_settings (
  key TEXT PRIMARY KEY,
  value JSONB NOT NULL,
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE public.app_settings ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Enable full access for app_settings" ON public.app_settings;
CREATE POLICY "Enable full access for app_settings"
ON public.app_settings
FOR ALL
TO public, anon, authenticated, service_role
USING (true)
WITH CHECK (true);

