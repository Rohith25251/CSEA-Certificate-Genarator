-- Create the invitation logging table
CREATE TABLE IF NOT EXISTS public.invitation (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  recipient_name TEXT NOT NULL,
  recipient_email TEXT NOT NULL,
  event_name TEXT,
  subject TEXT,
  status TEXT DEFAULT 'pending', -- 'sent', 'failed', 'pending'
  error_message TEXT,
  custom_data JSONB DEFAULT '{}'::jsonb,
  sent_at TIMESTAMPTZ DEFAULT NOW(),
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Enable Row-Level Security (RLS)
ALTER TABLE public.invitation ENABLE ROW LEVEL SECURITY;

-- Drop existing policies if any
DROP POLICY IF EXISTS "Enable full access for invitation" ON public.invitation;

-- Create permissive RLS policies matching the rest of the workspace tables
CREATE POLICY "Enable full access for invitation"
ON public.invitation
FOR ALL
TO public, anon, authenticated, service_role
USING (true)
WITH CHECK (true);
