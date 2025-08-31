-- Add beta access fields to profiles
ALTER TABLE profiles 
ADD COLUMN IF NOT EXISTS beta_requested_at TIMESTAMP WITH TIME ZONE,
ADD COLUMN IF NOT EXISTS beta_approved_at TIMESTAMP WITH TIME ZONE,
ADD COLUMN IF NOT EXISTS beta_access BOOLEAN DEFAULT FALSE,
ADD COLUMN IF NOT EXISTS beta_notes TEXT;

-- Create index for admin queries
CREATE INDEX IF NOT EXISTS idx_beta_requests ON profiles(beta_requested_at) WHERE beta_requested_at IS NOT NULL;

-- Create a view for admin to see pending beta requests
CREATE OR REPLACE VIEW beta_requests AS
SELECT 
  id,
  email,
  full_name,
  beta_requested_at,
  beta_approved_at,
  beta_access,
  beta_notes,
  created_at
FROM profiles
WHERE beta_requested_at IS NOT NULL
ORDER BY beta_requested_at DESC;