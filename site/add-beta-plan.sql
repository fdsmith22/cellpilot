-- Add beta plan type and admin capabilities
-- Beta is just a subscription tier, NOT admin access

-- Update subscription_tier to include beta option
ALTER TABLE profiles 
ALTER COLUMN subscription_tier 
SET DEFAULT 'free';

-- Add is_admin flag for admin users (separate from beta)
ALTER TABLE profiles 
ADD COLUMN IF NOT EXISTS is_admin BOOLEAN DEFAULT false;

-- Update operations limit based on tier
CREATE OR REPLACE FUNCTION update_operations_limit()
RETURNS TRIGGER AS $$
BEGIN
  CASE NEW.subscription_tier
    WHEN 'beta' THEN 
      NEW.operations_limit := 1000;  -- Beta users get 1000 operations
    WHEN 'pro' THEN 
      NEW.operations_limit := 5000;  -- Future pro plan
    WHEN 'enterprise' THEN 
      NEW.operations_limit := 99999; -- Future enterprise plan
    ELSE 
      NEW.operations_limit := 25;    -- Free tier
  END CASE;
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Create trigger to auto-update operations limit when tier changes
DROP TRIGGER IF EXISTS update_operations_limit_trigger ON profiles;
CREATE TRIGGER update_operations_limit_trigger
  BEFORE UPDATE OF subscription_tier ON profiles
  FOR EACH ROW
  EXECUTE FUNCTION update_operations_limit();

-- Make ONLY yourself an admin (not all beta users)
-- This is separate from subscription tier
UPDATE profiles 
SET is_admin = true 
WHERE email = 'fdsmith.mail@gmail.com';

-- Create a view for admins to see all users
CREATE OR REPLACE VIEW admin_user_view AS
SELECT 
  p.id,
  p.email,
  p.full_name,
  p.company,
  p.subscription_tier,
  p.operations_used,
  p.operations_limit,
  p.email_verified,
  p.newsletter_subscribed,
  p.created_at,
  p.last_login,
  p.is_admin,
  u.last_sign_in_at,
  u.email_confirmed_at
FROM profiles p
LEFT JOIN auth.users u ON p.id = u.id;

-- Grant access to admin view only for admin users
CREATE OR REPLACE FUNCTION is_admin()
RETURNS BOOLEAN AS $$
BEGIN
  RETURN EXISTS (
    SELECT 1 FROM profiles 
    WHERE id = auth.uid() 
    AND is_admin = true
  );
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Create policies for admin view
ALTER VIEW admin_user_view OWNER TO authenticated;

-- Create RLS policy for admin view
CREATE POLICY "Only admins can view admin_user_view" ON profiles
  FOR ALL
  USING (auth.uid() = id OR is_admin());

-- Function to update user tier (for admin use)
CREATE OR REPLACE FUNCTION admin_update_user_tier(
  target_user_id UUID,
  new_tier TEXT
)
RETURNS VOID AS $$
BEGIN
  -- Check if current user is admin
  IF NOT is_admin() THEN
    RAISE EXCEPTION 'Unauthorized: Admin access required';
  END IF;
  
  -- Update the user's tier
  UPDATE profiles 
  SET 
    subscription_tier = new_tier,
    updated_at = NOW()
  WHERE id = target_user_id;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Grant execute permission to authenticated users (will check admin internally)
GRANT EXECUTE ON FUNCTION admin_update_user_tier TO authenticated;
GRANT EXECUTE ON FUNCTION is_admin TO authenticated;