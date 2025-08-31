-- Grant admin users full access to all features
-- Admin access bypasses operation limits

-- First, ensure you have admin status
UPDATE profiles 
SET is_admin = true,
    subscription_tier = 'beta',  -- Give yourself beta tier as well
    operations_limit = 99999      -- Unlimited operations for admin
WHERE email = 'fdsmith.mail@gmail.com';

-- Create function to check if user has full access (admin or unlimited)
CREATE OR REPLACE FUNCTION has_full_access()
RETURNS BOOLEAN AS $$
BEGIN
  RETURN EXISTS (
    SELECT 1 FROM profiles 
    WHERE id = auth.uid() 
    AND (is_admin = true OR operations_limit >= 99999)
  );
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Create function to bypass operation limits for admins
CREATE OR REPLACE FUNCTION check_operation_limit()
RETURNS BOOLEAN AS $$
DECLARE
  user_profile RECORD;
BEGIN
  SELECT * INTO user_profile
  FROM profiles
  WHERE id = auth.uid();
  
  -- Admins bypass all limits
  IF user_profile.is_admin THEN
    RETURN true;
  END IF;
  
  -- Regular users check against their limit
  RETURN user_profile.operations_used < user_profile.operations_limit;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Update the trigger to give admins unlimited operations
CREATE OR REPLACE FUNCTION update_operations_limit()
RETURNS TRIGGER AS $$
BEGIN
  -- Admins always get unlimited
  IF NEW.is_admin THEN
    NEW.operations_limit := 99999;
    RETURN NEW;
  END IF;
  
  -- Regular tier-based limits
  CASE NEW.subscription_tier
    WHEN 'beta' THEN 
      NEW.operations_limit := 1000;
    WHEN 'pro' THEN 
      NEW.operations_limit := 5000;
    WHEN 'enterprise' THEN 
      NEW.operations_limit := 99999;
    ELSE 
      NEW.operations_limit := 25;
  END CASE;
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Recreate the trigger
DROP TRIGGER IF EXISTS update_operations_limit_trigger ON profiles;
CREATE TRIGGER update_operations_limit_trigger
  BEFORE INSERT OR UPDATE ON profiles
  FOR EACH ROW
  EXECUTE FUNCTION update_operations_limit();

-- Grant necessary permissions
GRANT EXECUTE ON FUNCTION has_full_access TO authenticated;
GRANT EXECUTE ON FUNCTION check_operation_limit TO authenticated;

-- Verify admin setup
SELECT email, is_admin, subscription_tier, operations_limit 
FROM profiles 
WHERE email = 'fdsmith.mail@gmail.com';