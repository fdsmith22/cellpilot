-- Create a function to check beta access by email
-- This function can be called publicly but only returns subscription tier info
CREATE OR REPLACE FUNCTION check_beta_access_by_email(user_email TEXT)
RETURNS JSON
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  result JSON;
BEGIN
  SELECT json_build_object(
    'hasBeta', CASE WHEN p.subscription_tier = 'beta' THEN true ELSE false END,
    'subscriptionTier', COALESCE(p.subscription_tier, 'none'),
    'found', true
  ) INTO result
  FROM auth.users u
  JOIN profiles p ON p.id = u.id
  WHERE LOWER(u.email) = LOWER(user_email)
  LIMIT 1;
  
  IF result IS NULL THEN
    RETURN json_build_object(
      'hasBeta', false,
      'subscriptionTier', 'none',
      'found', false
    );
  END IF;
  
  RETURN result;
END;
$$;

-- Grant execute permission to anon role so it can be called publicly
GRANT EXECUTE ON FUNCTION check_beta_access_by_email(TEXT) TO anon;

-- Add comment for documentation
COMMENT ON FUNCTION check_beta_access_by_email(TEXT) IS 'Check if a user has beta access by their email address';