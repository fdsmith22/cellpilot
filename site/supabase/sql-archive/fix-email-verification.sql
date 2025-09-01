-- Drop existing trigger and function
DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
DROP FUNCTION IF EXISTS public.handle_new_user();

-- Create improved function that properly handles email verification
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO public.profiles (
    id, 
    email, 
    email_verified,
    created_at,
    updated_at
  )
  VALUES (
    new.id, 
    new.email, 
    COALESCE(new.email_confirmed_at IS NOT NULL, false),
    NOW(),
    NOW()
  )
  ON CONFLICT (id) DO UPDATE
  SET 
    email = EXCLUDED.email,
    email_verified = COALESCE(new.email_confirmed_at IS NOT NULL, false),
    updated_at = NOW();
  
  RETURN new;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Create trigger for new users
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- Create function to update email verification status when email is confirmed
CREATE OR REPLACE FUNCTION public.handle_user_email_confirmed()
RETURNS TRIGGER AS $$
BEGIN
  -- Update the profile when email is confirmed
  UPDATE public.profiles
  SET 
    email_verified = true,
    updated_at = NOW()
  WHERE id = new.id;
  
  RETURN new;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Drop existing trigger if exists
DROP TRIGGER IF EXISTS on_auth_user_email_confirmed ON auth.users;

-- Create trigger for email confirmation
CREATE TRIGGER on_auth_user_email_confirmed
  AFTER UPDATE OF email_confirmed_at ON auth.users
  FOR EACH ROW 
  WHEN (old.email_confirmed_at IS NULL AND new.email_confirmed_at IS NOT NULL)
  EXECUTE FUNCTION public.handle_user_email_confirmed();

-- Update existing profiles to set email_verified based on current email_confirmed_at
UPDATE public.profiles p
SET email_verified = (u.email_confirmed_at IS NOT NULL)
FROM auth.users u
WHERE p.id = u.id;