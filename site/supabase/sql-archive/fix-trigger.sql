-- Drop the existing trigger and function
DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
DROP FUNCTION IF EXISTS public.handle_new_user();

-- Create an improved trigger function that handles errors better
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO public.profiles (id, email, email_verified)
  VALUES (
    new.id, 
    new.email, 
    COALESCE(new.email_confirmed_at IS NOT NULL, false)
  )
  ON CONFLICT (id) DO UPDATE
  SET 
    email = EXCLUDED.email,
    updated_at = NOW();
  
  RETURN new;
EXCEPTION
  WHEN others THEN
    -- Log the error but don't prevent user creation
    RAISE WARNING 'Error creating profile for user %: %', new.id, SQLERRM;
    RETURN new;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Create the trigger
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- Also ensure the profiles table has proper constraints
ALTER TABLE profiles 
  DROP CONSTRAINT IF EXISTS profiles_id_fkey,
  ADD CONSTRAINT profiles_id_fkey 
    FOREIGN KEY (id) 
    REFERENCES auth.users(id) 
    ON DELETE CASCADE;

-- Add an index for better performance
CREATE INDEX IF NOT EXISTS profiles_email_idx ON profiles(email);