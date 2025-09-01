-- Create a table for user profiles and metadata
CREATE TABLE IF NOT EXISTS profiles (
  id UUID REFERENCES auth.users ON DELETE CASCADE,
  email TEXT,
  full_name TEXT,
  company TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()),
  subscription_tier TEXT DEFAULT 'free',
  operations_used INTEGER DEFAULT 0,
  operations_limit INTEGER DEFAULT 25,
  last_login TIMESTAMP WITH TIME ZONE,
  email_verified BOOLEAN DEFAULT false,
  newsletter_subscribed BOOLEAN DEFAULT true,
  PRIMARY KEY (id)
);

-- Enable Row Level Security
ALTER TABLE profiles ENABLE ROW LEVEL SECURITY;

-- Create policy for users to see their own profile
CREATE POLICY "Users can view own profile" ON profiles
  FOR SELECT USING (auth.uid() = id);

-- Create policy for users to update their own profile  
CREATE POLICY "Users can update own profile" ON profiles
  FOR UPDATE USING (auth.uid() = id);

-- Create trigger to create profile on signup
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO public.profiles (id, email, email_verified)
  VALUES (new.id, new.email, new.email_verified);
  RETURN new;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Drop existing trigger if exists
DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;

-- Create trigger
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- Create a table for tracking installations
CREATE TABLE IF NOT EXISTS installations (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_id UUID REFERENCES auth.users ON DELETE CASCADE,
  installed_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()),
  sheet_id TEXT,
  sheet_name TEXT,
  installation_method TEXT,
  is_active BOOLEAN DEFAULT true
);

-- Enable RLS for installations
ALTER TABLE installations ENABLE ROW LEVEL SECURITY;

-- Policy for users to see their own installations
CREATE POLICY "Users can view own installations" ON installations
  FOR SELECT USING (auth.uid() = user_id);

-- Policy for users to create their own installations
CREATE POLICY "Users can create own installations" ON installations
  FOR INSERT WITH CHECK (auth.uid() = user_id);

-- Create table for email marketing preferences
CREATE TABLE IF NOT EXISTS email_preferences (
  id UUID REFERENCES auth.users ON DELETE CASCADE,
  marketing_emails BOOLEAN DEFAULT true,
  product_updates BOOLEAN DEFAULT true,
  tips_and_tricks BOOLEAN DEFAULT true,
  survey_invitations BOOLEAN DEFAULT false,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()),
  PRIMARY KEY (id)
);

-- Enable RLS
ALTER TABLE email_preferences ENABLE ROW LEVEL SECURITY;

-- Policies for email preferences
CREATE POLICY "Users can manage own email preferences" ON email_preferences
  FOR ALL USING (auth.uid() = id);