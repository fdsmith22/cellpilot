-- Create or replace the RLS policy to allow users to update their own subscription_tier to 'beta'
CREATE POLICY "Users can set their own beta tier" ON profiles
FOR UPDATE 
TO authenticated
USING (auth.uid() = id)
WITH CHECK (
  auth.uid() = id 
  AND (
    -- Allow updating to beta tier during beta phase
    subscription_tier = 'beta' 
    -- Or keeping existing tier
    OR subscription_tier = profiles.subscription_tier
  )
);

-- Note: This policy allows users to self-select into beta tier
-- Once beta phase ends, this policy should be updated to remove beta option