-- Create a function to delete user account
-- This needs to be run with service_role permissions

CREATE OR REPLACE FUNCTION public.delete_user()
RETURNS void AS $$
DECLARE
  user_id uuid;
BEGIN
  -- Get the current user's ID
  user_id := auth.uid();
  
  IF user_id IS NULL THEN
    RAISE EXCEPTION 'Not authenticated';
  END IF;

  -- Delete from related tables first (due to foreign key constraints)
  DELETE FROM public.installations WHERE user_id = user_id;
  DELETE FROM public.profiles WHERE id = user_id;
  
  -- Delete the user from auth.users (this requires elevated permissions)
  -- For security, this should be called through an edge function with service_role key
  -- For now, we'll just delete the profile data
  
  -- Note: Complete user deletion from auth.users table requires admin/service_role access
  -- You'll need to create a Supabase Edge Function for complete deletion
  
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Grant execute permission to authenticated users
GRANT EXECUTE ON FUNCTION public.delete_user() TO authenticated;

-- Alternative: Create an edge function for complete deletion
-- This would be a separate TypeScript file deployed as a Supabase Edge Function:
/*
// supabase/functions/delete-user/index.ts
import { serve } from 'https://deno.land/std@0.168.0/http/server.ts'
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

serve(async (req) => {
  const supabaseClient = createClient(
    Deno.env.get('SUPABASE_URL') ?? '',
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? ''
  )

  const token = req.headers.get('Authorization')?.replace('Bearer ', '')
  
  if (!token) {
    return new Response('Unauthorized', { status: 401 })
  }

  const { data: { user } } = await supabaseClient.auth.getUser(token)
  
  if (!user) {
    return new Response('Unauthorized', { status: 401 })
  }

  // Delete user data
  await supabaseClient.from('profiles').delete().eq('id', user.id)
  await supabaseClient.from('installations').delete().eq('user_id', user.id)
  
  // Delete the user account
  const { error } = await supabaseClient.auth.admin.deleteUser(user.id)
  
  if (error) {
    return new Response(JSON.stringify({ error: error.message }), { 
      status: 400,
      headers: { 'Content-Type': 'application/json' }
    })
  }

  return new Response(JSON.stringify({ success: true }), {
    headers: { 'Content-Type': 'application/json' }
  })
})
*/