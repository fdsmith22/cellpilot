import { redirect } from 'next/navigation'
import { createClient } from '@/lib/supabase/server'
import { createClient as createServiceClient } from '@supabase/supabase-js'
import AdminUserManagement from '@/components/AdminUserManagement'
import GridAnimation from '@/components/GridAnimation'

export default async function AdminPage() {
  const supabase = await createClient()
  
  const { data: { user } } = await supabase.auth.getUser()

  if (!user) {
    redirect('/auth')
  }

  // Check if user is admin
  const { data: profile } = await supabase
    .from('profiles')
    .select('is_admin')
    .eq('id', user.id)
    .single()

  if (!profile?.is_admin) {
    redirect('/dashboard')
  }

  // Create service client to bypass RLS and get ALL users
  const serviceClient = createServiceClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.SUPABASE_SERVICE_ROLE_KEY!,
    {
      auth: {
        autoRefreshToken: false,
        persistSession: false
      }
    }
  )

  // Get all users from database without any RLS restrictions
  const { data: users } = await serviceClient
    .from('profiles')
    .select('*')
    .order('created_at', { ascending: false })

  console.log('Fetched ALL users from database:', users?.length)
  
  // Get tier counts efficiently (optional - for future optimization)
  // This could replace the filter counts if you have many users
  const tierCounts = {
    free: users?.filter(u => u.subscription_tier === 'free').length || 0,
    beta: users?.filter(u => u.subscription_tier === 'beta').length || 0,
    pro: users?.filter(u => u.subscription_tier === 'pro').length || 0,
    enterprise: users?.filter(u => u.subscription_tier === 'enterprise').length || 0,
    total: users?.length || 0
  }
  
  console.log('User tier distribution:', tierCounts)

  return (
    <>
      <GridAnimation />
      
      <div className="min-h-screen bg-gradient-to-b from-white/60 via-white/70 to-white/80">
        <div className="container-wrapper py-12">
          {/* Admin Header */}
          <div className="glass-card rounded-2xl p-8 mb-8">
            <div className="flex justify-between items-center">
              <div>
                <h1 className="text-3xl font-bold text-neutral-900 mb-2">Admin Dashboard</h1>
                <p className="text-neutral-600">Manage users and subscriptions</p>
              </div>
              <div className="flex gap-4">
                <div className="text-center px-4">
                  <div className="text-3xl font-bold text-primary-600">{tierCounts.total}</div>
                  <div className="text-sm text-neutral-600">Total Users</div>
                </div>
                <div className="h-12 w-px bg-neutral-200"></div>
                <div className="text-center px-3">
                  <div className="text-2xl font-bold text-neutral-500">
                    {tierCounts.free}
                  </div>
                  <div className="text-sm text-neutral-600">Free</div>
                </div>
                <div className="text-center px-3">
                  <div className="text-2xl font-bold text-purple-600">
                    {tierCounts.beta}
                  </div>
                  <div className="text-sm text-neutral-600">Beta</div>
                </div>
                <div className="text-center px-3">
                  <div className="text-2xl font-bold text-blue-600">
                    {tierCounts.pro}
                  </div>
                  <div className="text-sm text-neutral-600">Pro</div>
                </div>
                <div className="text-center px-3">
                  <div className="text-2xl font-bold text-green-600">
                    {tierCounts.enterprise}
                  </div>
                  <div className="text-sm text-neutral-600">Enterprise</div>
                </div>
              </div>
            </div>
          </div>

          {/* User Management Table */}
          <div className="glass-card rounded-2xl p-8">
            <AdminUserManagement users={users || []} />
          </div>

          {/* Quick Actions */}
          <div className="mt-8 grid md:grid-cols-3 gap-6">
            <div className="glass-card rounded-xl p-6">
              <h3 className="font-semibold text-neutral-900 mb-2">Subscription Tiers</h3>
              <ul className="space-y-2 text-sm text-neutral-600">
                <li className="flex justify-between">
                  <span><span className="font-medium">Free:</span> 25 ops/month</span>
                  <span className="text-neutral-500">{tierCounts.free} users</span>
                </li>
                <li className="flex justify-between">
                  <span><span className="font-medium text-purple-600">Beta:</span> 1,000 ops/month</span>
                  <span className="text-purple-600">{tierCounts.beta} users</span>
                </li>
                <li className="flex justify-between">
                  <span><span className="font-medium text-blue-600">Pro:</span> 5,000 ops/month</span>
                  <span className="text-blue-600">{tierCounts.pro} users</span>
                </li>
                <li className="flex justify-between">
                  <span><span className="font-medium text-green-600">Enterprise:</span> Unlimited</span>
                  <span className="text-green-600">{tierCounts.enterprise} users</span>
                </li>
              </ul>
            </div>

            <div className="glass-card rounded-xl p-6">
              <h3 className="font-semibold text-neutral-900 mb-2">Admin Actions</h3>
              <ul className="space-y-2 text-sm text-neutral-600">
                <li>• Click user row to edit</li>
                <li>• Change subscription tiers</li>
                <li>• View user details</li>
                <li>• Monitor usage statistics</li>
              </ul>
            </div>

            <div className="glass-card rounded-xl p-6">
              <h3 className="font-semibold text-neutral-900 mb-2">Statistics</h3>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between">
                  <span className="text-neutral-600">Verified Emails:</span>
                  <span className="font-medium">
                    {users?.filter(u => u.email_verified).length || 0}
                  </span>
                </div>
                <div className="flex justify-between">
                  <span className="text-neutral-600">Newsletter Subs:</span>
                  <span className="font-medium">
                    {users?.filter(u => u.newsletter_subscribed).length || 0}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </>
  )
}