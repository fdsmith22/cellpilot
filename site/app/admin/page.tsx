import { redirect } from 'next/navigation'
import { createClient } from '@/lib/supabase/server'
import AdminUserTable from '@/components/AdminUserTable'
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

  // Get all users
  const { data: users } = await supabase
    .from('profiles')
    .select('*')
    .order('created_at', { ascending: false })

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
                <div className="text-center">
                  <div className="text-2xl font-bold text-primary-600">{users?.length || 0}</div>
                  <div className="text-sm text-neutral-600">Total Users</div>
                </div>
                <div className="text-center">
                  <div className="text-2xl font-bold text-green-600">
                    {users?.filter(u => u.subscription_tier === 'beta').length || 0}
                  </div>
                  <div className="text-sm text-neutral-600">Beta Users</div>
                </div>
              </div>
            </div>
          </div>

          {/* User Management Table */}
          <div className="glass-card rounded-2xl p-8">
            <h2 className="text-xl font-semibold text-neutral-900 mb-6">User Management</h2>
            <AdminUserTable users={users || []} />
          </div>

          {/* Quick Actions */}
          <div className="mt-8 grid md:grid-cols-3 gap-6">
            <div className="glass-card rounded-xl p-6">
              <h3 className="font-semibold text-neutral-900 mb-2">Subscription Tiers</h3>
              <ul className="space-y-2 text-sm text-neutral-600">
                <li><span className="font-medium">Free:</span> 25 operations/month</li>
                <li><span className="font-medium">Beta:</span> 1,000 operations/month (Early access)</li>
                <li><span className="font-medium">Pro:</span> 5,000 operations/month (Coming soon)</li>
                <li><span className="font-medium">Enterprise:</span> Unlimited (Coming soon)</li>
              </ul>
              <p className="text-xs text-neutral-500 mt-3 pt-3 border-t">
                Note: Beta is a subscription tier, not admin access
              </p>
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