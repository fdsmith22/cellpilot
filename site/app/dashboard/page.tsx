import { redirect } from 'next/navigation'
import { createClient } from '@/lib/supabase/server'
import Link from 'next/link'
import GridAnimation from '@/components/GridAnimation'
import SignOutButton from '@/components/SignOutButton'
import ProfileForm from '@/components/ProfileForm'
import DangerZone from '@/components/DangerZone'
import BetaAccessCard from '@/components/BetaAccessCard'

export default async function DashboardPage() {
  const supabase = await createClient()
  
  const { data: { user } } = await supabase.auth.getUser()

  if (!user) {
    redirect('/auth')
  }

  // Get user profile data
  const { data: profile } = await supabase
    .from('profiles')
    .select('*')
    .eq('id', user.id)
    .single()

  return (
    <>
      <GridAnimation />
      
      <div className="min-h-screen bg-gradient-to-b from-white/60 via-white/70 to-white/80">
        <div className="container-wrapper py-12">
          {/* Dashboard Header */}
          <div className="glass-card rounded-2xl p-8 mb-8">
            <div className="flex justify-between items-start">
              <div>
                <h1 className="text-3xl font-bold text-neutral-900 mb-2">
                  Welcome back{profile?.full_name ? `, ${profile.full_name.split(' ')[0]}` : ''}!
                </h1>
                <p className="text-neutral-600">{user.email}</p>
                {profile?.company && (
                  <p className="text-sm text-neutral-500 mt-1">{profile.company}</p>
                )}
              </div>
              <SignOutButton />
            </div>
          </div>

          {/* Admin Link if user is admin */}
          {profile?.is_admin && (
            <div className="mb-6">
              <Link 
                href="/admin" 
                className="inline-flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-purple-600 to-purple-700 text-white rounded-lg hover:from-purple-700 hover:to-purple-800 font-medium text-sm transition-all"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4" />
                </svg>
                Admin Dashboard
              </Link>
            </div>
          )}

          {/* Quick Actions */}
          <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-6 mb-8">
            <Link href="/install" className="glass-card rounded-xl p-6 hover:shadow-xl transition-all hover:scale-102">
              <div className="w-12 h-12 rounded-lg bg-gradient-to-br from-primary-100 to-primary-50 flex items-center justify-center mb-4">
                <svg className="w-6 h-6 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
              </div>
              <h3 className="text-lg font-semibold text-neutral-900 mb-2">Install CellPilot</h3>
              <p className="text-sm text-neutral-600">Get started with CellPilot for Google Sheets</p>
            </Link>

            <Link href="/docs" className="glass-card rounded-xl p-6 hover:shadow-xl transition-all hover:scale-102">
              <div className="w-12 h-12 rounded-lg bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 flex items-center justify-center mb-4">
                <svg className="w-6 h-6 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253" />
                </svg>
              </div>
              <h3 className="text-lg font-semibold text-neutral-900 mb-2">Documentation</h3>
              <p className="text-sm text-neutral-600">Learn how to use CellPilot effectively</p>
            </Link>

            <Link href="/support" className="glass-card rounded-xl p-6 hover:shadow-xl transition-all hover:scale-102">
              <div className="w-12 h-12 rounded-lg bg-gradient-to-br from-pastel-lavender/30 to-pastel-lavender/20 flex items-center justify-center mb-4">
                <svg className="w-6 h-6 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M18.364 5.636l-3.536 3.536m0 5.656l3.536 3.536M9.172 9.172L5.636 5.636m3.536 9.192l-3.536 3.536M21 12a9 9 0 11-18 0 9 9 0 0118 0zm-5 0a4 4 0 11-8 0 4 4 0 018 0z" />
                </svg>
              </div>
              <h3 className="text-lg font-semibold text-neutral-900 mb-2">Get Support</h3>
              <p className="text-sm text-neutral-600">Need help? Contact our support team</p>
            </Link>
          </div>

          {/* Beta Access Card - Show for users who haven't requested or have requested/been approved */}
          {!profile?.beta_access && (
            <div className="mb-8">
              <BetaAccessCard profile={profile} userId={user.id} />
            </div>
          )}
          
          {/* Show success card for approved users */}
          {profile?.beta_access && (
            <div className="mb-8">
              <BetaAccessCard profile={profile} userId={user.id} />
            </div>
          )}

          <div className="grid lg:grid-cols-2 gap-8">
            {/* Profile Information */}
            <div className="glass-card rounded-2xl p-8">
              <h2 className="text-xl font-semibold text-neutral-900 mb-6">Profile Information</h2>
              <ProfileForm user={user} profile={profile} />
            </div>

            {/* Account Details */}
            <div className="glass-card rounded-2xl p-8">
              <h2 className="text-xl font-semibold text-neutral-900 mb-6">Account Details</h2>
              
              <div className="space-y-4">
                <div className="flex justify-between items-center py-3 border-b border-neutral-200">
                  <span className="text-neutral-600">Email</span>
                  <div className="flex items-center gap-2">
                    <span className="text-neutral-900 font-medium">{user.email}</span>
                    {profile?.email_verified ? (
                      <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
                        Verified
                      </span>
                    ) : (
                      <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-yellow-100 text-yellow-800">
                        Unverified
                      </span>
                    )}
                  </div>
                </div>
                
                <div className="flex justify-between items-center py-3 border-b border-neutral-200">
                  <span className="text-neutral-600">User ID</span>
                  <span className="text-neutral-900 font-mono text-sm">{user.id}</span>
                </div>
                
                <div className="flex justify-between items-center py-3 border-b border-neutral-200">
                  <span className="text-neutral-600">Account Created</span>
                  <span className="text-neutral-900">{new Date(user.created_at).toLocaleDateString()}</span>
                </div>
                
                <div className="flex justify-between items-center py-3 border-b border-neutral-200">
                  <span className="text-neutral-600">Current Plan</span>
                  <span className="text-neutral-900 font-medium">{profile?.subscription_tier || 'Free'}</span>
                </div>

                <div className="flex justify-between items-center py-3">
                  <span className="text-neutral-600">Operations Used</span>
                  <span className="text-neutral-900 font-medium">
                    {profile?.operations_used || 0} / {profile?.operations_limit || 25}
                  </span>
                </div>
              </div>

              <div className="mt-8 pt-6 border-t border-neutral-200">
                <h3 className="text-lg font-semibold text-neutral-900 mb-4">Subscription</h3>
                <div className="bg-gradient-to-r from-pastel-mint/20 to-pastel-sky/20 rounded-lg p-4 mb-4">
                  <div className="flex justify-between items-center">
                    <div>
                      <p className="font-medium text-neutral-900">Free Plan</p>
                      <p className="text-sm text-neutral-600">25 operations per month</p>
                    </div>
                    <Link href="/pricing" className="btn-primary px-4 py-2 text-sm">
                      Upgrade
                    </Link>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Danger Zone - Smaller and subtle */}
          <div className="glass-card rounded-xl p-4 border border-red-100 bg-red-50/10">
            <DangerZone userEmail={user.email || ''} />
          </div>
        </div>
      </div>
    </>
  )
}