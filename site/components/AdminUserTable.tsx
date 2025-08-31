'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabase/client'

interface User {
  id: string
  email: string
  full_name: string | null
  company: string | null
  subscription_tier: string
  operations_used: number
  operations_limit: number
  email_verified: boolean
  newsletter_subscribed: boolean
  created_at: string
  last_login: string | null
  is_admin?: boolean
}

export default function AdminUserTable({ users: initialUsers }: { users: User[] }) {
  const [users, setUsers] = useState(initialUsers)
  const [editingUser, setEditingUser] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)
  const supabase = createClient()

  const handleTierChange = async (userId: string, newTier: string) => {
    setLoading(true)
    
    try {
      // Call the admin function to update tier
      const { error } = await supabase.rpc('admin_update_user_tier', {
        target_user_id: userId,
        new_tier: newTier
      })

      if (error) throw error

      // Update local state
      setUsers(users.map(user => 
        user.id === userId 
          ? { 
              ...user, 
              subscription_tier: newTier,
              operations_limit: newTier === 'beta' ? 1000 : newTier === 'pro' ? 5000 : 25
            }
          : user
      ))
      
      setEditingUser(null)
    } catch (error) {
      console.error('Error updating tier:', error)
      alert('Failed to update user tier')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-neutral-200">
            <th className="text-left py-3 px-2 font-medium text-neutral-700">User</th>
            <th className="text-left py-3 px-2 font-medium text-neutral-700">Company</th>
            <th className="text-left py-3 px-2 font-medium text-neutral-700">Plan</th>
            <th className="text-left py-3 px-2 font-medium text-neutral-700">Usage</th>
            <th className="text-left py-3 px-2 font-medium text-neutral-700">Status</th>
            <th className="text-left py-3 px-2 font-medium text-neutral-700">Joined</th>
          </tr>
        </thead>
        <tbody>
          {users.map((user) => (
            <tr key={user.id} className="border-b border-neutral-100 hover:bg-neutral-50">
              <td className="py-3 px-2">
                <div>
                  <div className="font-medium text-neutral-900">
                    {user.full_name || 'No name'}
                    {user.is_admin && (
                      <span className="ml-2 text-xs bg-primary-100 text-primary-700 px-2 py-0.5 rounded">
                        Admin
                      </span>
                    )}
                  </div>
                  <div className="text-xs text-neutral-500">{user.email}</div>
                </div>
              </td>
              <td className="py-3 px-2 text-neutral-600">
                {user.company || '-'}
              </td>
              <td className="py-3 px-2">
                {editingUser === user.id ? (
                  <select
                    value={user.subscription_tier}
                    onChange={(e) => handleTierChange(user.id, e.target.value)}
                    onBlur={() => setEditingUser(null)}
                    disabled={loading}
                    className="px-2 py-1 border border-neutral-300 rounded text-sm focus:outline-none focus:ring-2 focus:ring-primary-500"
                    autoFocus
                  >
                    <option value="free">Free</option>
                    <option value="beta">Beta</option>
                    <option value="pro">Pro</option>
                    <option value="enterprise">Enterprise</option>
                  </select>
                ) : (
                  <button
                    onClick={() => setEditingUser(user.id)}
                    className={`px-3 py-1 rounded text-xs font-medium ${
                      user.subscription_tier === 'beta' 
                        ? 'bg-green-100 text-green-700'
                        : user.subscription_tier === 'pro'
                        ? 'bg-primary-100 text-primary-700'
                        : user.subscription_tier === 'enterprise'
                        ? 'bg-purple-100 text-purple-700'
                        : 'bg-neutral-100 text-neutral-700'
                    } hover:opacity-80 transition-opacity`}
                  >
                    {user.subscription_tier}
                  </button>
                )}
              </td>
              <td className="py-3 px-2">
                <div className="text-xs">
                  <div>{user.operations_used} / {user.operations_limit}</div>
                  <div className="w-24 bg-neutral-200 rounded-full h-1.5 mt-1">
                    <div 
                      className="bg-primary-600 h-1.5 rounded-full"
                      style={{ 
                        width: `${Math.min((user.operations_used / user.operations_limit) * 100, 100)}%` 
                      }}
                    />
                  </div>
                </div>
              </td>
              <td className="py-3 px-2">
                <div className="flex gap-1">
                  {user.email_verified && (
                    <span className="text-xs bg-green-100 text-green-700 px-2 py-0.5 rounded">
                      Verified
                    </span>
                  )}
                  {user.newsletter_subscribed && (
                    <span className="text-xs bg-blue-100 text-blue-700 px-2 py-0.5 rounded">
                      Newsletter
                    </span>
                  )}
                </div>
              </td>
              <td className="py-3 px-2 text-neutral-600 text-xs">
                {new Date(user.created_at).toLocaleDateString()}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}