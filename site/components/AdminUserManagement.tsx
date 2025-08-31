'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import { useRouter } from 'next/navigation'

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
  is_admin?: boolean
  beta_access?: boolean
  beta_requested_at?: string
  beta_approved_at?: string
}

export default function AdminUserManagement({ users: initialUsers }: { users: User[] }) {
  const [users, setUsers] = useState(initialUsers)
  const [editingUser, setEditingUser] = useState<string | null>(null)
  const [editForm, setEditForm] = useState<Partial<User>>({})
  const [loading, setLoading] = useState(false)
  const [showAddUser, setShowAddUser] = useState(false)
  const [currentPage, setCurrentPage] = useState(1)
  const usersPerPage = 10
  const [newUser, setNewUser] = useState({
    email: '',
    full_name: '',
    company: '',
    subscription_tier: 'free',
    is_admin: false,
    beta_access: false
  })
  const router = useRouter()
  const supabase = createClient()
  
  // Calculate pagination
  const totalPages = Math.ceil(users.length / usersPerPage)
  const startIndex = (currentPage - 1) * usersPerPage
  const endIndex = startIndex + usersPerPage
  const currentUsers = users.slice(startIndex, endIndex)
  
  const handleSyncProfiles = async () => {
    setLoading(true)
    try {
      const response = await fetch('/api/admin/sync-profiles', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
      })
      
      const data = await response.json()
      
      if (response.ok) {
        alert(`Success: ${data.message}`)
        window.location.reload()
      } else {
        alert(`Error: ${data.error}`)
      }
    } catch (error) {
      console.error('Sync error:', error)
      alert('Failed to sync profiles')
    } finally {
      setLoading(false)
    }
  }

  const handleEdit = (user: User) => {
    setEditingUser(user.id)
    setEditForm({
      full_name: user.full_name || '',
      company: user.company || '',
      subscription_tier: user.subscription_tier,
      email_verified: user.email_verified,
      is_admin: user.is_admin || false,
      beta_access: user.beta_access || false
    })
  }

  const handleSave = async (userId: string) => {
    setLoading(true)
    try {
      const { error } = await supabase
        .from('profiles')
        .update({
          full_name: editForm.full_name,
          company: editForm.company,
          subscription_tier: editForm.subscription_tier,
          email_verified: editForm.email_verified,
          is_admin: editForm.is_admin,
          beta_access: editForm.beta_access,
          beta_approved_at: editForm.beta_access ? new Date().toISOString() : null,
          operations_limit: 
            editForm.subscription_tier === 'beta' ? 1000 :
            editForm.subscription_tier === 'pro' ? 5000 :
            editForm.subscription_tier === 'enterprise' ? 999999 : 25,
          updated_at: new Date().toISOString()
        })
        .eq('id', userId)

      if (error) throw error

      // Update local state
      setUsers(users.map(user => 
        user.id === userId ? { ...user, ...editForm } : user
      ))
      
      setEditingUser(null)
      router.refresh()
    } catch (error) {
      console.error('Error updating user:', error)
      alert('Failed to update user')
    } finally {
      setLoading(false)
    }
  }

  const handleDelete = async (userId: string) => {
    if (!confirm('Are you sure you want to delete this user? This action cannot be undone.')) {
      return
    }

    setLoading(true)
    try {
      const response = await fetch(`/api/admin/delete-user?userId=${userId}`, {
        method: 'DELETE',
        headers: {
          'Content-Type': 'application/json',
        },
      })
      
      const data = await response.json()
      
      if (!response.ok) {
        throw new Error(data.error || 'Failed to delete user')
      }

      // Update local state
      setUsers(users.filter(user => user.id !== userId))
      router.refresh()
    } catch (error) {
      console.error('Error deleting user:', error)
      alert('Failed to delete user: ' + (error as Error).message)
    } finally {
      setLoading(false)
    }
  }

  const handleAddUser = async () => {
    setLoading(true)
    try {
      const response = await fetch('/api/admin/create-user', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(newUser)
      })
      
      const data = await response.json()
      
      if (!response.ok) {
        throw new Error(data.error || 'Failed to create user')
      }

      // Add the new user to the local state immediately
      if (data.user) {
        setUsers([data.user, ...users])
      }

      setShowAddUser(false)
      setNewUser({
        email: '',
        full_name: '',
        company: '',
        subscription_tier: 'free',
        is_admin: false,
        beta_access: false
      })
      router.refresh()
    } catch (error) {
      console.error('Error adding user:', error)
      alert('Failed to add user: ' + (error as Error).message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="space-y-6">
      {/* Add User Button */}
      <div className="flex justify-between items-center">
        <h2 className="text-xl font-semibold text-neutral-900">User Management</h2>
        <div className="flex gap-2">
          <button
            onClick={handleSyncProfiles}
            disabled={loading}
            className="px-4 py-2 border border-amber-300 text-amber-700 rounded-lg hover:bg-amber-50 transition-colors font-medium"
            title="Sync missing profiles from auth users"
          >
            Sync Profiles
          </button>
          <button
            onClick={() => window.location.reload()}
            className="px-4 py-2 border border-neutral-300 text-neutral-700 rounded-lg hover:bg-neutral-50 transition-colors font-medium"
            title="Refresh user list"
          >
            Refresh
          </button>
          <button
            onClick={() => setShowAddUser(!showAddUser)}
            className="px-4 py-2 bg-primary-600 text-white rounded-lg hover:bg-primary-700 transition-colors font-medium"
          >
            {showAddUser ? 'Cancel' : 'Add User'}
          </button>
        </div>
      </div>

      {/* Add User Form */}
      {showAddUser && (
        <div className="glass-card rounded-xl p-6 bg-gradient-to-br from-primary-50 to-blue-50">
          <h3 className="text-lg font-semibold text-neutral-900 mb-4">Add New User</h3>
          <div className="grid md:grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-medium text-neutral-700 mb-1">Email</label>
              <input
                type="email"
                value={newUser.email}
                onChange={(e) => setNewUser({ ...newUser, email: e.target.value })}
                className="w-full px-3 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500"
                required
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-neutral-700 mb-1">Full Name</label>
              <input
                type="text"
                value={newUser.full_name}
                onChange={(e) => setNewUser({ ...newUser, full_name: e.target.value })}
                className="w-full px-3 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-neutral-700 mb-1">Company</label>
              <input
                type="text"
                value={newUser.company}
                onChange={(e) => setNewUser({ ...newUser, company: e.target.value })}
                className="w-full px-3 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-neutral-700 mb-1">Subscription Tier</label>
              <select
                value={newUser.subscription_tier}
                onChange={(e) => setNewUser({ ...newUser, subscription_tier: e.target.value })}
                className="w-full px-3 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500"
              >
                <option value="free">Free</option>
                <option value="beta">Beta</option>
                <option value="pro">Pro</option>
                <option value="enterprise">Enterprise</option>
              </select>
            </div>
            <div className="flex items-center space-x-4">
              <label className="flex items-center">
                <input
                  type="checkbox"
                  checked={newUser.is_admin}
                  onChange={(e) => setNewUser({ ...newUser, is_admin: e.target.checked })}
                  className="mr-2"
                />
                <span className="text-sm font-medium text-neutral-700">Admin Access</span>
              </label>
              <label className="flex items-center">
                <input
                  type="checkbox"
                  checked={newUser.beta_access}
                  onChange={(e) => setNewUser({ ...newUser, beta_access: e.target.checked })}
                  className="mr-2"
                />
                <span className="text-sm font-medium text-neutral-700">Beta Access</span>
              </label>
            </div>
          </div>
          <div className="mt-4 flex justify-end gap-3">
            <button
              onClick={() => setShowAddUser(false)}
              className="px-4 py-2 border border-neutral-300 text-neutral-700 rounded-lg hover:bg-neutral-50"
            >
              Cancel
            </button>
            <button
              onClick={handleAddUser}
              disabled={loading || !newUser.email}
              className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 disabled:opacity-50"
            >
              {loading ? 'Adding...' : 'Add User'}
            </button>
          </div>
        </div>
      )}

      {/* Users Table */}
      <div className="overflow-x-auto">
        <table className="min-w-full">
          <thead>
            <tr className="border-b border-neutral-200">
              <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">User</th>
              <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">Company</th>
              <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">Tier</th>
              <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">Status</th>
              <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">Usage</th>
              <th className="text-right py-3 px-4 text-sm font-medium text-neutral-700">Actions</th>
            </tr>
          </thead>
          <tbody>
            {currentUsers.map(user => (
              <tr key={user.id} className="border-b border-neutral-100 hover:bg-neutral-50">
                <td className="py-3 px-4">
                  <div>
                    <p className="font-medium text-neutral-900">
                      {editingUser === user.id ? (
                        <input
                          type="text"
                          value={editForm.full_name || ''}
                          onChange={(e) => setEditForm({ ...editForm, full_name: e.target.value })}
                          className="px-2 py-1 border rounded"
                        />
                      ) : (
                        user.full_name || 'No name'
                      )}
                    </p>
                    <p className="text-sm text-neutral-500">{user.email}</p>
                  </div>
                </td>
                <td className="py-3 px-4">
                  {editingUser === user.id ? (
                    <input
                      type="text"
                      value={editForm.company || ''}
                      onChange={(e) => setEditForm({ ...editForm, company: e.target.value })}
                      className="px-2 py-1 border rounded w-full"
                    />
                  ) : (
                    <span className="text-sm text-neutral-600">{user.company || '-'}</span>
                  )}
                </td>
                <td className="py-3 px-4">
                  {editingUser === user.id ? (
                    <select
                      value={editForm.subscription_tier}
                      onChange={(e) => setEditForm({ ...editForm, subscription_tier: e.target.value })}
                      className="px-2 py-1 border rounded"
                    >
                      <option value="free">Free</option>
                      <option value="beta">Beta</option>
                      <option value="pro">Pro</option>
                      <option value="enterprise">Enterprise</option>
                    </select>
                  ) : (
                    <span className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${
                      user.subscription_tier === 'beta' ? 'bg-purple-100 text-purple-800' :
                      user.subscription_tier === 'pro' ? 'bg-blue-100 text-blue-800' :
                      user.subscription_tier === 'enterprise' ? 'bg-green-100 text-green-800' :
                      'bg-neutral-100 text-neutral-800'
                    }`}>
                      {user.subscription_tier}
                    </span>
                  )}
                </td>
                <td className="py-3 px-4">
                  <div className="flex items-center gap-2">
                    {editingUser === user.id ? (
                      <>
                        <label className="flex items-center text-xs">
                          <input
                            type="checkbox"
                            checked={editForm.email_verified}
                            onChange={(e) => setEditForm({ ...editForm, email_verified: e.target.checked })}
                            className="mr-1"
                          />
                          Verified
                        </label>
                        <label className="flex items-center text-xs">
                          <input
                            type="checkbox"
                            checked={editForm.is_admin}
                            onChange={(e) => setEditForm({ ...editForm, is_admin: e.target.checked })}
                            className="mr-1"
                          />
                          Admin
                        </label>
                        <label className="flex items-center text-xs">
                          <input
                            type="checkbox"
                            checked={editForm.beta_access}
                            onChange={(e) => setEditForm({ ...editForm, beta_access: e.target.checked })}
                            className="mr-1"
                          />
                          Beta
                        </label>
                      </>
                    ) : (
                      <>
                        {user.email_verified && (
                          <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
                            Verified
                          </span>
                        )}
                        {user.is_admin && (
                          <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-purple-100 text-purple-800">
                            Admin
                          </span>
                        )}
                        {user.beta_access && (
                          <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-800">
                            Beta
                          </span>
                        )}
                      </>
                    )}
                  </div>
                </td>
                <td className="py-3 px-4">
                  <p className="text-sm text-neutral-600">
                    {user.operations_used} / {user.operations_limit}
                  </p>
                </td>
                <td className="py-3 px-4 text-right">
                  {editingUser === user.id ? (
                    <div className="flex justify-end gap-2">
                      <button
                        onClick={() => handleSave(user.id)}
                        disabled={loading}
                        className="px-3 py-1 bg-green-600 text-white rounded text-sm hover:bg-green-700 disabled:opacity-50"
                      >
                        Save
                      </button>
                      <button
                        onClick={() => setEditingUser(null)}
                        className="px-3 py-1 bg-neutral-300 text-neutral-700 rounded text-sm hover:bg-neutral-400"
                      >
                        Cancel
                      </button>
                    </div>
                  ) : (
                    <div className="flex justify-end gap-2">
                      <button
                        onClick={() => handleEdit(user)}
                        className="px-3 py-1 bg-primary-100 text-primary-700 rounded text-sm hover:bg-primary-200"
                      >
                        Edit
                      </button>
                      <button
                        onClick={() => handleDelete(user.id)}
                        className="px-3 py-1 bg-red-100 text-red-700 rounded text-sm hover:bg-red-200"
                      >
                        Delete
                      </button>
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      {/* Pagination Controls */}
      {totalPages > 1 && (
        <div className="flex items-center justify-between mt-4">
          <div className="text-sm text-neutral-600">
            Showing {startIndex + 1} to {Math.min(endIndex, users.length)} of {users.length} users
          </div>
          <div className="flex gap-2">
            <button
              onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
              disabled={currentPage === 1}
              className="px-3 py-1 border border-neutral-300 rounded-lg text-sm hover:bg-neutral-50 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              Previous
            </button>
            
            {/* Page numbers */}
            <div className="flex gap-1">
              {Array.from({ length: totalPages }, (_, i) => i + 1).map(page => {
                // Show first page, last page, current page, and pages around current
                if (
                  page === 1 ||
                  page === totalPages ||
                  (page >= currentPage - 1 && page <= currentPage + 1)
                ) {
                  return (
                    <button
                      key={page}
                      onClick={() => setCurrentPage(page)}
                      className={`px-3 py-1 rounded-lg text-sm ${
                        currentPage === page
                          ? 'bg-primary-600 text-white'
                          : 'border border-neutral-300 hover:bg-neutral-50'
                      }`}
                    >
                      {page}
                    </button>
                  )
                } else if (
                  page === currentPage - 2 ||
                  page === currentPage + 2
                ) {
                  return <span key={page} className="px-1 text-neutral-400">...</span>
                }
                return null
              })}
            </div>
            
            <button
              onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
              disabled={currentPage === totalPages}
              className="px-3 py-1 border border-neutral-300 rounded-lg text-sm hover:bg-neutral-50 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              Next
            </button>
          </div>
        </div>
      )}
    </div>
  )
}