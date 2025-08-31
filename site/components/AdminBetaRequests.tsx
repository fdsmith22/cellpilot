'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import { useRouter } from 'next/navigation'

interface BetaRequest {
  id: string
  email: string
  full_name: string | null
  beta_requested_at: string
  beta_approved_at: string | null
  beta_access: boolean
  beta_notes: string | null
  created_at: string
}

interface AdminBetaRequestsProps {
  requests: BetaRequest[]
}

export default function AdminBetaRequests({ requests: initialRequests }: AdminBetaRequestsProps) {
  const [requests, setRequests] = useState(initialRequests)
  const [loading, setLoading] = useState<string | null>(null)
  const router = useRouter()
  const supabase = createClient()

  const approveBetaAccess = async (userId: string) => {
    setLoading(userId)
    try {
      const { error } = await supabase
        .from('profiles')
        .update({ 
          beta_access: true,
          beta_approved_at: new Date().toISOString()
        })
        .eq('id', userId)

      if (error) throw error
      
      // Update local state
      setRequests(prev => prev.map(req => 
        req.id === userId 
          ? { ...req, beta_access: true, beta_approved_at: new Date().toISOString() }
          : req
      ))
      
      router.refresh()
    } catch (error) {
      console.error('Error approving beta access:', error)
      alert('Failed to approve beta access')
    } finally {
      setLoading(null)
    }
  }

  const revokeBetaAccess = async (userId: string) => {
    setLoading(userId)
    try {
      const { error } = await supabase
        .from('profiles')
        .update({ 
          beta_access: false,
          beta_approved_at: null
        })
        .eq('id', userId)

      if (error) throw error
      
      // Update local state
      setRequests(prev => prev.map(req => 
        req.id === userId 
          ? { ...req, beta_access: false, beta_approved_at: null }
          : req
      ))
      
      router.refresh()
    } catch (error) {
      console.error('Error revoking beta access:', error)
      alert('Failed to revoke beta access')
    } finally {
      setLoading(null)
    }
  }

  const pendingRequests = requests.filter(r => r.beta_requested_at && !r.beta_access)
  const approvedRequests = requests.filter(r => r.beta_access)

  return (
    <div className="space-y-8">
      {/* Pending Requests */}
      {pendingRequests.length > 0 && (
        <div className="glass-card rounded-2xl p-8">
          <h2 className="text-xl font-semibold text-amber-900 mb-6 flex items-center">
            <span className="w-2 h-2 bg-amber-500 rounded-full mr-2 animate-pulse"></span>
            Pending Beta Requests ({pendingRequests.length})
          </h2>
          
          <div className="overflow-x-auto">
            <table className="min-w-full">
              <thead>
                <tr className="border-b border-neutral-200">
                  <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">User</th>
                  <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">Requested</th>
                  <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">Notes</th>
                  <th className="text-right py-3 px-4 text-sm font-medium text-neutral-700">Action</th>
                </tr>
              </thead>
              <tbody>
                {pendingRequests.map(request => (
                  <tr key={request.id} className="border-b border-neutral-100 hover:bg-neutral-50">
                    <td className="py-3 px-4">
                      <div>
                        <p className="font-medium text-neutral-900">{request.email}</p>
                        {request.full_name && (
                          <p className="text-sm text-neutral-500">{request.full_name}</p>
                        )}
                      </div>
                    </td>
                    <td className="py-3 px-4">
                      <p className="text-sm text-neutral-600">
                        {new Date(request.beta_requested_at).toLocaleDateString()}
                      </p>
                      <p className="text-xs text-neutral-500">
                        {new Date(request.beta_requested_at).toLocaleTimeString()}
                      </p>
                    </td>
                    <td className="py-3 px-4">
                      <p className="text-sm text-neutral-600">{request.beta_notes || '-'}</p>
                    </td>
                    <td className="py-3 px-4 text-right">
                      <button
                        onClick={() => approveBetaAccess(request.id)}
                        disabled={loading === request.id}
                        className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors text-sm font-medium disabled:opacity-50"
                      >
                        {loading === request.id ? 'Approving...' : 'Approve'}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Approved Users */}
      {approvedRequests.length > 0 && (
        <div className="glass-card rounded-2xl p-8">
          <h2 className="text-xl font-semibold text-green-900 mb-6 flex items-center">
            <span className="w-2 h-2 bg-green-500 rounded-full mr-2"></span>
            Approved Beta Users ({approvedRequests.length})
          </h2>
          
          <div className="overflow-x-auto">
            <table className="min-w-full">
              <thead>
                <tr className="border-b border-neutral-200">
                  <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">User</th>
                  <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">Approved</th>
                  <th className="text-left py-3 px-4 text-sm font-medium text-neutral-700">Status</th>
                  <th className="text-right py-3 px-4 text-sm font-medium text-neutral-700">Action</th>
                </tr>
              </thead>
              <tbody>
                {approvedRequests.map(request => (
                  <tr key={request.id} className="border-b border-neutral-100 hover:bg-neutral-50">
                    <td className="py-3 px-4">
                      <div>
                        <p className="font-medium text-neutral-900">{request.email}</p>
                        {request.full_name && (
                          <p className="text-sm text-neutral-500">{request.full_name}</p>
                        )}
                      </div>
                    </td>
                    <td className="py-3 px-4">
                      <p className="text-sm text-neutral-600">
                        {request.beta_approved_at 
                          ? new Date(request.beta_approved_at).toLocaleDateString()
                          : 'Legacy access'}
                      </p>
                    </td>
                    <td className="py-3 px-4">
                      <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
                        Active
                      </span>
                    </td>
                    <td className="py-3 px-4 text-right">
                      <button
                        onClick={() => revokeBetaAccess(request.id)}
                        disabled={loading === request.id}
                        className="px-4 py-2 bg-red-100 text-red-700 rounded-lg hover:bg-red-200 transition-colors text-sm font-medium disabled:opacity-50"
                      >
                        {loading === request.id ? 'Revoking...' : 'Revoke'}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* No Requests */}
      {requests.length === 0 && (
        <div className="glass-card rounded-2xl p-8 text-center">
          <p className="text-neutral-600">No beta access requests yet</p>
        </div>
      )}
    </div>
  )
}