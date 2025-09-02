'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import { User } from '@supabase/supabase-js'

interface ProfileFormProps {
  user: User
  profile: any
}

export default function ProfileForm({ user, profile }: ProfileFormProps) {
  const [loading, setLoading] = useState(false)
  const [success, setSuccess] = useState(false)
  const [formData, setFormData] = useState({
    firstName: profile?.first_name || profile?.full_name?.split(' ')[0] || '',
    surname: profile?.surname || profile?.full_name?.split(' ').slice(1).join(' ') || '',
    company: profile?.company || '',
    newsletter: profile?.newsletter_subscribed ?? true,
  })

  const supabase = createClient()

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setLoading(true)
    setSuccess(false)

    const fullName = formData.surname 
      ? `${formData.firstName} ${formData.surname}` 
      : formData.firstName

    const { error } = await supabase
      .from('profiles')
      .update({
        full_name: fullName,
        first_name: formData.firstName || null,
        surname: formData.surname || null,
        company: formData.company || null,
        newsletter_subscribed: formData.newsletter,
        updated_at: new Date().toISOString(),
      })
      .eq('id', user.id)

    if (!error) {
      setSuccess(true)
      setTimeout(() => setSuccess(false), 3000)
    }

    setLoading(false)
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-4">
      <div className="grid grid-cols-2 gap-4">
        <div>
          <label htmlFor="firstName" className="block text-sm font-medium text-neutral-700 mb-1">
            First Name
          </label>
          <input
            id="firstName"
            type="text"
            value={formData.firstName}
            onChange={(e) => setFormData({ ...formData, firstName: e.target.value })}
            className="w-full px-4 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all"
            placeholder="John"
          />
        </div>
        <div>
          <label htmlFor="surname" className="block text-sm font-medium text-neutral-700 mb-1">
            Surname
          </label>
          <input
            id="surname"
            type="text"
            value={formData.surname}
            onChange={(e) => setFormData({ ...formData, surname: e.target.value })}
            className="w-full px-4 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all"
            placeholder="Doe"
          />
        </div>
      </div>

      <div>
        <label htmlFor="company" className="block text-sm font-medium text-neutral-700 mb-1">
          Company
        </label>
        <input
          id="company"
          type="text"
          value={formData.company}
          onChange={(e) => setFormData({ ...formData, company: e.target.value })}
          className="w-full px-4 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all"
          placeholder="Acme Inc."
        />
      </div>

      <div className="pt-4 border-t border-neutral-200">
        <h3 className="text-sm font-medium text-neutral-700 mb-3">Email Preferences</h3>
        <div className="flex items-start">
          <input
            id="newsletter"
            type="checkbox"
            checked={formData.newsletter}
            onChange={(e) => setFormData({ ...formData, newsletter: e.target.checked })}
            className="mt-1 h-4 w-4 text-primary-600 border-neutral-300 rounded focus:ring-primary-500"
          />
          <label htmlFor="newsletter" className="ml-2 text-sm text-neutral-600">
            Receive product updates and newsletter emails
          </label>
        </div>
      </div>

      <div className="flex items-center justify-between pt-4">
        <button
          type="submit"
          disabled={loading}
          className="px-6 py-2 bg-gradient-to-r from-primary-600 to-primary-700 text-white font-medium rounded-lg hover:from-primary-700 hover:to-primary-800 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
        >
          {loading ? 'Saving...' : 'Save Changes'}
        </button>
        
        {success && (
          <span className="text-sm text-green-600 font-medium">
            ✓ Profile updated successfully
          </span>
        )}
      </div>
    </form>
  )
}