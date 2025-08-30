'use client'

import { useState, useEffect } from 'react'
import { useAuth } from '@/components/AuthProvider'
import { useRouter, useSearchParams } from 'next/navigation'
import Link from 'next/link'
import GridAnimation from '@/components/GridAnimation'
import Image from 'next/image'

export default function AuthPage() {
  const searchParams = useSearchParams()
  const [isSignUp, setIsSignUp] = useState(searchParams.get('signup') === 'true')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [firstName, setFirstName] = useState('')
  const [surname, setSurname] = useState('')
  const [company, setCompany] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [newsletter, setNewsletter] = useState(true)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [success, setSuccess] = useState<string | null>(null)
  
  const { signIn, signUp, signInWithEmail } = useAuth()
  const router = useRouter()

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setLoading(true)
    setError(null)
    setSuccess(null)

    if (isSignUp) {
      const fullName = surname ? `${firstName} ${surname}` : firstName
      const metadata = {
        full_name: fullName,
        first_name: firstName,
        surname: surname || null,
        company: company || null,
        newsletter_subscribed: newsletter
      }
      
      const { error } = await signUp(email, password, metadata)
      if (error) {
        setError(error.message)
      } else {
        setSuccess('Check your email to confirm your account!')
      }
    } else {
      const { error } = await signIn(email, password)
      if (error) {
        setError(error.message)
      } else {
        router.push('/dashboard')
      }
    }
    
    setLoading(false)
  }

  const handleMagicLink = async () => {
    setLoading(true)
    setError(null)
    setSuccess(null)

    const { error } = await signInWithEmail(email)
    if (error) {
      setError(error.message)
    } else {
      setSuccess('Check your email for the login link!')
    }
    
    setLoading(false)
  }

  return (
    <>
      <GridAnimation />
      
      <div className="min-h-screen bg-gradient-to-b from-white/60 via-white/70 to-white/80 flex items-center justify-center px-4 py-12">
        <div className="w-full max-w-md">
          <div className="glass-card rounded-2xl p-8">
            {/* Logo */}
            <Link href="/" className="flex justify-center mb-8">
              <div className="relative h-12 w-[160px]">
                <Image 
                  src="/logo/combined/horizontal-standard-200x60.svg" 
                  alt="CellPilot" 
                  fill
                  className="object-contain"
                />
              </div>
            </Link>

            <h2 className="text-2xl font-bold text-center text-neutral-900 mb-2">
              {isSignUp ? 'Create Your Account' : 'Welcome Back'}
            </h2>
            <p className="text-center text-neutral-600 mb-8">
              {isSignUp 
                ? 'Get access to CellPilot for Google Sheets' 
                : 'Sign in to access your dashboard'}
            </p>

            {error && (
              <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">
                {error}
              </div>
            )}

            {success && (
              <div className="mb-4 p-3 bg-green-50 border border-green-200 rounded-lg text-green-700 text-sm">
                {success}
              </div>
            )}

            <form onSubmit={handleSubmit} className="space-y-4">
              {isSignUp && (
                <>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label htmlFor="firstName" className="block text-sm font-medium text-neutral-700 mb-1">
                        First Name
                      </label>
                      <input
                        id="firstName"
                        type="text"
                        value={firstName}
                        onChange={(e) => setFirstName(e.target.value)}
                        required={isSignUp}
                        className="w-full px-4 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all"
                        placeholder="John"
                      />
                    </div>
                    <div>
                      <label htmlFor="surname" className="block text-sm font-medium text-neutral-700 mb-1">
                        Surname <span className="text-neutral-400">(optional)</span>
                      </label>
                      <input
                        id="surname"
                        type="text"
                        value={surname}
                        onChange={(e) => setSurname(e.target.value)}
                        className="w-full px-4 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all"
                        placeholder="Doe"
                      />
                    </div>
                  </div>

                  <div>
                    <label htmlFor="company" className="block text-sm font-medium text-neutral-700 mb-1">
                      Company <span className="text-neutral-400">(optional)</span>
                    </label>
                    <input
                      id="company"
                      type="text"
                      value={company}
                      onChange={(e) => setCompany(e.target.value)}
                      className="w-full px-4 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all"
                      placeholder="Acme Inc."
                    />
                  </div>
                </>
              )}

              <div>
                <label htmlFor="email" className="block text-sm font-medium text-neutral-700 mb-1">
                  Email Address
                </label>
                <input
                  id="email"
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  required
                  className="w-full px-4 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all"
                  placeholder="you@example.com"
                />
              </div>

              <div>
                <label htmlFor="password" className="block text-sm font-medium text-neutral-700 mb-1">
                  Password
                </label>
                <div className="relative">
                  <input
                    id="password"
                    type={showPassword ? 'text' : 'password'}
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    required
                    minLength={6}
                    className="w-full px-4 py-2 pr-12 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all"
                    placeholder="••••••••"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-neutral-500 hover:text-neutral-700 focus:outline-none"
                  >
                    {showPassword ? (
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" />
                      </svg>
                    ) : (
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
                      </svg>
                    )}
                  </button>
                </div>
              </div>

              {isSignUp && (
                <div className="flex items-start">
                  <input
                    id="newsletter"
                    type="checkbox"
                    checked={newsletter}
                    onChange={(e) => setNewsletter(e.target.checked)}
                    className="mt-1 h-4 w-4 text-primary-600 border-neutral-300 rounded focus:ring-primary-500"
                  />
                  <label htmlFor="newsletter" className="ml-2 text-sm text-neutral-600">
                    Send me product updates and newsletter emails
                  </label>
                </div>
              )}

              <button
                type="submit"
                disabled={loading}
                className="w-full py-3 px-4 bg-gradient-to-r from-primary-600 to-primary-700 text-white font-semibold rounded-lg hover:from-primary-700 hover:to-primary-800 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
              >
                {loading ? 'Processing...' : (isSignUp ? 'Sign Up' : 'Sign In')}
              </button>
            </form>

            <div className="relative my-6">
              <div className="absolute inset-0 flex items-center">
                <div className="w-full border-t border-neutral-300"></div>
              </div>
              <div className="relative flex justify-center text-sm">
                <span className="px-2 bg-white text-neutral-500">Or</span>
              </div>
            </div>

            <button
              onClick={handleMagicLink}
              disabled={loading || !email}
              className="w-full py-3 px-4 bg-white border border-neutral-300 text-neutral-700 font-medium rounded-lg hover:bg-neutral-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
            >
              Send Magic Link to Email
            </button>

            <p className="mt-6 text-center text-sm text-neutral-600">
              {isSignUp ? 'Already have an account?' : "Don't have an account?"}{' '}
              <button
                onClick={() => {
                  setIsSignUp(!isSignUp)
                  setError(null)
                  setSuccess(null)
                }}
                className="text-primary-600 hover:text-primary-700 font-medium"
              >
                {isSignUp ? 'Sign In' : 'Sign Up'}
              </button>
            </p>
          </div>

          <p className="mt-8 text-center text-xs text-neutral-500">
            By signing up, you agree to our{' '}
            <Link href="/terms" className="text-primary-600 hover:text-primary-700">
              Terms of Service
            </Link>{' '}
            and{' '}
            <Link href="/privacy" className="text-primary-600 hover:text-primary-700">
              Privacy Policy
            </Link>
          </p>
        </div>
      </div>
    </>
  )
}