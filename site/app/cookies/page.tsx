import Link from 'next/link'
import type { Metadata } from 'next'
import GridAnimation from '@/components/GridAnimation'

export const metadata: Metadata = {
  title: 'Cookie Policy - CellPilot',
  description: 'Understand how CellPilot uses cookies and similar technologies to improve your experience.',
}

export default function CookiePolicy() {
  return (
    <main className="min-h-screen relative">
      <GridAnimation />
      {/* Background decoration matching main site */}
      <div className="absolute inset-0 overflow-hidden pointer-events-none" style={{zIndex: 0}}>
        <div className="absolute top-0 left-1/2 w-96 h-96 bg-pastel-mint/10 rounded-full blur-3xl"></div>
        <div className="absolute bottom-0 right-1/3 w-96 h-96 bg-pastel-sky/10 rounded-full blur-3xl"></div>
      </div>
      
      <div className="container-wrapper relative pt-[100px] sm:pt-[110px] md:pt-[120px] lg:pt-[130px] pb-20" style={{zIndex: 10}}>
        <div className="max-w-4xl mx-auto">
          {/* Back link */}
          <Link href="/" className="inline-flex items-center text-primary-600 hover:text-primary-700 transition-colors mb-8">
            <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
            </svg>
            Back to Home
          </Link>
          
          <h1 className="text-4xl sm:text-5xl font-bold text-neutral-900 mb-8">Cookie Policy</h1>
          
          <div className="prose prose-lg max-w-none space-y-8">
            <div className="glass-card rounded-2xl p-8">
              <p className="text-lg text-neutral-600 leading-relaxed">
                Last updated: {new Date().toLocaleDateString('en-GB', { year: 'numeric', month: 'long', day: 'numeric' })}
              </p>
              <p className="mt-4 text-neutral-600 leading-relaxed">
                This Cookie Policy explains how CellPilot uses cookies and similar technologies to recognize you when you visit our website 
                and use our Google Sheets add-on.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">What Are Cookies?</h2>
              <p className="text-neutral-600 leading-relaxed">
                Cookies are small data files placed on your device when you visit a website. They help websites remember information 
                about your visit, making your next visit easier and the site more useful to you.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Types of Cookies We Use</h2>
              
              <div className="space-y-4">
                <div className="p-4 bg-pastel-sky/10 rounded-lg">
                  <h3 className="text-lg font-semibold text-neutral-800 mb-2">Essential Cookies</h3>
                  <p className="text-neutral-600">
                    Required for the website to function properly. These cookies enable core functionality such as security, 
                    network management, and accessibility.
                  </p>
                  <ul className="list-disc list-inside mt-2 text-sm text-neutral-600">
                    <li>Session management</li>
                    <li>Authentication tokens</li>
                    <li>Security features</li>
                  </ul>
                </div>

                <div className="p-4 bg-pastel-mint/10 rounded-lg">
                  <h3 className="text-lg font-semibold text-neutral-800 mb-2">Functional Cookies</h3>
                  <p className="text-neutral-600">
                    Help us provide enhanced functionality and personalization. These cookies remember your preferences and choices.
                  </p>
                  <ul className="list-disc list-inside mt-2 text-sm text-neutral-600">
                    <li>Language preferences</li>
                    <li>User settings</li>
                    <li>Feature preferences</li>
                  </ul>
                </div>

                <div className="p-4 bg-pastel-lavender/10 rounded-lg">
                  <h3 className="text-lg font-semibold text-neutral-800 mb-2">Analytics Cookies</h3>
                  <p className="text-neutral-600">
                    Help us understand how visitors interact with our website by collecting and reporting information anonymously.
                  </p>
                  <ul className="list-disc list-inside mt-2 text-sm text-neutral-600">
                    <li>Page views and navigation</li>
                    <li>Time spent on pages</li>
                    <li>Error tracking</li>
                  </ul>
                </div>
              </div>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Third-Party Cookies</h2>
              <p className="text-neutral-600 leading-relaxed">
                We use services from third-party providers that may set their own cookies:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li><strong>Google Analytics:</strong> For website usage statistics</li>
                <li><strong>Stripe:</strong> For secure payment processing</li>
                <li><strong>Google OAuth:</strong> For authentication services</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Cookie Duration</h2>
              <div className="space-y-4 text-neutral-600">
                <p><strong>Session Cookies:</strong> Deleted when you close your browser</p>
                <p><strong>Persistent Cookies:</strong> Remain on your device for a set period:</p>
                <ul className="list-disc list-inside space-y-2 ml-4">
                  <li>Authentication: 30 days</li>
                  <li>Preferences: 1 year</li>
                  <li>Analytics: 2 years</li>
                </ul>
              </div>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Managing Cookies</h2>
              <p className="text-neutral-600 leading-relaxed mb-4">
                You can control and manage cookies in various ways:
              </p>
              
              <div className="space-y-4">
                <div className="p-4 bg-neutral-50 rounded-lg">
                  <h3 className="font-semibold text-neutral-800 mb-2">Browser Settings</h3>
                  <p className="text-sm text-neutral-600">
                    Most browsers allow you to refuse or accept cookies. Check your browser's help section for instructions.
                  </p>
                </div>
                
                <div className="p-4 bg-neutral-50 rounded-lg">
                  <h3 className="font-semibold text-neutral-800 mb-2">Cookie Preferences</h3>
                  <p className="text-sm text-neutral-600">
                    You can adjust your cookie preferences when you first visit our website or anytime through our cookie settings.
                  </p>
                </div>
                
                <div className="p-4 bg-neutral-50 rounded-lg">
                  <h3 className="font-semibold text-neutral-800 mb-2">Opt-Out Links</h3>
                  <p className="text-sm text-neutral-600">
                    Google Analytics: <a href="https://tools.google.com/dlpage/gaoptout" className="text-primary-600 hover:text-primary-700" target="_blank" rel="noopener noreferrer">Opt-out Browser Add-on</a>
                  </p>
                </div>
              </div>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Impact of Disabling Cookies</h2>
              <p className="text-neutral-600 leading-relaxed">
                If you disable cookies, some features of CellPilot may not function properly:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>You may need to re-enter login information</li>
                <li>Some features may be unavailable</li>
                <li>Preferences won't be saved between sessions</li>
                <li>Usage tracking for your account may be affected</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Updates to This Policy</h2>
              <p className="text-neutral-600 leading-relaxed">
                We may update this Cookie Policy from time to time to reflect changes in our practices or for operational, legal, 
                or regulatory reasons. We will notify you of any material changes by posting the new policy on this page.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Contact Us</h2>
              <p className="text-neutral-600 leading-relaxed">
                If you have questions about our use of cookies, please contact us at:
              </p>
              <div className="mt-4 p-4 bg-neutral-50 rounded-lg">
                <p className="text-neutral-700">Email: privacy@cellpilot.app</p>
                <p className="text-neutral-700">Address: CellPilot Ltd, London, United Kingdom</p>
              </div>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Related Policies</h2>
              <p className="text-neutral-600 leading-relaxed">
                For more information about how we handle your data, please see:
              </p>
              <div className="flex gap-4 mt-4">
                <Link href="/privacy" className="text-primary-600 hover:text-primary-700 transition-colors">
                  Privacy Policy
                </Link>
                <Link href="/terms" className="text-primary-600 hover:text-primary-700 transition-colors">
                  Terms of Service
                </Link>
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>
  )
}