import Link from 'next/link'
import type { Metadata } from 'next'

export const metadata: Metadata = {
  title: 'Privacy Policy - CellPilot',
  description: 'Learn how CellPilot protects your data and respects your privacy while using our Google Sheets automation tools.',
}

export default function PrivacyPolicy() {
  return (
    <main className="min-h-screen relative">
      {/* Background decoration matching main site */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute top-0 left-0 w-96 h-96 bg-pastel-sky/10 rounded-full blur-3xl"></div>
        <div className="absolute bottom-0 right-0 w-96 h-96 bg-pastel-mint/10 rounded-full blur-3xl"></div>
      </div>
      
      <div className="container-wrapper relative z-10 pt-[100px] sm:pt-32 md:pt-28 pb-20">
        <div className="max-w-4xl mx-auto">
          {/* Back link */}
          <Link href="/" className="inline-flex items-center text-primary-600 hover:text-primary-700 transition-colors mb-8">
            <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
            </svg>
            Back to Home
          </Link>
          
          <h1 className="text-4xl sm:text-5xl font-bold text-neutral-900 mb-8">Privacy Policy</h1>
          
          <div className="prose prose-lg max-w-none space-y-8">
            <div className="glass-card rounded-2xl p-8">
              <p className="text-lg text-neutral-600 leading-relaxed">
                Last updated: {new Date().toLocaleDateString('en-GB', { year: 'numeric', month: 'long', day: 'numeric' })}
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">1. Information We Collect</h2>
              <p className="text-neutral-600 leading-relaxed">
                CellPilot collects minimal information necessary to provide our services:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Google account information (email, name) when you authorize the add-on</li>
                <li>Usage data to improve our services (operations performed, features used)</li>
                <li>Spreadsheet metadata required for operations (no content is stored)</li>
                <li>Payment information for paid subscriptions (processed securely via Stripe)</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">2. How We Use Your Information</h2>
              <p className="text-neutral-600 leading-relaxed">
                We use collected information to:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Provide and maintain CellPilot services</li>
                <li>Process your operations and requests</li>
                <li>Send important service updates and notifications</li>
                <li>Improve and optimize our features</li>
                <li>Provide customer support</li>
                <li>Comply with legal obligations</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">3. Data Security</h2>
              <p className="text-neutral-600 leading-relaxed">
                We implement industry-standard security measures to protect your data:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>All data transmission is encrypted using SSL/TLS</li>
                <li>Access to user data is strictly limited and monitored</li>
                <li>We do not store your spreadsheet content</li>
                <li>Regular security audits and updates</li>
                <li>Compliance with Google's OAuth 2.0 security standards</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">4. Data Sharing</h2>
              <p className="text-neutral-600 leading-relaxed">
                We do not sell, trade, or rent your personal information. We may share data only:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>With your explicit consent</li>
                <li>To comply with legal requirements</li>
                <li>With service providers who assist our operations (under strict confidentiality)</li>
                <li>To protect rights, safety, and property</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">5. Your Rights</h2>
              <p className="text-neutral-600 leading-relaxed">
                You have the right to:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Access your personal data</li>
                <li>Correct inaccurate data</li>
                <li>Request deletion of your data</li>
                <li>Object to data processing</li>
                <li>Export your data</li>
                <li>Withdraw consent at any time</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">6. Cookies</h2>
              <p className="text-neutral-600 leading-relaxed">
                We use essential cookies to maintain your session and preferences. You can control cookie settings in your browser. 
                For more details, see our <Link href="/cookies" className="text-primary-600 hover:text-primary-700">Cookie Policy</Link>.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">7. Children's Privacy</h2>
              <p className="text-neutral-600 leading-relaxed">
                CellPilot is not intended for children under 13. We do not knowingly collect personal information from children under 13.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">8. Changes to This Policy</h2>
              <p className="text-neutral-600 leading-relaxed">
                We may update this Privacy Policy periodically. We will notify you of significant changes via email or through the service.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">9. Contact Us</h2>
              <p className="text-neutral-600 leading-relaxed">
                If you have questions about this Privacy Policy or our data practices, please contact us at:
              </p>
              <div className="mt-4 p-4 bg-neutral-50 rounded-lg">
                <p className="text-neutral-700">Email: privacy@cellpilot.app</p>
                <p className="text-neutral-700">Address: CellPilot Ltd, London, United Kingdom</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>
  )
}