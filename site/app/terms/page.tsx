import Link from 'next/link'
import type { Metadata } from 'next'

export const metadata: Metadata = {
  title: 'Terms of Service - CellPilot',
  description: 'Read the CellPilot terms of service for using our Google Sheets automation add-on and services.',
}

export default function TermsOfService() {
  return (
    <main className="min-h-screen relative">
      {/* Background decoration matching main site */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute top-0 right-0 w-96 h-96 bg-pastel-lavender/10 rounded-full blur-3xl"></div>
        <div className="absolute bottom-0 left-0 w-96 h-96 bg-pastel-peach/10 rounded-full blur-3xl"></div>
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
          
          <h1 className="text-4xl sm:text-5xl font-bold text-neutral-900 mb-8">Terms of Service</h1>
          
          <div className="prose prose-lg max-w-none space-y-8">
            <div className="glass-card rounded-2xl p-8">
              <p className="text-lg text-neutral-600 leading-relaxed">
                Last updated: {new Date().toLocaleDateString('en-GB', { year: 'numeric', month: 'long', day: 'numeric' })}
              </p>
              <p className="mt-4 text-neutral-600 leading-relaxed">
                By using CellPilot, you agree to these Terms of Service. Please read them carefully.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">1. Acceptance of Terms</h2>
              <p className="text-neutral-600 leading-relaxed">
                By accessing or using CellPilot, you agree to be bound by these Terms of Service and all applicable laws and regulations. 
                If you do not agree with any of these terms, you are prohibited from using or accessing this service.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">2. Service Description</h2>
              <p className="text-neutral-600 leading-relaxed">
                CellPilot is a Google Sheets add-on that provides automation tools and features including:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Data cleaning and formatting tools</li>
                <li>Duplicate removal functionality</li>
                <li>Formula generation assistance</li>
                <li>Email automation features</li>
                <li>Usage tracking and analytics</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">3. User Accounts</h2>
              <p className="text-neutral-600 leading-relaxed">
                To use CellPilot, you must:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Have a valid Google account</li>
                <li>Provide accurate and complete information</li>
                <li>Maintain the security of your account</li>
                <li>Notify us immediately of any unauthorized access</li>
                <li>Be responsible for all activities under your account</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">4. Subscription and Billing</h2>
              <div className="space-y-4 text-neutral-600">
                <p><strong>Free Plan:</strong> 25 operations per month at no cost</p>
                <p><strong>Paid Plans:</strong> Starter (£5.99/month) and Professional (£11.99/month)</p>
                <ul className="list-disc list-inside space-y-2 mt-4">
                  <li>7-day free trial for paid plans</li>
                  <li>Automatic monthly renewal unless cancelled</li>
                  <li>Cancellation takes effect at the end of the billing period</li>
                  <li>No refunds for partial months</li>
                  <li>Prices subject to change with 30 days notice</li>
                </ul>
              </div>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">5. Acceptable Use</h2>
              <p className="text-neutral-600 leading-relaxed">
                You agree not to:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Use the service for any unlawful purpose</li>
                <li>Attempt to gain unauthorized access to any part of the service</li>
                <li>Interfere with or disrupt the service</li>
                <li>Transmit viruses or malicious code</li>
                <li>Share your account with others</li>
                <li>Reverse engineer or attempt to extract source code</li>
                <li>Use the service to process sensitive personal data without proper authorization</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">6. Intellectual Property</h2>
              <p className="text-neutral-600 leading-relaxed">
                All content, features, and functionality of CellPilot are owned by CellPilot Ltd and are protected by international copyright, 
                trademark, and other intellectual property laws. You retain all rights to your spreadsheet data.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">7. Data Processing</h2>
              <p className="text-neutral-600 leading-relaxed">
                CellPilot processes your spreadsheet data to provide the requested services. We:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Process data only as necessary to perform operations</li>
                <li>Do not store your spreadsheet content permanently</li>
                <li>Implement automatic backups before operations</li>
                <li>Delete temporary data after processing</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">8. Limitation of Liability</h2>
              <p className="text-neutral-600 leading-relaxed">
                CellPilot is provided "as is" without warranties of any kind. We are not liable for:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Data loss or corruption</li>
                <li>Service interruptions or delays</li>
                <li>Indirect, incidental, or consequential damages</li>
                <li>Loss of profits or business opportunities</li>
              </ul>
              <p className="mt-4 text-neutral-600">
                Our total liability shall not exceed the amount paid by you in the last 12 months.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">9. Indemnification</h2>
              <p className="text-neutral-600 leading-relaxed">
                You agree to indemnify and hold harmless CellPilot Ltd, its officers, directors, employees, and agents from any claims, 
                damages, or expenses arising from your use of the service or violation of these terms.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">10. Termination</h2>
              <p className="text-neutral-600 leading-relaxed">
                We may terminate or suspend your account immediately, without prior notice, for:
              </p>
              <ul className="list-disc list-inside space-y-2 text-neutral-600">
                <li>Breach of these Terms of Service</li>
                <li>Fraudulent or illegal activity</li>
                <li>Non-payment of fees</li>
                <li>At our sole discretion</li>
              </ul>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">11. Changes to Terms</h2>
              <p className="text-neutral-600 leading-relaxed">
                We reserve the right to modify these terms at any time. Material changes will be notified via email or through the service 
                at least 30 days before taking effect. Continued use after changes constitutes acceptance.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">12. Governing Law</h2>
              <p className="text-neutral-600 leading-relaxed">
                These Terms of Service are governed by the laws of the United Kingdom. Any disputes shall be resolved in the courts of 
                England and Wales.
              </p>
            </div>

            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">13. Contact Information</h2>
              <p className="text-neutral-600 leading-relaxed">
                For questions about these Terms of Service, please contact us at:
              </p>
              <div className="mt-4 p-4 bg-neutral-50 rounded-lg">
                <p className="text-neutral-700">Email: legal@cellpilot.app</p>
                <p className="text-neutral-700">Address: CellPilot Ltd, London, United Kingdom</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>
  )
}