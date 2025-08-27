import Link from 'next/link'

const plans = [
  {
    name: 'Free',
    price: 0,
    currency: '£',
    period: 'forever',
    stripeId: null,
    description: 'Perfect for trying out CellPilot',
    features: [
      '25 operations per month',
      'Basic duplicate removal',
      'Text standardization',
      'Community support',
      'Basic analytics'
    ],
    limitations: [
      'No formula builder',
      'No automation features',
      'No priority support'
    ],
    cta: 'Get Started',
    ctaLink: '/install',
    popular: false,
  },
  {
    name: 'Starter',
    price: 5.99,
    currency: '£',
    period: 'per month',
    stripeId: 'price_starter_monthly',
    description: 'For individuals and small teams',
    features: [
      '500 operations per month',
      'All data cleaning tools',
      'Formula builder',
      'Email support',
      'Automatic backups',
      'Basic automation',
      'Export to CSV/PDF'
    ],
    limitations: [
      'No calendar integration',
      'Standard support only'
    ],
    cta: 'Start 14-day Trial',
    ctaLink: '/signup?plan=starter',
    popular: true,
  },
  {
    name: 'Professional',
    price: 11.99,
    currency: '£',
    period: 'per month',
    stripeId: 'price_professional_monthly',
    description: 'For power users and growing teams',
    features: [
      'Unlimited operations',
      'All features included',
      'Email automation',
      'Calendar integration',
      'Priority support',
      'Advanced analytics',
      'Custom formulas',
      'API access',
      'Team collaboration'
    ],
    limitations: [],
    cta: 'Start 14-day Trial',
    ctaLink: '/signup?plan=professional',
    popular: false,
  },
  {
    name: 'Business',
    price: 19.99,
    currency: '£',
    period: 'per month per user',
    stripeId: 'price_business_monthly',
    description: 'For teams and organizations',
    features: [
      'Everything in Professional',
      'Unlimited team members',
      'SSO/SAML',
      'Admin dashboard',
      'Usage reporting',
      'Custom integrations',
      'SLA guarantee',
      'Dedicated support',
      'Training sessions'
    ],
    limitations: [],
    cta: 'Contact Sales',
    ctaLink: '/contact',
    popular: false,
  }
]

export default function PricingPage() {
  return (
    <main className="min-h-screen relative">
      {/* Background decoration matching main site */}
      <div className="absolute inset-0 overflow-hidden pointer-events-none" style={{zIndex: 0}}>
        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[800px] h-[800px] bg-gradient-to-r from-pastel-peach/20 to-pastel-lavender/20 rounded-full blur-3xl"></div>
      </div>
      
      <div className="container mx-auto px-4 pt-[100px] sm:pt-32 md:pt-28 pb-16 sm:px-6 lg:px-8 relative" style={{zIndex: 10}}>
        <div className="text-center mb-12">
          <h1 className="text-4xl font-bold text-neutral-900 sm:text-5xl">
            Simple, transparent pricing
          </h1>
          <p className="mt-4 text-xl text-neutral-600">
            Choose the plan that fits your needs. Upgrade or downgrade anytime.
          </p>
        </div>

        {/* Billing Toggle */}
        <div className="flex justify-center mb-12">
          <div className="glass-card p-1 rounded-lg flex">
            <button className="px-4 py-2 bg-white rounded-md text-neutral-900 font-medium">
              Monthly
            </button>
            <button className="px-4 py-2 text-neutral-600 font-medium">
              Annual (Save 20%)
            </button>
          </div>
        </div>

        {/* Pricing Cards */}
        <div className="grid grid-cols-1 gap-8 lg:grid-cols-4 max-w-7xl mx-auto">
          {plans.map((plan) => (
            <div
              key={plan.name}
              className={`relative bg-white rounded-lg ${
                plan.popular 
                  ? 'ring-2 ring-primary-500 shadow-xl scale-105' 
                  : 'border border-gray-200 shadow-sm'
              }`}
            >
              {plan.popular && (
                <div className="absolute -top-5 left-0 right-0 mx-auto w-32">
                  <div className="bg-primary-500 text-white text-sm text-center py-1 px-3 rounded-full">
                    Most Popular
                  </div>
                </div>
              )}

              <div className="p-6">
                <h3 className="text-xl font-semibold text-neutral-900">{plan.name}</h3>
                <p className="mt-2 text-sm text-neutral-600">{plan.description}</p>
                
                <div className="mt-4 flex items-baseline">
                  <span className="text-4xl font-bold text-neutral-900">
                    {plan.currency}{plan.price}
                  </span>
                  <span className="ml-2 text-neutral-600">{plan.period}</span>
                </div>

                <Link
                  href={plan.ctaLink}
                  className={`mt-6 block w-full py-3 px-4 rounded-lg text-center font-semibold transition-colors ${
                    plan.popular
                      ? 'bg-primary-600 text-white hover:bg-primary-700'
                      : 'glass-card text-neutral-900 hover:bg-gray-200'
                  }`}
                >
                  {plan.cta}
                </Link>

                <div className="mt-6">
                  <h4 className="text-sm font-semibold text-neutral-900 mb-3">
                    What&apos;s included:
                  </h4>
                  <ul className="space-y-2">
                    {plan.features.map((feature, index) => (
                      <li key={index} className="flex items-start">
                        <svg
                          className="h-5 w-5 text-green-500 mt-0.5 mr-2 flex-shrink-0"
                          fill="currentColor"
                          viewBox="0 0 20 20"
                        >
                          <path
                            fillRule="evenodd"
                            d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z"
                            clipRule="evenodd"
                          />
                        </svg>
                        <span className="text-sm text-gray-700">{feature}</span>
                      </li>
                    ))}
                  </ul>

                  {plan.limitations.length > 0 && (
                    <>
                      <div className="mt-4 pt-4 border-t border-gray-200">
                        <ul className="space-y-2">
                          {plan.limitations.map((limitation, index) => (
                            <li key={index} className="flex items-start">
                              <svg
                                className="h-5 w-5 text-gray-400 mt-0.5 mr-2 flex-shrink-0"
                                fill="currentColor"
                                viewBox="0 0 20 20"
                              >
                                <path
                                  fillRule="evenodd"
                                  d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z"
                                  clipRule="evenodd"
                                />
                              </svg>
                              <span className="text-sm text-gray-500">{limitation}</span>
                            </li>
                          ))}
                        </ul>
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>
          ))}
        </div>

        {/* FAQ Section */}
        <div className="mt-20 max-w-3xl mx-auto">
          <h2 className="text-2xl font-bold text-neutral-900 text-center mb-8">
            Frequently Asked Questions
          </h2>
          
          <div className="space-y-6">
            <div className="bg-white rounded-lg p-6 shadow-sm">
              <h3 className="font-semibold text-neutral-900 mb-2">
                Can I change plans anytime?
              </h3>
              <p className="text-neutral-600">
                Yes, you can upgrade or downgrade your plan at any time. Changes take effect immediately, 
                and we&apos;ll prorate any payments.
              </p>
            </div>

            <div className="bg-white rounded-lg p-6 shadow-sm">
              <h3 className="font-semibold text-neutral-900 mb-2">
                Do you offer refunds?
              </h3>
              <p className="text-neutral-600">
                We offer a 14-day money-back guarantee on all paid plans. If you&apos;re not satisfied, 
                contact support for a full refund.
              </p>
            </div>

            <div className="bg-white rounded-lg p-6 shadow-sm">
              <h3 className="font-semibold text-neutral-900 mb-2">
                What happens when I reach my operation limit?
              </h3>
              <p className="text-neutral-600">
                You&apos;ll receive a notification when you reach 80% of your limit. Once you hit 100%, 
                you can either wait for the next billing cycle or upgrade your plan for immediate access.
              </p>
            </div>

            <div className="bg-white rounded-lg p-6 shadow-sm">
              <h3 className="font-semibold text-neutral-900 mb-2">
                Is my data secure?
              </h3>
              <p className="text-neutral-600">
                Absolutely. CellPilot only accesses the spreadsheets you explicitly work with. 
                We never store your spreadsheet data on our servers, and all operations happen 
                directly in your Google account.
              </p>
            </div>
          </div>
        </div>

        {/* CTA Section */}
        <div className="mt-20 text-center">
          <h2 className="text-2xl font-bold text-neutral-900 mb-4">
            Ready to get started?
          </h2>
          <p className="text-lg text-neutral-600 mb-8">
            Try CellPilot free for 14 days. No credit card required.
          </p>
          <Link
            href="/install"
            className="inline-block bg-primary-600 text-white px-8 py-3 rounded-lg font-semibold hover:bg-primary-700 transition-colors"
          >
            Start Free Trial
          </Link>
        </div>
      </div>
    </main>
  )
}