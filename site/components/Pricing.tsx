'use client'

import Link from 'next/link'

const plans = [
  {
    name: 'Free',
    price: '£0',
    period: 'forever',
    description: 'Perfect for trying out CellPilot',
    features: [
      '25 operations per month',
      'Basic duplicate removal',
      'Text standardization',
      'Community support',
    ],
    cta: 'Get Started',
    ctaLink: '/install',
    popular: false,
  },
  {
    name: 'Starter',
    price: '£5.99',
    period: 'per month',
    description: 'For individuals and small teams',
    features: [
      '500 operations per month',
      'All data cleaning tools',
      'Formula builder',
      'Email support',
      'Automatic backups',
    ],
    cta: 'Start Free Trial',
    ctaLink: '/install?plan=starter',
    popular: true,
  },
  {
    name: 'Professional',
    price: '£11.99',
    period: 'per month',
    description: 'For power users and growing teams',
    features: [
      'Unlimited operations',
      'All features included',
      'Email automation',
      'Calendar integration',
      'Priority support',
      'Advanced analytics',
    ],
    cta: 'Start Free Trial',
    ctaLink: '/install?plan=professional',
    popular: false,
  },
]

const Pricing = () => {
  return (
    <section id="pricing" className="py-20 bg-gray-50">
      <div className="container mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center">
          <h2 className="text-3xl font-bold text-gray-900 sm:text-4xl">
            Simple, transparent pricing
          </h2>
          <p className="mt-4 text-lg text-gray-600">
            Choose the plan that fits your needs. Upgrade or downgrade anytime.
          </p>
        </div>

        <div className="mt-16 grid grid-cols-1 gap-8 lg:grid-cols-3">
          {plans.map((plan) => (
            <div
              key={plan.name}
              className={`relative bg-white rounded-lg shadow-sm ${
                plan.popular ? 'ring-2 ring-primary-500' : 'border border-gray-200'
              }`}
            >
              {plan.popular && (
                <div className="absolute -top-4 left-1/2 -translate-x-1/2">
                  <span className="bg-primary-500 text-white px-3 py-1 text-sm font-medium rounded-full">
                    Most Popular
                  </span>
                </div>
              )}

              <div className="p-6">
                <h3 className="text-xl font-semibold text-gray-900">{plan.name}</h3>
                <p className="mt-2 text-gray-600">{plan.description}</p>
                
                <div className="mt-4">
                  <span className="text-4xl font-bold text-gray-900">{plan.price}</span>
                  <span className="text-gray-600 ml-2">{plan.period}</span>
                </div>

                <ul className="mt-6 space-y-3">
                  {plan.features.map((feature, index) => (
                    <li key={index} className="flex items-start">
                      <svg
                        className="h-5 w-5 text-green-500 mt-0.5 mr-3 flex-shrink-0"
                        fill="currentColor"
                        viewBox="0 0 20 20"
                      >
                        <path
                          fillRule="evenodd"
                          d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z"
                          clipRule="evenodd"
                        />
                      </svg>
                      <span className="text-gray-700">{feature}</span>
                    </li>
                  ))}
                </ul>

                <Link
                  href={plan.ctaLink}
                  className={`mt-8 block w-full py-3 px-4 rounded-lg text-center font-semibold transition-colors ${
                    plan.popular
                      ? 'bg-primary-600 text-white hover:bg-primary-700'
                      : 'bg-gray-100 text-gray-900 hover:bg-gray-200'
                  }`}
                >
                  {plan.cta}
                </Link>
              </div>
            </div>
          ))}
        </div>

        <div className="mt-12 text-center text-sm text-gray-600">
          <p>All plans include a 14-day free trial. No credit card required.</p>
          <p className="mt-2">
            Need a custom plan?{' '}
            <Link href="/contact" className="text-primary-600 hover:text-primary-700 font-medium">
              Contact us
            </Link>
          </p>
        </div>
      </div>
    </section>
  )
}

export default Pricing