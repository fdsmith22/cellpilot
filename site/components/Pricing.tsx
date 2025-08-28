'use client'

import { useState } from 'react'
import Link from 'next/link'
import ScrollObserver from './ScrollObserver'

const plans = [
  {
    name: 'Free',
    price: '£0',
    period: 'forever',
    description: 'Get started with basic features',
    features: [
      '25 operations per month',
      'Duplicate removal',
      'Text cleaning',
      'Basic formulas',
      'Email support',
    ],
    cta: 'Get Started',
    ctaLink: '/install',
    popular: false,
  },
  {
    name: 'Starter',
    price: '£5.99',
    period: 'per month',
    description: 'For regular spreadsheet users',
    features: [
      '500 operations per month',
      'All cleaning tools',
      'Advanced formulas',
      'Automatic backups',
      'Priority support',
      'Usage analytics',
    ],
    cta: 'Try Free for 7 Days',
    ctaLink: '/install?plan=starter',
    popular: true,
  },
  {
    name: 'Professional',
    price: '£11.99',
    period: 'per month',
    description: 'For power users and teams',
    features: [
      'Unlimited operations',
      'All features included',
      'Email automation',
      'Calendar integration',
      'API access',
      'Dedicated support',
    ],
    cta: 'Try Free for 7 Days',
    ctaLink: '/install?plan=professional',
    popular: false,
  },
]

const Pricing = () => {
  const [selectedPlan, setSelectedPlan] = useState<string>('Starter')
  
  return (
    <section id="pricing" className="snap-section h-screen-vh md:h-screen relative overflow-hidden bg-transparent">
      {/* Background decoration */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[400px] sm:w-[600px] md:w-[800px] h-[400px] sm:h-[600px] md:h-[800px] bg-gradient-to-r from-pastel-peach/20 to-pastel-lavender/20 rounded-full blur-3xl"></div>
      </div>
      
      {/* Seamless transition */}
      <div className="absolute inset-0 bg-gradient-to-b from-white/60 via-white/70 to-white/80"></div>
      
      {/* Removed grid pattern - using global grid */}
      
      <div className="container-wrapper relative z-elevated h-screen-vh md:h-screen pt-header-total-mobile md:pt-header-total pb-safe-bottom md:pb-4 flex flex-col justify-center">
        <div className="w-full max-h-full overflow-y-auto">
        <ScrollObserver className="text-center max-w-3xl mx-auto mb-6 sm:mb-8 px-4">
          <h2 className="text-2xl sm:text-3xl md:text-4xl lg:text-5xl font-bold text-neutral-900 mb-3 sm:mb-4">
            Simple Pricing
          </h2>
          <p className="text-sm sm:text-base md:text-lg text-neutral-600 leading-relaxed">
            Start free and upgrade when you need more. Cancel anytime.
          </p>
        </ScrollObserver>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-3 sm:gap-4 items-stretch px-4 sm:px-0">
          {plans.map((plan, index) => (
            <ScrollObserver
              key={plan.name}
              className={`scroll-observer-delay-${index * 100}`}
            >
              <div
                onClick={() => setSelectedPlan(plan.name)}
                className={`glass-card rounded-xl sm:rounded-2xl overflow-hidden h-full cursor-pointer transition-all duration-300 ${
                  selectedPlan === plan.name 
                    ? 'ring-2 ring-primary-500 scale-105 shadow-2xl bg-white/90' 
                    : 'hover:scale-102 hover:shadow-xl'
                }`}
              >
              {(selectedPlan === plan.name || (plan.popular && !selectedPlan)) && (
                <div className="bg-gradient-to-r from-primary-500 to-accent-teal px-4 py-2 text-center">
                  <span className="text-white text-sm font-medium">
                    {selectedPlan === plan.name ? 'Selected Plan' : 'Most Popular'}
                  </span>
                </div>
              )}
              
              <div className="p-4 sm:p-6">
                <h3 className="text-xl sm:text-2xl font-bold text-neutral-900 mb-2">{plan.name}</h3>
                <p className="text-sm sm:text-base text-neutral-600 mb-4 sm:mb-6">{plan.description}</p>
                
                <div className="mb-6 sm:mb-8">
                  <span className="text-3xl sm:text-4xl font-bold text-neutral-900">{plan.price}</span>
                  <span className="text-sm sm:text-base text-neutral-600 ml-1 sm:ml-2">/{plan.period}</span>
                </div>

                <ul className="space-y-2 sm:space-y-3 mb-6 sm:mb-8">
                  {plan.features.map((feature, index) => (
                    <li key={index} className="flex items-start">
                      <svg
                        className="h-5 w-5 text-accent-teal mt-0.5 mr-3 flex-shrink-0"
                        fill="currentColor"
                        viewBox="0 0 20 20"
                      >
                        <path
                          fillRule="evenodd"
                          d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z"
                          clipRule="evenodd"
                        />
                      </svg>
                      <span className="text-sm sm:text-base text-neutral-700">{feature}</span>
                    </li>
                  ))}
                </ul>

                <Link
                  href={plan.ctaLink}
                  className={`block w-full py-3 px-6 rounded-full text-center font-medium transition-all ${
                    plan.popular
                      ? 'btn-primary'
                      : 'bg-white border border-neutral-300 text-neutral-700 hover:bg-neutral-100'
                  }`}
                >
                  {plan.cta}
                </Link>
              </div>
              </div>
            </ScrollObserver>
          ))}
        </div>

        <div className="mt-8 sm:mt-10 text-center px-4">
          <div className="inline-flex flex-col sm:flex-row items-start sm:items-center gap-4 sm:gap-8 p-4 sm:p-6 rounded-xl sm:rounded-2xl glass-card">
            <div className="flex items-center">
              <svg className="w-4 sm:w-5 h-4 sm:h-5 text-accent-teal mr-2 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              <span className="text-sm sm:text-base text-neutral-700 font-medium">No credit card required</span>
            </div>
            <div className="flex items-center">
              <svg className="w-4 sm:w-5 h-4 sm:h-5 text-accent-teal mr-2 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              <span className="text-sm sm:text-base text-neutral-700 font-medium">Cancel anytime</span>
            </div>
            <div className="flex items-center">
              <svg className="w-4 sm:w-5 h-4 sm:h-5 text-accent-teal mr-2 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              <span className="text-sm sm:text-base text-neutral-700 font-medium">Secure payments</span>
            </div>
          </div>
          </div>
        </div>
      </div>
    </section>
  )
}

export default Pricing