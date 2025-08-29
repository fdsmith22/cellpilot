'use client'

import Link from 'next/link'
import Image from 'next/image'

const Footer = () => {
  const currentYear = new Date().getFullYear()

  const footerLinks = {
    Product: [
      { name: 'Features', href: '#features' },
      { name: 'Pricing', href: '#pricing' },
      { name: 'Documentation', href: '/docs' },
      { name: 'Changelog', href: '/changelog' },
    ],
    Support: [
      { name: 'Help Center', href: '/help' },
      { name: 'Contact', href: '/contact' },
      { name: 'Status', href: '/status' },
      { name: 'API', href: '/api' },
    ],
    Company: [
      { name: 'About', href: '/about' },
      { name: 'Privacy', href: '/privacy' },
      { name: 'Terms', href: '/terms' },
      { name: 'Cookies', href: '/cookies' },
    ],
  }

  return (
    <footer className="relative overflow-hidden" style={{scrollSnapAlign: 'none', scrollSnapStop: 'normal'}}>
      {/* Subtle animated background */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute -bottom-32 -left-32 w-64 h-64 bg-pastel-lavender/10 rounded-full blur-3xl"></div>
        <div className="absolute -bottom-32 -right-32 w-64 h-64 bg-pastel-mint/10 rounded-full blur-3xl"></div>
      </div>
      
      {/* Seamless transition from sections above */}
      <div className="absolute inset-0 bg-gradient-to-b from-neutral-50/90 via-white/95 to-primary-50/5"></div>
      
      {/* Grid pattern overlay */}
      <div className="absolute inset-0 grid-pattern opacity-15"></div>
      
      <div className="relative z-elevated">
        <div className="container-wrapper py-16 lg:py-20">
          {/* Main Footer Content */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-8 mb-12">
            {/* Brand Column */}
            <div className="col-span-2 md:col-span-1">
              <Link href="/" className="inline-block mb-6">
                <div className="relative h-12 w-[160px]">
                  <Image 
                    src="/logo/combined/horizontal-standard-200x60.svg" 
                    alt="CellPilot - Google Sheets Automation" 
                    fill
                    className="object-contain object-left"
                  />
                </div>
              </Link>
              <p className="text-sm text-neutral-600 mb-6 leading-relaxed">
                Transform your spreadsheet workflow with intelligent automation.
              </p>
            </div>

          {/* Links Columns */}
          {Object.entries(footerLinks).map(([category, links]) => (
            <div key={category}>
              <h3 className="text-sm font-semibold text-neutral-800 tracking-wide mb-4">{category}</h3>
              <ul className="space-y-3">
                {links.map((link) => (
                  <li key={link.name}>
                    <a
                      href={link.href}
                      className="text-sm text-neutral-600 hover:text-primary-600 transition-all hover:translate-x-1 inline-block"
                      onClick={(e) => {
                        if (link.href.startsWith('#')) {
                          e.preventDefault()
                          const element = document.querySelector(link.href)
                          element?.scrollIntoView({ behavior: 'smooth' })
                        }
                      }}
                    >
                      {link.name}
                    </a>
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>

          {/* Bottom Bar */}
          <div className="pt-6 sm:pt-8 mt-6 sm:mt-8 border-t border-neutral-200/50">
            <div className="flex justify-center">
              <p className="text-xs sm:text-sm text-neutral-500 text-center">
                © {currentYear} CellPilot. All rights reserved.
              </p>
            </div>
          </div>
        </div>
      </div>
    </footer>
  )
}

export default Footer