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
    <footer className="relative overflow-hidden bg-gradient-to-b from-white via-pastel-lavender/5 to-pastel-lavender/10 flex-shrink-0">
      {/* Subtle animated background */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute -bottom-32 -left-32 w-96 h-96 bg-pastel-mint/20 rounded-full blur-3xl"></div>
        <div className="absolute -bottom-32 -right-32 w-96 h-96 bg-pastel-sky/20 rounded-full blur-3xl"></div>
      </div>
      
      {/* Grid pattern overlay */}
      <div className="absolute inset-0 grid-pattern opacity-10"></div>
      
      <div className="relative z-elevated">
        <div className="container-wrapper pt-16 pb-12">
          {/* Main Footer Content */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-8 lg:gap-12">
            {/* Brand Column - Wider */}
            <div className="col-span-2 md:col-span-1 lg:pr-8">
              <Link href="/" className="inline-block mb-4 group">
                <div className="relative h-10 w-[140px] transition-transform group-hover:scale-105">
                  <Image 
                    src="/logo/combined/horizontal-standard-200x60.svg" 
                    alt="CellPilot - Google Sheets Automation" 
                    fill
                    className="object-contain object-left"
                  />
                </div>
              </Link>
              <p className="text-sm text-neutral-600 leading-relaxed mb-6">
                Transform your spreadsheet workflow with intelligent automation.
              </p>
              <p className="text-xs text-neutral-500">
                © {currentYear} CellPilot
                <br />
                All rights reserved.
              </p>
            </div>

          {/* Links Columns - Better spacing */}
          {Object.entries(footerLinks).map(([category, links]) => (
            <div key={category} className="col-span-1">
              <h3 className="text-sm font-bold text-neutral-900 mb-4 uppercase tracking-wider">{category}</h3>
              <ul className="space-y-2.5">
                {links.map((link) => (
                  <li key={link.name}>
                    <a
                      href={link.href}
                      className="text-sm text-neutral-600 hover:text-primary-600 transition-all duration-200 hover:translate-x-1 inline-block relative after:content-[''] after:absolute after:bottom-0 after:left-0 after:w-0 after:h-[1px] after:bg-primary-600 after:transition-all hover:after:w-full"
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
        </div>
      </div>
    </footer>
  )
}

export default Footer