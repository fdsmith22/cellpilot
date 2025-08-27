'use client'

import Link from 'next/link'

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
      { name: 'Blog', href: '/blog' },
      { name: 'Privacy', href: '/privacy' },
      { name: 'Terms', href: '/terms' },
    ],
    Connect: [
      { name: 'GitHub', href: 'https://github.com' },
      { name: 'Twitter', href: 'https://twitter.com' },
      { name: 'LinkedIn', href: 'https://linkedin.com' },
      { name: 'YouTube', href: 'https://youtube.com' },
    ],
  }

  return (
    <footer className="relative overflow-hidden">
      {/* Subtle animated background */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute -bottom-32 -left-32 w-64 h-64 bg-pastel-lavender/10 rounded-full blur-3xl"></div>
        <div className="absolute -bottom-32 -right-32 w-64 h-64 bg-pastel-mint/10 rounded-full blur-3xl"></div>
      </div>
      
      {/* Seamless transition from sections above */}
      <div className="absolute inset-0 bg-gradient-to-b from-neutral-50/90 via-white/95 to-primary-50/5"></div>
      
      {/* Grid pattern overlay */}
      <div className="absolute inset-0 grid-pattern opacity-15"></div>
      
      <div className="relative z-10">
        <div className="container-wrapper py-16 lg:py-20">
          {/* Main Footer Content */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-8 mb-12">
            {/* Brand Column */}
            <div className="col-span-2 md:col-span-1">
              <Link href="/" className="inline-block mb-6">
                <img 
                  src="/logo/combined/horizontal-standard-200x60.svg" 
                  alt="CellPilot" 
                  className="h-12 w-auto"
                />
              </Link>
              <p className="text-sm text-neutral-600 mb-6 leading-relaxed">
                Transform your spreadsheet workflow with intelligent automation.
              </p>
              {/* Social Links - Mobile */}
              <div className="flex space-x-3 md:hidden">
                {footerLinks.Connect.map((link) => (
                  <a
                    key={link.name}
                    href={link.href}
                    className="w-9 h-9 rounded-lg bg-neutral-100/80 hover:bg-pastel-sky/20 flex items-center justify-center text-neutral-500 hover:text-primary-600 transition-all hover:scale-110"
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    <span className="sr-only">{link.name}</span>
                    <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 24 24">
                      <path d="M12 2C6.477 2 2 6.477 2 12c0 4.42 2.865 8.17 6.839 9.49.5.092.682-.217.682-.482 0-.237-.008-.866-.013-1.7-2.782.603-3.369-1.34-3.369-1.34-.454-1.156-1.11-1.463-1.11-1.463-.908-.62.069-.608.069-.608 1.003.07 1.531 1.03 1.531 1.03.892 1.529 2.341 1.087 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.11-4.555-4.943 0-1.091.39-1.984 1.029-2.683-.103-.253-.446-1.27.098-2.647 0 0 .84-.269 2.75 1.025A9.578 9.578 0 0112 6.836c.85.004 1.705.114 2.504.336 1.909-1.294 2.747-1.025 2.747-1.025.546 1.377.203 2.394.1 2.647.64.699 1.028 1.592 1.028 2.683 0 3.842-2.339 4.687-4.566 4.935.359.309.678.919.678 1.852 0 1.336-.012 2.415-.012 2.743 0 .267.18.578.688.48C19.138 20.167 22 16.418 22 12c0-5.523-4.477-10-10-10z" />
                    </svg>
                  </a>
                ))}
              </div>
            </div>

          {/* Links Columns */}
          {Object.entries(footerLinks).slice(0, 3).map(([category, links]) => (
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

          {/* Social Links - Desktop */}
          <div className="hidden md:block">
            <h3 className="text-sm font-semibold text-neutral-800 tracking-wide mb-4">Connect</h3>
            <div className="flex space-x-3">
              {footerLinks.Connect.map((link) => (
                <a
                  key={link.name}
                  href={link.href}
                  className="w-9 h-9 rounded-lg bg-neutral-100/80 hover:bg-pastel-sky/20 flex items-center justify-center text-neutral-500 hover:text-primary-600 transition-all hover:scale-110"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  <span className="sr-only">{link.name}</span>
                  <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 24 24">
                    <path d="M12 2C6.477 2 2 6.477 2 12c0 4.42 2.865 8.17 6.839 9.49.5.092.682-.217.682-.482 0-.237-.008-.866-.013-1.7-2.782.603-3.369-1.34-3.369-1.34-.454-1.156-1.11-1.463-1.11-1.463-.908-.62.069-.608.069-.608 1.003.07 1.531 1.03 1.531 1.03.892 1.529 2.341 1.087 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.11-4.555-4.943 0-1.091.39-1.984 1.029-2.683-.103-.253-.446-1.27.098-2.647 0 0 .84-.269 2.75 1.025A9.578 9.578 0 0112 6.836c.85.004 1.705.114 2.504.336 1.909-1.294 2.747-1.025 2.747-1.025.546 1.377.203 2.394.1 2.647.64.699 1.028 1.592 1.028 2.683 0 3.842-2.339 4.687-4.566 4.935.359.309.678.919.678 1.852 0 1.336-.012 2.415-.012 2.743 0 .267.18.578.688.48C19.138 20.167 22 16.418 22 12c0-5.523-4.477-10-10-10z" />
                  </svg>
                </a>
              ))}
            </div>
          </div>
        </div>

          {/* Bottom Bar */}
          <div className="pt-8 mt-8 border-t border-neutral-200/50">
            <div className="flex flex-col md:flex-row justify-between items-center space-y-4 md:space-y-0">
              <p className="text-sm text-neutral-500">
                © {currentYear} CellPilot. All rights reserved.
              </p>
              <div className="flex items-center space-x-6">
                <Link href="/privacy" className="text-sm text-neutral-500 hover:text-primary-600 transition-all hover:translate-y-[-1px]">
                  Privacy Policy
                </Link>
                <Link href="/terms" className="text-sm text-neutral-500 hover:text-primary-600 transition-all hover:translate-y-[-1px]">
                  Terms of Service
                </Link>
                <Link href="/cookies" className="text-sm text-neutral-500 hover:text-primary-600 transition-all hover:translate-y-[-1px]">
                  Cookie Policy
                </Link>
              </div>
            </div>
          </div>
        </div>
      </div>
    </footer>
  )
}

export default Footer