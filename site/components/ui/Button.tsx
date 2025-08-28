import React from 'react'
import Link from 'next/link'
import { cn } from '@/lib/utils'

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: 'primary' | 'secondary' | 'ghost' | 'cell'
  size?: 'sm' | 'md' | 'lg'
  href?: string
  external?: boolean
  children: React.ReactNode
  className?: string
  cellRef?: string // For cell-style buttons
}

const Button = React.forwardRef<HTMLButtonElement | HTMLAnchorElement, ButtonProps>(
  ({ variant = 'primary', size = 'md', href, external, children, className, cellRef, ...props }, ref) => {
    const baseStyles = 'inline-flex items-center justify-center font-semibold transition-all duration-300 rounded-lg'
    
    const variants = {
      primary: 'bg-primary-600 text-white hover:bg-primary-700 hover:shadow-xl hover:scale-[1.02] active:scale-[0.98]',
      secondary: 'glass-card text-neutral-900 hover:bg-gray-200',
      ghost: 'text-neutral-600 hover:text-primary-600 hover:bg-neutral-100',
      cell: 'bg-white/80 backdrop-blur-md border-2 border-primary-300/50 shadow-lg hover:shadow-xl hover:scale-105 group'
    }
    
    const sizes = {
      sm: 'px-3 py-1.5 text-sm',
      md: 'px-4 py-2 text-base',
      lg: 'px-6 py-3 text-lg'
    }
    
    const classes = cn(baseStyles, variants[variant], sizes[size], className)
    
    // If it's a link
    if (href) {
      if (external) {
        return (
          <a
            href={href}
            target="_blank"
            rel="noopener noreferrer"
            className={classes}
            ref={ref as React.Ref<HTMLAnchorElement>}
          >
            {cellRef && (
              <div className="absolute -top-3 -left-3 text-xs text-neutral-500 font-mono">{cellRef}</div>
            )}
            {children}
          </a>
        )
      }
      return (
        <Link href={href} className={classes} ref={ref as React.Ref<HTMLAnchorElement>}>
          {cellRef && (
            <div className="absolute -top-3 -left-3 text-xs text-neutral-500 font-mono">{cellRef}</div>
          )}
          {children}
        </Link>
      )
    }
    
    // Regular button
    return (
      <button
        className={classes}
        ref={ref as React.Ref<HTMLButtonElement>}
        {...props}
      >
        {cellRef && (
          <div className="absolute -top-3 -left-3 text-xs text-neutral-500 font-mono">{cellRef}</div>
        )}
        {children}
      </button>
    )
  }
)

Button.displayName = 'Button'

export default Button