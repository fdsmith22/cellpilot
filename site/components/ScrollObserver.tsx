'use client'

import { useEffect, useRef } from 'react'

interface ScrollObserverProps {
  children: React.ReactNode
  className?: string
  threshold?: number
  rootMargin?: string
}

const ScrollObserver: React.FC<ScrollObserverProps> = ({
  children,
  className = '',
  threshold = 0.1,
  rootMargin = '0px',
}) => {
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add('in-view')
          }
        })
      },
      {
        threshold,
        rootMargin,
      }
    )

    const element = ref.current
    if (element) {
      observer.observe(element)
    }

    return () => {
      if (element) {
        observer.unobserve(element)
      }
    }
  }, [threshold, rootMargin])

  return (
    <div ref={ref} className={`scroll-observer ${className}`}>
      {children}
    </div>
  )
}

export default ScrollObserver