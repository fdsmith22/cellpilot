'use client'

import dynamic from 'next/dynamic'
import Hero from '@/components/Hero'
import Features from '@/components/Features'
import Pricing from '@/components/Pricing'
import CTA from '@/components/CTA'

// Lazy load the GridAnimation component for better performance
const GridAnimation = dynamic(() => import('@/components/GridAnimation'), {
  ssr: false, // Disable SSR for canvas-based animation
  loading: () => null // No loading indicator needed as it's a background element
})

export default function Home() {
  return (
    <main className="relative">
      <GridAnimation />
      <div className="relative z-10">
        <Hero />
        <Features />
        <Pricing />
        <CTA />
      </div>
    </main>
  )
}