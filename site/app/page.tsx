'use client'

import dynamic from 'next/dynamic'
import Hero from '@/components/Hero'
import Features from '@/components/Features'
import Pricing from '@/components/Pricing'
import CTA from '@/components/CTA'
import Footer from '@/components/Footer'
import ViewportProvider from '@/components/ViewportProvider'

// Lazy load the GridAnimation component for better performance
const GridAnimation = dynamic(() => import('@/components/GridAnimation'), {
  ssr: false, // Disable SSR for canvas-based animation
  loading: () => null // No loading indicator needed as it's a background element
})

export default function Home() {
  return (
    <ViewportProvider>
      <main className="relative">
        <GridAnimation />
        <div className="relative z-elevated">
          <Hero />
          <Features />
          <Pricing />
          <CTA />
          <Footer />
        </div>
      </main>
    </ViewportProvider>
  )
}