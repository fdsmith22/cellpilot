import Hero from '@/components/Hero'
import Features from '@/components/Features'
import Pricing from '@/components/Pricing'
import CTA from '@/components/CTA'
import GridAnimation from '@/components/GridAnimation'

export default function Home() {
  return (
    <main className="min-h-screen relative">
      {/* Fixed background grid that spans all sections */}
      <div className="fixed inset-0 grid-pattern opacity-20 z-0"></div>
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