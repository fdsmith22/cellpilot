import Hero from '@/components/Hero'
import Features from '@/components/Features'
import Pricing from '@/components/Pricing'
import CTA from '@/components/CTA'
import GridAnimation from '@/components/GridAnimation'

export default function Home() {
  return (
    <main className="min-h-screen relative">
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