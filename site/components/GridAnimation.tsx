'use client'

import { useEffect, useRef } from 'react'

const GridAnimation = () => {
  const canvasRef = useRef<HTMLCanvasElement>(null)

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return

    const ctx = canvas.getContext('2d')
    if (!ctx) return

    canvas.width = window.innerWidth
    canvas.height = window.innerHeight

    // Grid configuration - rectangular cells like a spreadsheet
    const gridSizeX = 80  // Wider cells
    const gridSizeY = 35  // Shorter cells
    const perspective = 800
    let time = 0
    let mouseX = canvas.width / 2
    let mouseY = canvas.height / 2

    // Create grid points
    const createGrid = () => {
      const points = []
      const rows = Math.ceil(canvas.height / gridSizeY) + 2
      const cols = Math.ceil(canvas.width / gridSizeX) + 2

      for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
          points.push({
            x: j * gridSizeX - gridSizeX,
            y: i * gridSizeY - gridSizeY,
            z: 0,
            originalX: j * gridSizeX - gridSizeX,
            originalY: i * gridSizeY - gridSizeY,
          })
        }
      }
      return points
    }

    let gridPoints = createGrid()

    // Mouse tracking
    const handleMouseMove = (e: MouseEvent) => {
      mouseX = e.clientX
      mouseY = e.clientY
    }

    window.addEventListener('mousemove', handleMouseMove)

    // Animation loop
    const animate = () => {
      ctx.clearRect(0, 0, canvas.width, canvas.height)
      
      // Update grid points with wave effect
      gridPoints.forEach((point) => {
        const distX = point.originalX - mouseX
        const distY = point.originalY - mouseY
        const distance = Math.sqrt(distX * distX + distY * distY)
        const maxDistance = 200
        
        if (distance < maxDistance) {
          const influence = (1 - distance / maxDistance) * 30
          point.z = Math.sin(time * 0.003 + distance * 0.01) * influence
        } else {
          point.z = Math.sin(time * 0.002 + point.originalX * 0.005 + point.originalY * 0.005) * 10
        }
      })

      // Draw connections - slightly more visible
      ctx.strokeStyle = 'rgba(94, 124, 226, 0.1)'
      ctx.lineWidth = 1

      const cols = Math.ceil(canvas.width / gridSizeX) + 2

      gridPoints.forEach((point, index) => {
        // Calculate perspective projection
        const scale = perspective / (perspective + point.z)
        const projectedX = point.originalX * scale + (1 - scale) * canvas.width / 2
        const projectedY = point.originalY * scale + (1 - scale) * canvas.height / 2

        // Draw horizontal lines
        if ((index + 1) % cols !== 0) {
          const nextPoint = gridPoints[index + 1]
          const nextScale = perspective / (perspective + nextPoint.z)
          const nextX = nextPoint.originalX * nextScale + (1 - nextScale) * canvas.width / 2
          const nextY = nextPoint.originalY * nextScale + (1 - nextScale) * canvas.height / 2

          ctx.beginPath()
          ctx.moveTo(projectedX, projectedY)
          ctx.lineTo(nextX, nextY)
          ctx.stroke()
        }

        // Draw vertical lines
        if (index + cols < gridPoints.length) {
          const nextPoint = gridPoints[index + cols]
          const nextScale = perspective / (perspective + nextPoint.z)
          const nextX = nextPoint.originalX * nextScale + (1 - nextScale) * canvas.width / 2
          const nextY = nextPoint.originalY * nextScale + (1 - nextScale) * canvas.height / 2

          ctx.beginPath()
          ctx.moveTo(projectedX, projectedY)
          ctx.lineTo(nextX, nextY)
          ctx.stroke()
        }

        // Draw nodes at intersections
        const nodeSize = 2 * scale
        const opacity = 0.2 * scale
        ctx.fillStyle = `rgba(121, 208, 208, ${opacity})`
        ctx.beginPath()
        ctx.arc(projectedX, projectedY, nodeSize, 0, Math.PI * 2)
        ctx.fill()
      })

      time++
      requestAnimationFrame(animate)
    }

    // Handle resize
    const handleResize = () => {
      canvas.width = window.innerWidth
      canvas.height = window.innerHeight
      gridPoints = createGrid()
    }

    window.addEventListener('resize', handleResize)
    animate()

    return () => {
      window.removeEventListener('mousemove', handleMouseMove)
      window.removeEventListener('resize', handleResize)
    }
  }, [])

  return (
    <canvas
      ref={canvasRef}
      className="fixed inset-0 pointer-events-none"
      style={{ zIndex: 0 }}
    />
  )
}

export default GridAnimation