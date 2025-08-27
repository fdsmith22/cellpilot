# CellPilot Brand Assets

## Logo Structure
This directory contains all CellPilot logo variations and brand assets.

### Directory Structure

```
/logo
├── icons/              # Icon-only versions (no text)
│   ├── favicon-16x16.svg     # Browser tab favicon
│   ├── icon-32x32.svg        # Small icon for apps
│   ├── icon-64x64.svg        # Medium icon for desktop
│   └── icon-64x64-dark.svg   # Dark theme variant
│
├── wordmarks/          # Text-only versions (no icon)
│   ├── text-standard-24px.svg  # Standard text size
│   ├── text-large-36px.svg     # Large presentations
│   ├── text-dark-24px.svg      # Dark background version
│   └── text-small-16px.svg     # Small/footer usage
│
├── combined/           # Icon + Text combinations
│   ├── horizontal-standard-200x60.svg  # Standard header
│   ├── horizontal-large-300x90.svg     # Large displays
│   ├── horizontal-compact-160x50.svg   # Compact/mobile
│   └── horizontal-dark-200x60.svg      # Dark theme
│
└── stacked/           # Vertical layout
    └── vertical-120x90.svg  # Square format, social media

```

## Brand Colors

- **Primary Blue**: #2563eb
- **Secondary Blue**: #3b82f6  
- **Accent Blue**: #60a5fa
- **Light Blue**: #93c5fd
- **Text Dark**: #1f2937
- **Text Light**: #ffffff

## Usage Guidelines

### When to Use Each Logo

**Icons Only**
- Favicon (browser tabs)
- App icons
- Small UI elements where text won't be legible

**Wordmarks Only**
- When icon is displayed separately
- Text-heavy contexts
- Email signatures

**Combined Horizontal**
- Website headers
- Business cards
- Documentation headers
- Marketing materials

**Stacked Vertical**
- Social media profile images
- Square format requirements
- App splash screens

### Spacing Requirements

- Minimum clear space around logo: 0.5x the height of the icon
- Never stretch or distort logos
- Maintain aspect ratios

### Background Usage

- Use standard versions on light backgrounds (#ffffff to #f3f4f6)
- Use dark versions on dark backgrounds (#1f2937 to #000000)
- Ensure sufficient contrast for accessibility

## File Formats

All logos are provided in SVG format for:
- Perfect scaling at any size
- Small file sizes
- Web optimization
- Easy color customization if needed

To convert to PNG/JPG:
1. Use online converters like CloudConvert
2. Export from design tools (Figma, Illustrator)
3. Recommended sizes for PNG exports:
   - Favicon: 16x16, 32x32
   - App icons: 192x192, 512x512
   - Social media: 1200x1200

## Implementation Examples

### HTML Favicon
```html
<link rel="icon" type="image/svg+xml" href="/logo/icons/favicon-16x16.svg">
<link rel="apple-touch-icon" href="/apple-touch-icon.png">
```

### React Component
```jsx
import logo from '/logo/combined/horizontal-standard-200x60.svg'

<img src={logo} alt="CellPilot" className="h-12" />
```

### CSS Background
```css
.logo {
  background-image: url('/logo/icons/icon-64x64.svg');
  background-size: contain;
  background-repeat: no-repeat;
}
```

## Contact

For brand usage questions or additional format requests, contact: branding@cellpilot.app