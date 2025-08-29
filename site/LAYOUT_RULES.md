# Layout Rules for CellPilot Site

## Critical: Preventing Header Overlap

### The Problem
The site has a fixed header that is 72px tall (--header-height). When content is centered vertically using `flex items-center justify-center` on full-height sections, it centers in the full viewport, causing content to slide under the header.

### The Solution

#### ❌ DON'T DO THIS (causes overlap):
```jsx
<section className="snap-section">
  <div className="h-full flex items-center justify-center">
    <!-- Content will be centered in full viewport and go under header -->
  </div>
</section>
```

#### ✅ DO THIS INSTEAD:
```jsx
<section className="snap-section">
  <div className="h-full flex flex-col justify-center pt-24 sm:pt-28 lg:pt-32">
    <!-- Content starts below header with proper padding -->
  </div>
</section>
```

### Key Rules:

1. **Never use `items-center` on hero section containers** - This centers content in the full viewport height
2. **Always use `flex-col justify-center` with `pt-*` padding** - This pushes content down from the top
3. **The first section gets special CSS treatment** - See `globals.css` line ~80
4. **Header height is defined as CSS variable** - `--header-height: 72px`

### CSS Classes:
- `.snap-section` - Full viewport height sections with scroll snap
- `.container-wrapper` - Standard container with responsive padding
- First snap-section automatically gets extra top padding via CSS

### Testing:
Always check that the topmost content in the hero section is fully visible and not hidden behind the header on:
- Desktop (1920x1080)
- Laptop (1440x900)
- Tablet (768x1024)
- Mobile (375x667)