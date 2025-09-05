# CellPilot Master Documentation
*Last Updated: 2025-09-05*
*This document consolidates all project documentation to prevent cross-contamination and maintain single source of truth*

---

## 🎯 Project Overview

CellPilot is a comprehensive Google Sheets add-on with 40K+ lines of Apps Script code, featuring advanced data manipulation, AI-powered analysis (CellM8), and professional presentation generation capabilities. The project includes a Next.js dashboard for user management and Supabase backend infrastructure.

### Technology Stack
- **Primary**: Google Apps Script (V8 Runtime) - 36 interdependent modules
- **Web**: Next.js 14, TypeScript, Tailwind CSS
- **Backend**: Supabase (Auth, Database)
- **Deployment**: Google CLASP, Vercel
- **Payments**: Stripe integration

### Project Structure
```
cellpilot/
├── apps-script/              # Google Apps Script Projects
│   ├── main-library/         # Core CellPilot add-on (40K+ lines)
│   │   ├── src/             # 30+ JavaScript modules
│   │   └── .clasp.json      # Deployment configuration
│   ├── beta-installer/      # Installation web app
│   └── package.json         # CLASP deployment scripts
├── site/                    # Next.js Web Application
├── supabase/               # Database configuration
└── docs/                   # Documentation
```

---

## 📋 MARKETPLACE REQUIREMENTS & ACTION PLAN

### Current Status
- **Beta Readiness**: 75% complete
- **Marketplace Readiness**: 60% complete
- **Estimated Beta Launch**: 1-2 weeks
- **Estimated Marketplace**: 4-5 weeks

### Critical Blockers (Must Fix)
1. **Logo Hosting Issue** - Sidebar logo returns 404
2. **Google Cloud Project Setup** - Missing project configuration
3. **API Endpoints** - Missing post-installation tracking endpoints

### Pre-Submission Requirements
- ✅ Functional add-on core
- ✅ OAuth scopes defined
- ✅ Legal documents (Privacy Policy, Terms)
- ❌ Google Cloud Project with billing
- ❌ OAuth consent screen configuration
- ❌ Domain verification (for restricted scopes)
- ❌ Screenshots (1280x800 or 640x400)
- ❌ User documentation

### OAuth Verification Timeline
- **Unverified Apps**: Immediate (100 user limit)
- **Verified Apps**: 2-3 business days
- **Restricted Scopes**: Several weeks review

### App Listing Requirements
- **Logo**: 128x128 PNG (under 100KB)
- **Screenshots**: At least 2, max 5 (PNG/JPG)
- **Banner**: 440x280 optional promotional graphic
- **Description**: 1000 character limit
- **Support Links**: Privacy policy, terms, support URL

### 6-Week Implementation Timeline
**Week 1-2: Preparation**
- Fix logo hosting issue
- Set up Google Cloud Project
- Create missing API endpoints
- Prepare screenshots

**Week 3-4: Documentation & Testing**
- Complete user documentation
- Run comprehensive testing
- Fix identified bugs
- Prepare submission materials

**Week 5: Submission**
- Submit for OAuth verification
- Submit marketplace listing
- Internal beta testing

**Week 6: Launch**
- Address review feedback
- Marketing preparation
- Beta user onboarding

---

## 🔧 TECHNICAL INFRASTRUCTURE

### Supabase Database Connection
**Project Reference**: `grhcxavwagumzceauthr`

```bash
# Direct Connection
Host: aws-0-us-west-1.pooler.supabase.com
Port: 5432
Database: postgres
User: postgres.grhcxavwagumzceauthr

# CLI Commands
npx supabase login
npx supabase link --project-ref grhcxavwagumzceauthr
npx supabase db pull
```

### Logo Setup Solution
**Quick Fix Using Vercel Deployment**:
1. Upload logo to: `site/public/images/cellpilot-logo-48.png`
2. Deploy to Vercel: `cd site && vercel --prod`
3. Update Code.js: `logoUrl: 'https://cellpilot.io/images/cellpilot-logo-48.png'`
4. Push changes: `npm run push:library`

**Requirements**:
- 48x48 pixels
- HTTPS URL required
- Under 100KB file size
- PNG or SVG format

---

## 🚀 CELLM8 INTEGRATION ARCHITECTURE

### Critical Multi-Layer Proxy Structure
CellM8 requires functions in **4 locations** to work:

1. **Library.js** - Main export point
2. **Code.js** - Menu integration
3. **CellM8.js** - Core implementation  
4. **CellM8Template.html** - UI interface

### Required Scope Structure
```javascript
// CRITICAL: Must maintain this exact structure
const CellM8 = {
  SlideGenerator: {
    analyzeDataIntelligently: function() { /* ... */ },
    createPresentationFromAnalysis: function() { /* ... */ },
    createDashboardLayout: function() { /* ... */ }
    // ... 34 total functions
  }
};
```

### Development Workflow
1. **Always check** existing patterns in CellM8.js
2. **Add function** to CellM8.SlideGenerator object
3. **Export** through Library.js
4. **Test** without deployment using test approaches
5. **Verify** all 4 proxy layers work
6. **Deploy** only after successful testing

### Common Issues & Solutions
- **TypeError**: Missing scope structure → Check CellM8.SlideGenerator exists
- **undefined function**: Not exported in Library.js → Add to exports
- **Template errors**: Proxy function missing → Add to template HTML
- **Chart duplication**: Missing unique IDs → Use timestamp identifiers

---

## 🧪 TESTING APPROACHES

### Option 1: Direct Testing (Recommended)
```javascript
// In Apps Script Editor after clasp push
function testCellM8Directly() {
  const testData = [
    ['Product', 'Sales', 'Growth'],
    ['A', 100, 10],
    ['B', 200, 20]
  ];
  
  const result = CellM8.SlideGenerator.analyzeDataIntelligently(testData);
  console.log('Analysis result:', result);
}
```

### Option 2: HTTP Testing Endpoint
```javascript
// Add to Code.js
function doGet(e) {
  if (e.parameter.test === 'cellm8') {
    return ContentService.createTextOutput(
      JSON.stringify(testCellM8Functions())
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
```

### Option 3: Test Through Sidebar
1. Push changes: `npm run push:library`
2. Open test spreadsheet
3. CellPilot menu → Open CellM8
4. Test functions through UI
5. Check logs in Apps Script editor

### Testing Checklist
- [ ] All functions accessible through Library.js
- [ ] Menu items trigger correct functions
- [ ] Sidebar loads without errors
- [ ] API calls return expected data
- [ ] No permission errors
- [ ] Charts generate with unique IDs
- [ ] Templates process correctly
- [ ] Scope structure maintained
- [ ] Error handling works
- [ ] Logging provides useful debug info
- [ ] Performance acceptable (<30s timeout)

---

## 📊 CELLM8 PHASE 2 IMPLEMENTATION PLAN

### Enhancement Roadmap (4 Weeks)

**Phase 2.1: Professional Theme System (Week 1)**
- Native Google Slides theming
- Color scheme management
- Font and style controls
- Brand consistency tools

**Phase 2.2: Advanced Chart Customization (Week 1)**
- Granular chart controls
- Professional styling options
- Custom color palettes
- Advanced formatting

**Phase 2.3: Animation Controls (Week 2)**
- Slide transitions
- Object animations
- Timing controls
- Professional effects

**Phase 2.4: Chart Templates (Week 2)**
- Business scenario presets
- Industry-specific charts
- Quick styling options
- Template library

**Phase 2.5: Dashboard Layouts (Week 3)**
- Multiple layout options
- Responsive designs
- Grid-based positioning
- Layout templates

**Phase 2.6: Smart Content (Week 4)**
- AI-powered insights
- Narrative generation
- Key finding highlights
- Executive summaries

### Success Metrics
- Generate professional presentations in <2 minutes
- 20+ chart customization options
- 10+ animation presets
- 15+ dashboard layouts
- 95% user satisfaction score

---

## 🔐 DEVELOPMENT GUIDELINES

### Golden Rules
1. **ALWAYS** reference this master documentation
2. **NEVER** create new .md files - update this one
3. **TEST** in sandbox before production
4. **MAINTAIN** backwards compatibility
5. **DOCUMENT** all API changes here

### Code Modification Workflow
1. Check this documentation first
2. Review existing patterns in target file
3. Make changes following conventions
4. Test using recommended approaches
5. Update this documentation if needed
6. Deploy using CLASP commands

### Deployment Commands
```bash
# Apps Script
cd apps-script
npm run push:library    # Deploy main add-on
npm run push:installer  # Deploy installer
npm run pull:library    # Pull latest code

# Web Application
cd site
npm run dev            # Local development
npm run build          # Production build
vercel --prod          # Deploy to production
```

### Critical Files to Monitor
- `apps-script/main-library/src/Code.js` - Main entry point (4K lines)
- `apps-script/main-library/src/CellM8.js` - AI engine (5K lines)
- `apps-script/main-library/src/Library.js` - Export definitions
- `apps-script/main-library/appsscript.json` - Manifest configuration
- `site/app/page.tsx` - Web app landing page

---

## 🛠️ TROUBLESHOOTING GUIDE

### Common Issues

**Logo Not Showing in Sidebar**
- Solution: Deploy logo to https://cellpilot.io/images/
- Verify: URL returns image, not 404
- Update: Code.js logoUrl property

**CellM8 Functions Undefined**
- Check: CellM8.SlideGenerator scope exists
- Verify: Function exported in Library.js
- Confirm: Template has proxy function

**Permission Errors**
- Review: appsscript.json OAuth scopes
- Check: User authorization status
- Verify: API quotas not exceeded

**Chart Duplication**
- Add: Unique timestamp IDs to charts
- Check: Chart deletion before creation
- Verify: Slide clearing logic

**Deployment Failures**
- Check: .clasp.json configuration
- Verify: Script ID correct
- Confirm: File push order in config

### Debug Commands
```bash
# Check deployment status
clasp deployments

# View logs
clasp logs --watch

# Check project settings
clasp status

# Run specific function
clasp run testFunction
```

---

## 📈 PROJECT METRICS

### Current Statistics
- **Total Files**: 36 Apps Script modules
- **Code Volume**: 40,000+ lines
- **Functions**: 500+ exported functions
- **Templates**: 8 HTML templates
- **Documentation**: Consolidated into this file

### Performance Targets
- Menu load: <2 seconds
- Sidebar init: <3 seconds
- Data processing: <30 seconds
- Chart generation: <5 seconds
- API response: <1 second

---

## 🔄 VERSION HISTORY

### Recent Updates
- **2025-09-05**: Consolidated all documentation
- **2025-09-02**: CellM8 full integration completed
- **v50**: Comprehensive contrast improvements
- **v49**: Major presentation enhancements
- **v47**: Dashboard functions scope fix

### Git Workflow
```bash
# Check status
git status

# Create feature branch
git checkout -b feature/description

# Commit with semantic message
git add .
git commit -m "type: description"

# Types: feat, fix, docs, style, refactor, test, chore
```

---

## 📝 NOTES FOR AI ASSISTANTS

When working on this project:
1. **Start here** - Read relevant sections before coding
2. **Check patterns** - Follow existing code conventions
3. **Test first** - Use recommended testing approaches
4. **Update docs** - Keep this file current
5. **No new .md files** - All docs go here

### Context Commands
- "Check CellPilot documentation" → Refer to this file
- "Follow CellM8 integration pattern" → See CellM8 section
- "Test without deployment" → Use testing approaches
- "Fix marketplace blocker" → Check requirements section

---

*End of Master Documentation - All other .md files should be considered deprecated*