# CellPilot AI Assistant Context

## Project Overview
CellPilot is a Google Sheets add-on with 40K+ lines of Apps Script code. This is a production system with active users - exercise caution with all changes.

## CRITICAL: Before Making Any Changes
1. **ALWAYS** read MASTER_DOCUMENTATION.md first
2. **NEVER** create new .md files - update MASTER_DOCUMENTATION.md instead
3. **TEST** all changes using the documented testing approaches
4. **MAINTAIN** backwards compatibility - existing users depend on this

## Project Structure
```
cellpilot/
├── apps-script/           # Google Apps Script code
│   ├── main-library/      # Core add-on (36 modules)
│   │   ├── src/          # JavaScript source files
│   │   └── *.html        # UI templates
│   └── beta-installer/    # Installation helper
├── site/                  # Next.js dashboard
├── scripts/               # Development utilities
└── MASTER_DOCUMENTATION.md # Single source of truth
```

## Key Files to Monitor
- `apps-script/main-library/src/Code.js` - Main entry (4K lines)
- `apps-script/main-library/src/CellM8.js` - AI engine (5K lines)  
- `apps-script/main-library/src/Library.js` - Exports
- `apps-script/main-library/appsscript.json` - Manifest
- `MASTER_DOCUMENTATION.md` - All documentation

## Development Workflow

### Making Code Changes
```bash
# 1. Check documentation
cat MASTER_DOCUMENTATION.md | grep -A10 "section_name"

# 2. Make changes following patterns
# 3. Test without deployment
cd apps-script && npm run push:library
# Then test in Apps Script editor

# 4. Deploy if tests pass
cd apps-script && npm run deploy:library
```

### CellM8 Development (CRITICAL)
CellM8 requires functions in **4 locations**:
1. `CellM8.js` - Implementation
2. `Library.js` - Export
3. `Code.js` - Menu integration
4. `CellM8Template.html` - UI proxy

**Required Structure:**
```javascript
const CellM8 = {
  SlideGenerator: {
    functionName: function() { /* ... */ }
  }
};
```

### Testing Approaches
1. **Direct Testing** (Preferred):
   ```javascript
   function testCellM8Directly() {
     const result = CellM8.SlideGenerator.analyzeDataIntelligently(data);
     console.log(result);
   }
   ```

2. **Sidebar Testing**:
   - Push changes: `npm run push:library`
   - Open test sheet → CellPilot menu → Open CellM8
   - Check logs in Apps Script editor

## Common Commands
```bash
# Deployment
cd apps-script && npm run push:library    # Push changes
cd apps-script && npm run deploy:library  # Deploy version
cd apps-script && npm run logs:library    # View logs

# Status & Analysis
./scripts/status.sh              # Project health dashboard
codebase-digest                  # Generate AI context
npx dependency-cruiser --validate apps-script/main-library/src

# Documentation
npx jsdoc apps-script/main-library/src -r -d docs/api
```

## Current Issues & Blockers
1. **Logo Hosting** - Sidebar logo returns 404
2. **Google Cloud Project** - Not configured
3. **OAuth Verification** - Pending for marketplace

## Rules for AI Assistants
1. **Read First**: Always check MASTER_DOCUMENTATION.md
2. **Test Everything**: Use documented testing approaches
3. **No New .md Files**: Update MASTER_DOCUMENTATION.md only
4. **Preserve Compatibility**: Don't break existing functionality
5. **Follow Patterns**: Match existing code style and structure
6. **Document Changes**: Update MASTER_DOCUMENTATION.md when making significant changes

## Marketplace Timeline
- Beta Launch: 1-2 weeks (pending logo fix)
- Marketplace: 4-5 weeks (pending OAuth verification)

## Support Resources
- Script ID: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
- Dashboard: https://cellpilot.io
- Apps Script Console: https://script.google.com
- Issue Reports: https://github.com/anthropics/claude-code/issues

## Performance Targets
- Menu load: <2 seconds
- Sidebar init: <3 seconds  
- Data processing: <30 seconds
- Chart generation: <5 seconds

## Testing Checklist
Before any deployment:
- [ ] All functions accessible through Library.js
- [ ] Menu items trigger correct functions
- [ ] Sidebar loads without errors
- [ ] No permission errors
- [ ] Charts generate with unique IDs
- [ ] Performance within targets
- [ ] Backwards compatibility maintained

---
*Always refer to MASTER_DOCUMENTATION.md for detailed information*