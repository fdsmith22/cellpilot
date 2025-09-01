# CellPilot Project Context
*This file should be included in all development prompts to ensure proper workflow*

## 🎯 Project Overview
CellPilot is a Google Sheets add-on with a Next.js web dashboard for user management.

## 📁 Critical File Locations

### When Adding/Modifying Functions - UPDATE ALL 4 PLACES:
1. **Implementation**: `apps-script/main-library/src/[Feature].js`
2. **Export**: `apps-script/main-library/src/Library.js`
3. **Beta Proxy**: `apps-script/beta-installer/Code.gs`
4. **Test Proxy**: `test-sheet-proxy.js`

### Key Configuration Files:
- **Website Installer URL**: `site/components/BetaAccessCard.tsx`
- **Library Script ID**: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
- **Current Versions**: Library v11, Beta Installer v1.8

## 🔄 Required Development Workflow

### For ANY Code Changes:
```bash
# 1. Make changes in main library
cd apps-script/main-library
# Edit src files

# 2. If new function, add export
# Edit src/Library.js
var functionName = functionName || function() { return functionName(); };

# 3. Add proxy to beta installer
cd ../beta-installer
# Edit Code.gs
function functionName() { return CellPilot.functionName(); }

# 4. Add to test proxy
cd ../..
# Edit test-sheet-proxy.js (same proxy function)

# 5. Push changes
cd apps-script
npm run push:library
npm run push:installer

# 6. Test in Google Sheets

# 7. If proxies changed, deploy new version
cd beta-installer
npx clasp deploy -d "v1.X - Description"
# COPY THE DEPLOYMENT ID!

# 8. Update website with new deployment
cd ../../site
# Edit components/BetaAccessCard.tsx with new URL

# 9. Commit everything together
cd ..
git add -A
git commit -m "Feature: Description with deployment info"
git push
```

## ⚠️ CRITICAL RULES

### NEVER:
- ❌ Add functions to only one location
- ❌ Deploy without testing
- ❌ Forget to update BetaAccessCard.tsx after new deployment
- ❌ Mix up clasp projects (always check your directory!)
- ❌ Commit partial changes

### ALWAYS:
- ✅ Update all 4 locations for user-facing functions
- ✅ Run `./scripts/check-sync.sh` before deploying
- ✅ Test in actual Google Sheets
- ✅ Deploy new versions when proxies change
- ✅ Update website URL after deployment
- ✅ Commit all related changes together

## 📍 Directory Navigation

```bash
# Working locations
cd apps-script/main-library    # Main code
cd apps-script/beta-installer  # Installer
cd site                        # Website
cd /home/freddy/cellpilot      # Root

# Pushing code
cd apps-script && npm run push:library
cd apps-script && npm run push:installer
```

## 🛠️ Tool Commands

### Apps Script (from `/apps-script`):
```bash
npm run push:library    # Push main library
npm run push:installer  # Push beta installer
npm run pull:library    # Pull changes
npm run pull:installer  # Pull installer
```

### Deployment (from specific folders):
```bash
npx clasp deploy -d "Version description"
npx clasp deployments   # View all deployments
```

### Database (from `/site`):
```bash
npx supabase db push    # Push migrations
npx supabase status     # Check status
```

### Verification:
```bash
./scripts/check-sync.sh # Check function synchronization
```

## 📋 Pre-Deployment Checklist

Before ANY deployment:
- [ ] Functions exist in all 4 locations
- [ ] Ran check-sync.sh successfully
- [ ] Tested in Google Sheets
- [ ] Deployed new versions if needed
- [ ] Updated BetaAccessCard.tsx
- [ ] Ready to commit all changes

## 🔍 Current State
- **Library Version**: 11 (HEAD for testing)
- **Beta Installer**: v1.8
- **Proxy Functions**: 200+ synchronized
- **Database**: Supabase with RLS policies

## 📝 Commit Message Format
```
Type: Brief description

- Implementation details
- Proxy functions added/modified
- Deployment version if applicable
- Website URL updated if needed
```

## ⚠️ IMPORTANT REMINDERS

1. **Function Naming**: Must match EXACTLY across all files (case-sensitive)
2. **Testing**: Always use HEAD version for testing, numbered version for production
3. **Deployment**: New deployment needed when proxies change, not for internal logic
4. **Website**: Vercel auto-deploys from GitHub push
5. **Database**: Test migrations locally first with `npx supabase db reset`

---
**Include this context when asking for development help to ensure proper workflow is followed**