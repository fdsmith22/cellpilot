# CellPilot Quick Reference Card

## 🚀 Most Common Tasks

### Making Code Changes
```bash
# 1. Edit main library
cd apps-script/main-library
# Make your changes in src/

# 2. Push to Apps Script
cd ..
npm run push:library

# 3. Test & Commit
cd ..
git add -A && git commit -m "Description" && git push
```

### Adding New Function
**Remember: Update ALL 4 places!**

1. **Main Library** (`apps-script/main-library/src/`)
   - Implement function
   
2. **Library Export** (`apps-script/main-library/src/Library.js`)
   ```javascript
   var myFunction = myFunction || function() { return myFunction(); };
   ```

3. **Beta Installer** (`apps-script/beta-installer/Code.gs`)
   ```javascript
   function myFunction() { return CellPilot.myFunction(); }
   ```

4. **Test Proxy** (`test-sheet-proxy.js`)
   ```javascript
   function myFunction() { return CellPilot.myFunction(); }
   ```

### Deploying Updates
```bash
# After making changes and testing:

# 1. Deploy new beta installer (if proxies changed)
cd apps-script/beta-installer
npx clasp deploy -d "v1.X - Description"
# Copy deployment ID!

# 2. Update website with new URL
cd ../../site
nano components/BetaAccessCard.tsx
# Update the URL

# 3. Push everything
cd ..
git add -A && git commit -m "Deploy v1.X" && git push
```

### Testing Your Changes
```bash
# Quick test in Google Sheets:
1. Open https://sheets.new
2. Extensions > Apps Script
3. Delete all code
4. Add Library: 1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O
   - Version: HEAD (for testing)
   - Identifier: CellPilot
5. Copy test-sheet-proxy.js content
6. Save & refresh sheet
```

## 📂 Directory Cheat Sheet

```
Working on Apps Script?  → cd apps-script/main-library
Working on Installer?    → cd apps-script/beta-installer  
Working on Website?      → cd site
Pushing to Apps Script?  → cd apps-script && npm run push:library
Testing deployment?      → Open Google Sheets
```

## 🔧 Command Shortcuts

### From `/apps-script` folder:
```bash
npm run push:library     # Push main code
npm run push:installer   # Push installer
npm run pull:library     # Pull main code
npm run pull:installer   # Pull installer
```

### From `/site` folder:
```bash
npm run dev             # Start local server
npx supabase status     # Check database
npx supabase db push    # Push migrations
```

### From root `/`:
```bash
# Quick commit & push
git add -A && git commit -m "msg" && git push

# Check what changed
git status
```

## ⚠️ Don't Forget!

### When Adding Functions:
✅ Main implementation  
✅ Library export  
✅ Beta installer proxy  
✅ Test proxy file  

### When Deploying:
✅ Test first  
✅ Deploy new version  
✅ Update website URL  
✅ Commit everything  

### File Locations:
- **Main Code**: `apps-script/main-library/src/`
- **Templates**: `apps-script/main-library/*Template.html`
- **Installer**: `apps-script/beta-installer/Code.gs`
- **Test Proxy**: `test-sheet-proxy.js` (root)
- **Website URL**: `site/components/BetaAccessCard.tsx`

## 🔍 Current Deployment Info

**Library Script ID**: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`  
**Current Library Version**: 11  
**Current Installer**: v1.8  
**Beta Installer URL**: Check `site/components/BetaAccessCard.tsx`

## 🆘 Quick Fixes

**"Function not found"**: Check all 4 places have the function  
**"Not authorized"**: Run initializeCellPilot in Apps Script  
**"Changes not showing"**: Hard refresh (Ctrl+Shift+R)  
**"Library not defined"**: Check identifier is "CellPilot"  

---
*Keep this handy while developing!*