# CellPilot Beta Deployment Guide

## Overview
Deploy CellPilot for beta testing without exposing source code using Google Apps Script Library model.

## Architecture

```
┌─────────────────────────┐
│   Your Private Code     │
│  (Apps Script Library)  │
│   - All core logic      │
│   - Hidden from users   │
└──────────┬──────────────┘
           │
           ▼
┌─────────────────────────┐
│   Beta User's Script    │
│   (Thin Client)         │
│   - Only proxy calls    │
│   - No business logic   │
└─────────────────────────┘
```

## Step 1: Prepare Your Library (One-time setup)

### 1.1 In your main Apps Script project:

1. Open your main CellPilot Apps Script project
2. Click **Deploy** → **Manage deployments**
3. Click **Create** → **Library**
4. Add description: "CellPilot Beta Library"
5. Click **Deploy**
6. Copy the **Deployment ID** (looks like: `AKfycb...`)

### 1.2 Configure Library Access:

```javascript
// In your main Code.gs, ensure these functions are exposed:
function getLibraryVersion() {
  return '1.0.0-beta';
}

// All your existing functions remain as-is
// They're automatically available to library users
```

## Step 2: Create Beta Installer Script

Create a new Apps Script project for beta users to copy:

```javascript
/**
 * CellPilot Beta Installation
 * Version: 1.0.0-beta
 * 
 * This script connects to CellPilot services
 * No source code is exposed to users
 */

// Add CellPilot as a library
// 1. Click Libraries (+)
// 2. Add Script ID: YOUR_DEPLOYMENT_ID
// 3. Version: Use latest
// 4. Identifier: CellPilot

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  CellPilot.onOpen(e);
}

// Proxy all menu functions
function showCellPilotSidebar() { 
  return CellPilot.showCellPilotSidebar(); 
}

function tableize() { 
  return CellPilot.tableize(); 
}

function removeDuplicates() { 
  return CellPilot.removeDuplicates(); 
}

// Add all other menu functions as simple proxies...
```

## Step 3: Web App Deployment (For easier installation)

### 3.1 Create an Installer Web App:

```javascript
// installer-webapp.gs
function doGet(e) {
  // Check if user is authorized beta tester
  const userEmail = Session.getActiveUser().getEmail();
  
  if (!isBetaTester(userEmail)) {
    return HtmlService.createHtmlOutput('Not authorized for beta access');
  }
  
  // Return installation page
  return HtmlService.createTemplateFromFile('InstallPage')
    .evaluate()
    .setTitle('Install CellPilot Beta');
}

function isBetaTester(email) {
  // Check against your Supabase database
  // Or maintain a simple list
  const betaTesters = [
    'user1@example.com',
    'user2@example.com'
  ];
  return betaTesters.includes(email);
}

function installForUser() {
  // This creates the add-on for the user
  // Returns installation instructions
  return {
    scriptId: 'YOUR_LIBRARY_DEPLOYMENT_ID',
    instructions: [
      '1. Open a Google Sheet',
      '2. Click Extensions → Apps Script',
      '3. Copy the provided code',
      '4. Save and refresh your sheet'
    ]
  };
}
```

### 3.2 Create Installation HTML:

```html
<!-- InstallPage.html -->
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; }
    .container { max-width: 600px; margin: 0 auto; }
    .code-block { 
      background: #f5f5f5; 
      padding: 15px; 
      border-radius: 5px;
      font-family: monospace;
    }
    button { 
      background: #4285f4; 
      color: white; 
      padding: 10px 20px;
      border: none;
      border-radius: 5px;
      cursor: pointer;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>Install CellPilot Beta</h1>
    
    <div id="step1">
      <h2>Step 1: Authorize Installation</h2>
      <button onclick="startInstall()">Start Installation</button>
    </div>
    
    <div id="step2" style="display:none;">
      <h2>Step 2: Copy Installation Code</h2>
      <div class="code-block" id="code"></div>
      <button onclick="copyCode()">Copy Code</button>
    </div>
    
    <div id="step3" style="display:none;">
      <h2>Step 3: Install in Your Sheet</h2>
      <ol>
        <li>Open any Google Sheet</li>
        <li>Click Extensions → Apps Script</li>
        <li>Delete all existing code</li>
        <li>Paste the code you just copied</li>
        <li>Click Save (💾)</li>
        <li>Refresh your Google Sheet</li>
        <li>Find CellPilot in Extensions menu!</li>
      </ol>
    </div>
  </div>
  
  <script>
    function startInstall() {
      google.script.run
        .withSuccessHandler(showCode)
        .installForUser();
    }
    
    function showCode(result) {
      document.getElementById('step1').style.display = 'none';
      document.getElementById('step2').style.display = 'block';
      document.getElementById('step3').style.display = 'block';
      
      // Generate the installation code
      const code = generateInstallCode(result.scriptId);
      document.getElementById('code').textContent = code;
    }
    
    function generateInstallCode(scriptId) {
      return `// CellPilot Beta
// Auto-generated installation code

function onInstall(e) { onOpen(e); }
function onOpen(e) { 
  CellPilot.onOpen(e); 
}

// Menu proxies
function showCellPilotSidebar() { return CellPilot.showCellPilotSidebar(); }
function tableize() { return CellPilot.tableize(); }
function removeDuplicates() { return CellPilot.removeDuplicates(); }
// ... add all functions

// IMPORTANT: After saving, add the library:
// 1. Click Libraries (+)
// 2. Script ID: ${scriptId}
// 3. Click Add
// 4. Select Version: Latest
// 5. Click Add`;
    }
    
    function copyCode() {
      const code = document.getElementById('code').textContent;
      navigator.clipboard.writeText(code);
      alert('Code copied! Now follow Step 3.');
    }
  </script>
</body>
</html>
```

## Step 4: Deploy the Installer

1. In the installer project:
   - Click **Deploy** → **New Deployment**
   - Type: **Web app**
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone with Google account**
   - Click **Deploy**

2. Share the web app URL with beta testers

## Step 5: User Experience

### For Beta Testers:

1. **Visit your web app URL**: `https://script.google.com/macros/s/YOUR_WEB_APP_ID/exec`
2. **Authorize once**: Google sign-in
3. **Get installation code**: Automatically generated
4. **Install in their sheet**: Copy-paste process
5. **Start using**: CellPilot appears in Extensions menu

### What Users See:
- ✅ Professional installation page
- ✅ CellPilot menu in their sheets
- ✅ All features working
- ❌ No access to your source code
- ❌ Can't see implementation details

## Step 6: Managing Beta Access

### Option A: Supabase Integration

```javascript
function isBetaTester(email) {
  const response = UrlFetchApp.fetch('https://cellpilot.io/api/check-beta', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    payload: JSON.stringify({ email: email })
  });
  
  const result = JSON.parse(response.getContentText());
  return result.hasBetaAccess;
}
```

### Option B: Google Sheets List

Create a Google Sheet with beta tester emails:

```javascript
function isBetaTester(email) {
  const sheet = SpreadsheetApp.openById('YOUR_BETA_LIST_SHEET_ID');
  const emails = sheet.getRange('A:A').getValues().flat();
  return emails.includes(email);
}
```

## Step 7: Simplified Installation (Future)

### Create a one-click installer:

```javascript
function doGet(e) {
  if (e.parameter.install === 'true') {
    // Auto-install process
    return autoInstall();
  }
  // Show installation page
  return showInstallPage();
}

function autoInstall() {
  // This would require OAuth and additional permissions
  // But can create a smoother experience
}
```

## Advantages of This Approach

1. **Code Protection**: Your source code stays private
2. **Version Control**: Update library, all users get updates
3. **Access Control**: You control who can install
4. **Professional Experience**: Looks like a real add-on
5. **Easy Transition**: Similar to marketplace experience

## Next Steps

1. **Set up library deployment** from your main script
2. **Create installer web app** with the code above
3. **Test with your account** first
4. **Add beta testers** to access list
5. **Share web app URL** with testers

## Security Notes

- Library code is not visible to users
- Users only see proxy functions
- You can revoke library access anytime
- Track usage through your library
- No need to share script files

## Monitoring Usage

Add analytics to your library:

```javascript
function logUsage(functionName, userEmail) {
  // Send to your Supabase database
  const data = {
    function: functionName,
    user: userEmail,
    timestamp: new Date().toISOString()
  };
  
  // Log to your database
  UrlFetchApp.fetch('https://cellpilot.io/api/track', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    payload: JSON.stringify(data)
  });
}
```

## FAQ

**Q: Can users see my code?**
A: No, library code is completely hidden.

**Q: How do I update the beta?**
A: Deploy a new version of your library, users automatically get updates.

**Q: Can I revoke access?**
A: Yes, remove users from beta list or undeploy the library.

**Q: Is this secure?**
A: Yes, same security as Google Workspace Marketplace add-ons.

**Q: How is this different from current setup?**
A: Easier installation, no manual code copying, professional experience.