/**
 * CellPilot Beta Installer Web App
 * =================================
 * This web app helps beta users install CellPilot
 */

/**
 * Handle GET requests - show installer page
 */
function doGet(e) {
  // Try multiple methods to get the user email
  let userEmail = Session.getActiveUser().getEmail();
  
  // If email is empty, try effective user
  if (!userEmail) {
    userEmail = Session.getEffectiveUser().getEmail();
  }
  
  console.log('User email from session:', userEmail);
  console.log('Active user:', Session.getActiveUser().getEmail());
  console.log('Effective user:', Session.getEffectiveUser().getEmail());
  
  // Check if user has beta access
  const hasBeta = checkBetaAccess(userEmail);
  console.log('Beta access check result:', hasBeta);
  
  if (!hasBeta) {
    return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 16px 16px 0 0; text-align: center;">
          <h1 style="margin: 0;">CellPilot Beta</h1>
        </div>
        <div style="background: white; padding: 30px; border: 1px solid #e5e7eb; border-radius: 0 0 16px 16px;">
          <h2 style="color: #1a202c; margin-top: 0;">Beta Access Required</h2>
          <p style="color: #4a5568; line-height: 1.6;">
            To install CellPilot, you need to:
          </p>
          <ol style="color: #4a5568; line-height: 2;">
            <li>Sign up at <a href="https://www.cellpilot.io" target="_blank" style="color: #667eea;">www.cellpilot.io</a></li>
            <li>Go to your dashboard</li>
            <li>Click "Activate Beta Access"</li>
            <li>Return here to install</li>
          </ol>
          <p style="color: #718096; font-size: 14px; margin-top: 20px;">
            <strong>Your email:</strong> ${userEmail || 'Not signed in'}<br>
            <small>Make sure you're signed in with the same Google account you used to sign up.</small>
          </p>
          <div style="text-align: center; margin-top: 30px;">
            <a href="https://www.cellpilot.io" target="_blank" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; display: inline-block;">
              Sign Up for Beta Access
            </a>
          </div>
        </div>
      </div>
    `).setTitle('CellPilot Beta Access');
  }
  
  // Return the installer page
  return HtmlService.createTemplateFromFile('installer')
    .evaluate()
    .setTitle('Install CellPilot Beta')
    .setWidth(800)
    .setHeight(600);
}

/**
 * Check if user has beta access
 */
function checkBetaAccess(email) {
  if (!email) return false;
  
  try {
    // Check against CellPilot API
    const response = UrlFetchApp.fetch('https://www.cellpilot.io/api/check-beta', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ email: email }),
      muteHttpExceptions: true // Get response even on error
    });
    
    const responseText = response.getContentText();
    console.log('API Response:', responseText);
    console.log('Response Code:', response.getResponseCode());
    
    if (response.getResponseCode() !== 200) {
      console.error('API returned error code:', response.getResponseCode());
      return false;
    }
    
    const result = JSON.parse(responseText);
    console.log('Parsed result:', result);
    return result.hasBeta === true;
    
  } catch (e) {
    console.error('Error checking beta access:', e.toString());
    // If API is down, deny access for security
    return false;
  }
}

/**
 * Get the installation code
 */
function getInstallationCode() {
  const scriptId = '1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O';
  const version = '10';
  
  // Return the full installation code
  const code = getFullBetaInstallerCode(scriptId, version);
  
  return {
    scriptId: scriptId,
    version: version,
    code: code,
    success: true
  };
}

/**
 * Get the full beta installer code
 */
function getFullBetaInstallerCode(scriptId, version) {
  // Return the complete proxy code that users will paste
  return `/**
 * CellPilot Beta Installation Script
 * ===================================
 * Version: 1.0.0-beta
 * 
 * IMPORTANT: After pasting this code, you MUST add the CellPilot library:
 * 1. Click Libraries (+) in the left sidebar
 * 2. Script ID: ${scriptId}
 * 3. Version: ${version}
 * 4. Identifier: CellPilot
 * 5. Click Add
 */

function onOpen() {
  try {
    CellPilot.onOpen();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Please add CellPilot library. See instructions above.');
  }
}

function onInstall(e) {
  CellPilot.onInstall(e);
}

// Main Features
function showCellPilotSidebar() { CellPilot.showCellPilotSidebar(); }
function tableize() { return CellPilot.tableize(); }
function removeDuplicates() { return CellPilot.removeDuplicates(); }
function cleanData() { return CellPilot.cleanData(); }
function showSmartFormulaAssistant() { CellPilot.showSmartFormulaAssistant(); }
function showFormulaBuilder() { CellPilot.showFormulaBuilder(); }
function showAdvancedRestructuring() { CellPilot.showAdvancedRestructuring(); }
function showIndustryTemplates() { CellPilot.showIndustryTemplates(); }
function showConditionalFormattingWizard() { CellPilot.showConditionalFormattingWizard(); }
function showCrossSheetFormulaBuilder() { CellPilot.showCrossSheetFormulaBuilder(); }
function showDataPipeline() { CellPilot.showDataPipeline(); }
function showDataValidationGenerator() { CellPilot.showDataValidationGenerator(); }
function showSmartFormulaDebugger() { CellPilot.showSmartFormulaDebugger(); }
function showPivotTableAssistant() { CellPilot.showPivotTableAssistant(); }
function showSettings() { CellPilot.showSettings(); }
function showHelpFeedback() { CellPilot.showHelpFeedback(); }

// Supporting Functions
function include(filename) { return CellPilot.include(filename); }
function getCurrentUserContext() { return CellPilot.getCurrentUserContext(); }
function getExtendedUserContext() { return CellPilot.getExtendedUserContext(); }

// For complete list of 200+ functions, visit:
// https://www.cellpilot.io/docs/functions`;
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Test function to debug API access
 */
function testBetaAccess() {
  const testEmail = 'freddyfiesta@gmail.com';
  console.log('Testing beta access for:', testEmail);
  
  const result = checkBetaAccess(testEmail);
  console.log('Result:', result);
  
  // Also try a direct API call
  try {
    const response = UrlFetchApp.fetch('https://www.cellpilot.io/api/check-beta', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ email: testEmail }),
      muteHttpExceptions: true
    });
    
    console.log('Direct API test:');
    console.log('Status:', response.getResponseCode());
    console.log('Headers:', response.getAllHeaders());
    console.log('Content:', response.getContentText());
  } catch (e) {
    console.log('Direct API error:', e.toString());
  }
  
  return result;
}