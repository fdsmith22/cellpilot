/**
 * CellPilot Beta Installer Web App
 * =================================
 * This web app helps beta users install CellPilot
 */

/**
 * Handle GET requests - show installer page
 */
function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();
  
  // Check if user has beta access
  if (!checkBetaAccess(userEmail)) {
    return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 16px 16px 0 0; text-align: center;">
          <h1 style="margin: 0;">🚀 CellPilot Beta</h1>
        </div>
        <div style="background: white; padding: 30px; border: 1px solid #e5e7eb; border-radius: 0 0 16px 16px;">
          <h2 style="color: #1a202c; margin-top: 0;">Beta Access Required</h2>
          <p style="color: #4a5568; line-height: 1.6;">
            To install CellPilot, you need to:
          </p>
          <ol style="color: #4a5568; line-height: 2;">
            <li>Sign up at <a href="https://cellpilot.io" target="_blank" style="color: #667eea;">cellpilot.io</a></li>
            <li>Go to your dashboard</li>
            <li>Click "Activate Beta Access"</li>
            <li>Return here to install</li>
          </ol>
          <p style="color: #718096; font-size: 14px; margin-top: 20px;">
            <strong>Your email:</strong> ${userEmail || 'Not signed in'}<br>
            <small>Make sure you're signed in with the same Google account you used to sign up.</small>
          </p>
          <div style="text-align: center; margin-top: 30px;">
            <a href="https://cellpilot.io" target="_blank" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; display: inline-block;">
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
    const response = UrlFetchApp.fetch('https://cellpilot.io/api/check-beta', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ email: email })
    });
    
    const result = JSON.parse(response.getContentText());
    return result.hasBeta === true;
    
  } catch (e) {
    console.error('Error checking beta access:', e);
    // If API is down, deny access for security
    return false;
  }
}

/**
 * Get the installation code
 */
function getInstallationCode() {
  // Read the BETA_INSTALLER.gs content
  const scriptId = '1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O';
  const version = '10';
  
  return {
    scriptId: scriptId,
    version: version,
    success: true
  };
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}