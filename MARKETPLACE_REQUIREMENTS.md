# Google Workspace Marketplace Requirements Checklist
*Official requirements for marketplace approval - Updated January 2025*

## 🎯 Overview
This document contains the official requirements from Google's documentation for publishing CellPilot to the Google Workspace Marketplace. All requirements must be met before submission for review.

---

## ✅ Pre-Submission Requirements

### 1. Google Cloud Project Setup
- [ ] Create a Google Cloud Project
- [ ] Enable necessary APIs
- [ ] Configure OAuth consent screen
- [ ] Complete brand verification (2-3 business days)
- [ ] Set user type to "External" for public apps
- [ ] Set publishing status to "Published" (not "Testing")

### 2. OAuth Verification Process
- [ ] Submit for OAuth verification if using sensitive/restricted scopes
- [ ] Provide demo video showing scope usage
- [ ] Complete security assessment (if accessing server-side data)
- [ ] Allow several weeks for restricted scope verification
- [ ] Annual re-verification required for restricted scopes

### 3. Scope Configuration
- [ ] Match scopes across all configurations:
  - [ ] OAuth consent screen
  - [ ] Google Workspace Marketplace SDK
  - [ ] App manifest (appsscript.json)
- [ ] Use most narrowly focused scopes possible
- [ ] Avoid requesting unnecessary scopes
- [ ] Document justification for each scope

---

## 📋 App Listing Requirements

### 4. App Details
- [ ] App name (must be unique and not misleading)
- [ ] Developer name and contact information
- [ ] App description (clear, accurate, keyword-appropriate)
- [ ] Categories and tags
- [ ] Pricing information
- [ ] Support email address

### 5. Graphic Assets (REQUIRED)
- [ ] **Application Icon/Logo**
  - Format: PNG or JPEG
  - Size: 128x128 pixels minimum
  - High quality, represents app functionality
  - Cannot use Google logos
  - Must be hosted and accessible
  
- [ ] **Card Banner** (for marketplace listing)
  - High-quality banner image
  - Accurately represents app
  
- [ ] **Screenshots** (minimum 3-5 recommended)
  - High quality images
  - Accurately represent app capabilities
  - Show actual functionality
  - Standard Google service screenshots allowed (unaltered)
  - Must not be misleading or deceptive

### 6. Support Links
- [ ] **Privacy Policy URL** (REQUIRED)
  - Must be publicly accessible
  - GDPR compliant if serving EU users
  - Clear data handling practices
  
- [ ] **Terms of Service URL** (REQUIRED)
  - Publicly accessible
  - Professional and comprehensive
  
- [ ] **Support/Help URL**
  - Documentation or help center
  - Contact information for support

### 7. Optional Assets
- [ ] YouTube promo video(s)
- [ ] Additional marketing materials
- [ ] User testimonials/reviews

---

## 🔒 Security & Compliance

### 8. Security Requirements
- [ ] No hardcoded API keys or secrets
- [ ] Secure data transmission (HTTPS/SSL)
- [ ] Proper error handling without exposing sensitive data
- [ ] Input validation and sanitization
- [ ] Rate limiting implementation

### 9. Data Privacy
- [ ] Clear data collection disclosure
- [ ] User consent for data processing
- [ ] Data retention policies
- [ ] User data deletion capabilities
- [ ] GDPR compliance (if applicable)

### 10. OAuth Scope Verification
- [ ] **Sensitive Scopes**: Require Google review
- [ ] **Restricted Scopes**: Require extensive verification
  - Security assessment by approved third-party
  - Annual re-verification
  - Demo video required
  - Several weeks for approval

---

## 🧪 Testing Requirements

### 11. Functionality Testing
- [ ] App is fully functional
- [ ] No critical bugs or errors
- [ ] Positive user experience
- [ ] Works from Argentina, Canada, and United States (reviewer locations)
- [ ] All advertised features work as described

### 12. Browser Compatibility
- [ ] Chrome (latest version)
- [ ] Firefox (latest version)
- [ ] Safari (latest version)
- [ ] Edge (latest version)

### 13. Performance
- [ ] Reasonable load times
- [ ] Efficient resource usage
- [ ] No memory leaks
- [ ] Proper cleanup on uninstall

---

## 📝 Review Process

### 14. Submission Checklist
- [ ] OAuth verification complete
- [ ] All scopes properly configured and matched
- [ ] Store listing complete with all required fields
- [ ] Graphics and screenshots uploaded
- [ ] Privacy Policy and Terms of Service links working
- [ ] App fully tested and functional
- [ ] No use of misleading information
- [ ] Compliance with Google trademark guidelines

### 15. Common Rejection Reasons (AVOID)
- [ ] Incorrect OAuth consent screen setup
- [ ] User type set to "Internal" instead of "External"
- [ ] Publishing status set to "Testing"
- [ ] OAuth verification not completed
- [ ] Inappropriate use of Google trademarks
- [ ] Missing or low-quality graphics
- [ ] Misleading app description or functionality
- [ ] Broken features or critical bugs
- [ ] Missing privacy policy or terms of service

### 16. Review Timeline
- **Private Apps**: Immediate availability to organization
- **Public Apps**: Several days for review
- **OAuth Verification**: Several weeks for restricted scopes
- **Brand Verification**: 2-3 business days

---

## 🚀 Post-Approval

### 17. Maintenance Requirements
- [ ] Annual security reassessment (if using restricted scopes)
- [ ] Keep OAuth verification current
- [ ] Update listing for any functionality changes
- [ ] Maintain support channels
- [ ] Monitor and respond to user reviews

### 18. Updates and Changes
- [ ] New sensitive/restricted scopes require re-verification
- [ ] Major functionality changes may require re-review
- [ ] Keep all URLs (privacy, terms, support) active
- [ ] Update screenshots if UI changes significantly

---

## 📊 CellPilot Specific Action Items

### Immediate Actions Required:
1. **Host logo file**: Upload to https://www.cellpilot.io/logo-48x48.png
2. **Create Google Cloud Project**: Set up OAuth consent screen
3. **Prepare screenshots**: Capture 3-5 high-quality screenshots
4. **Record demo video**: Show OAuth scope usage
5. **Complete OAuth verification**: Submit for review

### Documentation Needed:
1. **User guide**: Comprehensive how-to documentation
2. **API documentation**: For developers/advanced users
3. **Video tutorials**: Feature demonstrations
4. **Help center**: FAQ and troubleshooting

### Technical Requirements:
1. **Fix logo URL** in appsscript.json manifest
2. **Implement API endpoints** for telemetry
3. **Add global error handler**
4. **Complete rate limiting**
5. **Performance optimization** for ML features

---

## 📅 Estimated Timeline

### Phase 1: Preparation (Week 1-2)
- Google Cloud Project setup
- OAuth consent configuration
- Asset preparation (logo, screenshots)
- Documentation creation

### Phase 2: Verification (Week 3-5)
- OAuth verification submission
- Brand verification
- Security assessment scheduling
- Demo video creation

### Phase 3: Review (Week 6-7)
- Marketplace submission
- Address reviewer feedback
- Final adjustments

### Phase 4: Launch (Week 8)
- Public availability
- Marketing announcement
- User onboarding

---

## 🔗 Official Resources

- [App Review Process](https://developers.google.com/workspace/marketplace/about-app-review)
- [OAuth Configuration](https://developers.google.com/workspace/marketplace/configure-oauth-consent-screen)
- [Create Store Listing](https://developers.google.com/workspace/marketplace/create-listing)
- [Publishing Guide](https://developers.google.com/workspace/marketplace/how-to-publish)
- [Branding Guidelines](https://developers.google.com/workspace/marketplace/terms/branding)
- [Program Policies](https://developers.google.com/workspace/marketplace/terms/policies)
- [OAuth Verification Help](https://support.google.com/cloud/answer/13463073)
- [Restricted Scope Verification](https://developers.google.com/identity/protocols/oauth2/production-readiness/restricted-scope-verification)

---

## ⚠️ Important Notes

1. **Visibility Cannot Change**: Once you choose private/public visibility, it cannot be changed later
2. **Regional Access**: Ensure app works from Argentina, Canada, and United States (reviewer locations)
3. **Annual Requirements**: Restricted scopes require annual re-verification and security assessment
4. **Timeline Variance**: Review times can vary based on complexity and current queue
5. **OAuth First**: OAuth verification must complete before marketplace approval

---

*Last Updated: January 2025*
*Based on official Google documentation current as of 2025*