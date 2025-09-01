# CellPilot Development Prompt Template

## How to Use This Template

When starting a new development session, copy this template and fill in your specific request.

---

## PROMPT TEMPLATE:

I'm working on CellPilot, a Google Sheets add-on project. Please follow the project context and workflow rules.

**Project Context File**: `/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md`

**Current Task**: [Describe what you want to do]

**Type of Change**:
- [ ] Adding new feature
- [ ] Modifying existing feature
- [ ] Bug fix
- [ ] UI/UX change
- [ ] Database change
- [ ] Documentation

**Specific Requirements**:
[List any specific requirements or constraints]

**Please ensure**:
1. Follow the 4-location update rule for any new functions
2. Use the correct workflow from PROJECT_CONTEXT.md
3. Run check-sync.sh before finalizing
4. Deploy and update URLs as needed
5. Commit all changes together

---

## EXAMPLE PROMPTS:

### Example 1: Adding New Feature
```
I'm working on CellPilot. Please follow the project context at /home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md

Task: Add a new "Export to PDF" feature

Requirements:
- Create showPdfExport() function
- Add menu item in main sidebar
- Follow the 4-location update rule
- Deploy new versions after implementation
```

### Example 2: Fixing a Bug
```
I'm working on CellPilot. Please follow the project context at /home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md

Task: Fix the duplicate removal function not detecting all duplicates

Current issue: [describe the bug]

Please ensure we update all necessary locations and test thoroughly.
```

### Example 3: UI Update
```
I'm working on CellPilot. Please follow the project context at /home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md

Task: Update the TableizeTemplate.html to improve the UI

Changes needed:
- Better color scheme
- Larger buttons
- Improved spacing

Ensure changes are pushed to Apps Script after modification.
```

## QUICK REFERENCE PATHS:

When you need to reference specific aspects:

**For workflow rules**: "Follow the workflow in `/home/freddy/cellpilot/docs/DEVELOPMENT_WORKFLOW.md`"

**For quick commands**: "Use commands from `/home/freddy/cellpilot/docs/QUICK_REFERENCE.md`"

**For sync checking**: "Run `/home/freddy/cellpilot/scripts/check-sync.sh` to verify"

**For project structure**: "Check `/home/freddy/cellpilot/docs/PROJECT_STRUCTURE.md`"

---

## CRITICAL REMINDERS TO INCLUDE:

Always mention these if relevant to your task:

1. "Ensure all 4 locations are updated for new functions"
2. "Deploy new versions and update BetaAccessCard.tsx"
3. "Test in Google Sheets before committing"
4. "Run check-sync.sh to verify synchronization"
5. "Commit all related changes together"

---

Save this template and customize it for each development session!