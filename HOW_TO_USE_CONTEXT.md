# How to Use CellPilot Context in Future Development

## 🎯 Quick Start

For ANY CellPilot development work, start your prompt with:

```
I'm working on CellPilot. Please follow the project context at:
/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md
```

## 📋 What This Ensures

When you include the context reference, the AI will:

1. **Follow the 4-location rule** for function updates
2. **Use correct directories** for each component
3. **Run proper commands** for deployment
4. **Update all necessary files** in sequence
5. **Deploy and update URLs** correctly
6. **Commit changes properly**

## 🚀 Example Opening Phrases

### For New Features:
```
I'm working on CellPilot. Please follow the project context at 
/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md

I want to add a new feature that [description]. Ensure we update all 
4 locations and follow the deployment workflow.
```

### For Bug Fixes:
```
I'm working on CellPilot. Please follow the project context at 
/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md

There's a bug in [feature]. Please fix it following the proper workflow.
```

### For UI Changes:
```
I'm working on CellPilot. Please follow the project context at 
/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md

I need to update the [template name] UI. Make sure to push changes 
to Apps Script after.
```

## 📁 Key Files to Reference

| Purpose | File to Reference |
|---------|------------------|
| General development | `/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md` |
| Detailed workflow | `/home/freddy/cellpilot/docs/DEVELOPMENT_WORKFLOW.md` |
| Quick commands | `/home/freddy/cellpilot/docs/QUICK_REFERENCE.md` |
| Check sync | `/home/freddy/cellpilot/scripts/check-sync.sh` |

## ⚡ Quick Copy-Paste Starter

Copy this for any CellPilot work:

```
I'm working on CellPilot. Please follow the project context at:
/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md

Task: [YOUR TASK HERE]

Ensure:
1. Update all 4 locations for new functions
2. Run check-sync.sh before deploying
3. Deploy new versions if proxies changed
4. Update BetaAccessCard.tsx with new URL
5. Commit all changes together
```

## 🔍 Verification Phrases

Add these to ensure things are done right:

- "Please verify all 4 locations are updated"
- "Run check-sync.sh to confirm synchronization"
- "Show me the deployment command and new URL"
- "Confirm all files are ready to commit"

## 💡 Pro Tips

1. **Always include the context path** - Even for small changes
2. **Be specific about the task** - Clear requirements = correct implementation
3. **Ask for verification** - "Show me what files were updated"
4. **Request the full workflow** - "Walk me through each step"

## 📝 The Golden Rule

**ALWAYS START WITH:**
```
I'm working on CellPilot. Please follow the project context at:
/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md
```

This single line ensures everything is done correctly according to your established workflow!

---

## Example Full Prompt

Here's a complete example of a well-structured prompt:

```
I'm working on CellPilot. Please follow the project context at:
/home/freddy/cellpilot/.claude/PROJECT_CONTEXT.md

Task: Add a new feature to analyze cell colors and patterns

Requirements:
- Create showColorAnalyzer() function
- Add it to the Data Cleaning menu section
- Include a simple HTML template

Please:
1. Implement in main library first
2. Add exports and proxies to all 4 locations
3. Create and push the HTML template
4. Run check-sync.sh to verify
5. Deploy new versions
6. Update the website URL
7. Prepare everything for commit

Show me each step as you do it.
```

---

Save this file and refer to it whenever you start a new CellPilot development session!