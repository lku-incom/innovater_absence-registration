---
name: build
description: Update documentation, build SPFx solution, package for SharePoint, commit and push to GitHub. Use when deploying the app or when the user says "build".
user-invocable: true
---

Build the SPFx absence-registration solution for production and deploy to GitHub.

## Prerequisites

- Node.js 18.x installed
- npm dependencies installed (`npm install` if needed)
- Git configured with remote origin

## Steps

Execute these steps in order:

### Step 1: Update Documentation

Review the current codebase and update `AbsenceRegistrationDocumentation.html` to reflect any changes:

1. Read the current documentation file
2. Scan the `src/webparts/absenceRegistration/` directory for:
   - New or removed components in `components/`
   - New or changed services in `services/`
   - New or changed models in `models/`
   - Changes to the main web part file
3. Check `package.json` for any dependency version changes
4. Update the documentation HTML file with:
   - New/removed components in the Project Structure section
   - New/changed services and their methods
   - Updated technology versions if changed
   - Any new features or configuration options
5. Only make changes if there are actual differences - don't update if everything is current

### Step 2: Clean Previous Build (Optional)

If there are build issues, run:
```bash
gulp clean
```

### Step 3: Create Production Bundle

Run the TypeScript compilation and webpack bundling:
```bash
gulp bundle --ship
```

**Expected output:**
- TypeScript compilation completes without errors
- Webpack bundles assets to `dist/` folder
- No warnings about missing dependencies

**Common issues:**
- TypeScript errors: Fix the code issues before proceeding
- Missing dependencies: Run `npm install`

### Step 4: Create SharePoint Package

Generate the .sppkg deployment package:
```bash
gulp package-solution --ship
```

**Expected output:**
- Package created at `sharepoint/solution/absence-registration.sppkg`
- No manifest errors

**Common issues:**
- Invalid manifest: Check `config/package-solution.json`
- Missing bundle: Ensure Step 3 completed successfully

### Step 5: Stage Changes

Stage all modified files for commit:
```bash
git add -A
```

**Files typically included:**
- Updated documentation (`AbsenceRegistrationDocumentation.html`)
- Build artifacts in `sharepoint/solution/`
- Any source code changes

### Step 6: Commit Changes

Create a commit with a descriptive message:
```bash
git commit -m "Build and deploy"
```

**Note:** If no changes were made, this step will indicate nothing to commit - that's okay.

### Step 7: Push to GitHub

Push the commit to the remote repository:
```bash
git push
```

**Expected output:**
- Changes pushed to `origin/main`
- No merge conflicts

## Output Artifacts

After successful completion:

| Artifact | Location |
|----------|----------|
| SPFx Package | `sharepoint/solution/absence-registration.sppkg` |
| Bundle JS | `dist/*.js` |
| Documentation | `AbsenceRegistrationDocumentation.html` |

## Deployment

After the build completes, the `.sppkg` file can be uploaded to:
- **SharePoint App Catalog:** `https://innovaterdk.sharepoint.com/sites/appcatalog`
- **Site Collection App Catalog:** For site-scoped deployment

## Error Handling

If any step fails:
1. **Stop immediately** - do not proceed to the next step
2. **Report the error** to the user with the full error message
3. **Suggest fixes** based on the error type:
   - TypeScript errors → Show the file and line number
   - npm errors → Suggest `npm install` or clearing node_modules
   - Git errors → Check for uncommitted changes or remote conflicts

## Rollback

If deployment causes issues, revert using:
```bash
git revert HEAD
git push
```
