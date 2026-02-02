---
name: build
description: Build SPFx solution, package for SharePoint, commit and push to GitHub. Use when deploying the app or when the user says "build".
user-invocable: true
---

Build the SPFx absence-registration solution for production and deploy to GitHub.

## Steps

Execute these steps in order:

1. Run `gulp bundle --ship` to create production bundle
2. Run `gulp package-solution --ship` to create the .sppkg file
3. Run `git add -A` to stage all changes
4. Run `git commit -m "Build and deploy"` to commit
5. Run `git push` to push to GitHub

## Execution

Run each command sequentially in the project directory.

Report the result of each step to the user. If any step fails, stop and report the error.

The .sppkg file will be at: `sharepoint/solution/absence-registration.sppkg`
