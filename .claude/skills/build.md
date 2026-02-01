# Build and Deploy Skill

Build the SPFx solution, create SharePoint package, commit to git, and push to GitHub.

## Steps

1. Run `gulp bundle --ship` to create production bundle
2. Run `gulp package-solution --ship` to create the .sppkg file for SharePoint
3. Stage all changes with `git add -A`
4. Commit with message "Build and deploy [timestamp]"
5. Push to GitHub with `git push`

## Execution

Run the following commands in sequence:

```bash
cd c:\Users\info\Innovater\absence-registration
gulp bundle --ship && gulp package-solution --ship && git add -A && git commit -m "Build and deploy $(date +%Y-%m-%d_%H:%M)" && git push
```

Report the result of each step to the user. If any step fails, stop and report the error.
