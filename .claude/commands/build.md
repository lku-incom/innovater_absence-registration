Build the SPFx absence-registration solution for production, create the SharePoint deployment package, and push all changes to GitHub.

Execute these steps in order:
1. Run `gulp bundle --ship` in the project directory to create production bundle
2. Run `gulp package-solution --ship` to create the .sppkg deployment file
3. Run `git add -A` to stage all changes
4. Run `git commit -m "Build and deploy"` to commit
5. Run `git push` to push to GitHub

Report progress after each step. The .sppkg file will be at: sharepoint/solution/absence-registration.sppkg
