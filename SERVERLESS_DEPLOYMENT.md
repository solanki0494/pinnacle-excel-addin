# Serverless Deployment Guide
## Pinnacle Real Estate Excel Add-in

### Overview
This guide shows how to deploy the Excel add-in without any server using GitHub Pages (free static hosting).

## Option 1: GitHub Pages Deployment (Recommended)

### Step 1: Create GitHub Repository
1. Create a new repository on GitHub (e.g., `pinnacle-excel-addin`)
2. Push this project to the repository:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO-NAME.git
   git push -u origin main
   ```

### Step 2: Enable GitHub Pages
1. Go to your repository on GitHub
2. Click **Settings** tab
3. Scroll down to **Pages** section
4. Under **Source**, select **Deploy from a branch**
5. Choose **main** branch and **/docs** folder
6. Click **Save**
7. GitHub will provide your URL: `https://YOUR-USERNAME.github.io/YOUR-REPO-NAME`

### Step 3: Update Manifest URLs
1. Open `docs/manifest-github-pages.xml`
2. Replace all instances of:
   - `YOUR-USERNAME` with your GitHub username
   - `YOUR-REPO-NAME` with your repository name
3. Save the file

### Step 4: Test the Deployment
1. Visit `https://YOUR-USERNAME.github.io/YOUR-REPO-NAME`
2. You should see the add-in landing page
3. Test the manifest: `https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/manifest-github-pages.xml`

### Step 5: Distribute to Users
1. Provide users with the updated `manifest-github-pages.xml` file
2. Users install using Excel's **Insert > Get Add-ins > Upload My Add-in**

## Option 2: Other Static Hosting Services

### Netlify
1. Sign up at netlify.com
2. Drag and drop the `docs` folder to Netlify
3. Get your URL (e.g., `https://amazing-name-123456.netlify.app`)
4. Update manifest URLs accordingly

### Vercel
1. Sign up at vercel.com
2. Connect your GitHub repository
3. Set build output directory to `docs`
4. Deploy and get your URL
5. Update manifest URLs accordingly

## Benefits of Serverless Deployment

### ✅ Advantages
- **No Server Costs**: Completely free hosting
- **High Availability**: 99.9% uptime with CDN
- **HTTPS by Default**: Secure connections required by Office
- **Global Distribution**: Fast loading worldwide
- **Easy Updates**: Push to GitHub to update
- **Version Control**: Full git history

### ⚠️ Considerations
- **Static Files Only**: No server-side processing (not needed for this add-in)
- **Public Repository**: Code is visible (can use private repos with paid plans)
- **GitHub Pages Limits**: 1GB storage, 100GB bandwidth/month (more than sufficient)

## Security Notes
- All add-in logic runs client-side in Excel
- No sensitive data is transmitted to servers
- HTTPS encryption for all communications
- Office.js provides sandboxed execution environment

## Troubleshooting

### Common Issues
1. **404 Errors**: Check that GitHub Pages is enabled and URLs are correct
2. **CORS Errors**: Ensure all URLs use HTTPS and match the manifest
3. **Manifest Validation**: Use Office Add-in Validator online tool

### Testing
1. Test manifest validation: https://dev.office.com/add-in-validator
2. Test in Excel Online first (easier debugging)
3. Check browser developer tools for errors

## Support
- GitHub Pages Documentation: https://pages.github.com/
- Office Add-ins Documentation: https://docs.microsoft.com/en-us/office/dev/add-ins/
- Manifest Reference: https://docs.microsoft.com/en-us/office/dev/add-ins/reference/manifest/

---
**Note**: This serverless deployment is production-ready and suitable for enterprise use.