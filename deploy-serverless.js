#!/usr/bin/env node
/**
 * Deploy Pinnacle Real Estate Excel Add-in to GitHub Pages (serverless)
 * This script prepares the add-in for serverless deployment
 */

const fs = require('fs');
const path = require('path');

function createServerlessDeployment() {
    console.log('🚀 Preparing Pinnacle Real Estate Excel Add-in for serverless deployment...\n');

    // 1. Build the production version
    console.log('📦 Building production version...');
    const { execSync } = require('child_process');
    
    try {
        execSync('npm run build', { stdio: 'inherit' });
        console.log('✅ Production build completed\n');
    } catch (error) {
        console.error('❌ Build failed:', error.message);
        return;
    }

    // 2. Create GitHub Pages deployment structure
    console.log('📁 Creating GitHub Pages structure...');
    
    const deployDir = 'docs';
    if (!fs.existsSync(deployDir)) {
        fs.mkdirSync(deployDir);
    }

    // Copy dist files to docs folder (GitHub Pages standard)
    const distFiles = fs.readdirSync('dist');
    distFiles.forEach(file => {
        const srcPath = path.join('dist', file);
        const destPath = path.join(deployDir, file);
        
        if (fs.statSync(srcPath).isDirectory()) {
            // Copy directory recursively
            copyDirectory(srcPath, destPath);
        } else {
            fs.copyFileSync(srcPath, destPath);
        }
    });

    // 3. Create a template manifest for GitHub Pages
    const githubPagesManifest = createGitHubPagesManifest();
    fs.writeFileSync(path.join(deployDir, 'manifest-github-pages.xml'), githubPagesManifest);

    // 4. Create deployment instructions
    const deploymentInstructions = createDeploymentInstructions();
    fs.writeFileSync('SERVERLESS_DEPLOYMENT.md', deploymentInstructions);

    // 5. Create a simple index.html for the GitHub Pages site
    const indexHtml = createIndexHtml();
    fs.writeFileSync(path.join(deployDir, 'index.html'), indexHtml);

    console.log('✅ Serverless deployment files created in /docs folder');
    console.log('✅ GitHub Pages manifest created: docs/manifest-github-pages.xml');
    console.log('✅ Deployment instructions created: SERVERLESS_DEPLOYMENT.md');
    console.log('\n🎯 Next steps:');
    console.log('1. Push this repository to GitHub');
    console.log('2. Enable GitHub Pages in repository settings');
    console.log('3. Update the manifest URLs with your GitHub Pages URL');
    console.log('4. Distribute the updated manifest.xml to users');
    console.log('\n📖 See SERVERLESS_DEPLOYMENT.md for detailed instructions');
}

function copyDirectory(src, dest) {
    if (!fs.existsSync(dest)) {
        fs.mkdirSync(dest, { recursive: true });
    }
    
    const files = fs.readdirSync(src);
    files.forEach(file => {
        const srcPath = path.join(src, file);
        const destPath = path.join(dest, file);
        
        if (fs.statSync(srcPath).isDirectory()) {
            copyDirectory(srcPath, destPath);
        } else {
            fs.copyFileSync(srcPath, destPath);
        }
    });
}

function createGitHubPagesManifest() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" 
           xsi:type="TaskPaneApp">

  <!-- Begin Basic Settings: Add-in metadata, used for all versions of Office unless override provided. -->
  <Id>12345678-1234-1234-1234-123456789012</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Pinnacle Real Estate</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Pinnacle Real Estate Add-in" />
  <Description DefaultValue="Excel add-in for Pinnacle Real Estate operating expenses calculation" />
  
  <!-- REPLACE WITH YOUR GITHUB PAGES URL -->
  <IconUrl DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://www.pinnaclerealestate.ca" />
  
  <AppDomains>
    <!-- REPLACE WITH YOUR GITHUB PAGES DOMAIN -->
    <AppDomain>https://YOUR-USERNAME.github.io</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Workbook" />
  </Hosts>

  <Requirements>
    <Sets>
      <Set Name="ExcelApi" MinVersion="1.7"/>
    </Sets>
  </Requirements>

  <DefaultSettings>
    <!-- REPLACE WITH YOUR GITHUB PAGES URL -->
    <SourceLocation DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/taskpane.html"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="Pinnacle.GetStarted.Title"/>
            <Description resid="Pinnacle.GetStarted.Description"/>
            <!-- REPLACE WITH YOUR GITHUB PAGES URL -->
            <LearnMoreUrl resid="Pinnacle.GetStarted.LearnMoreUrl"/>
          </GetStarted>

          <FunctionFile resid="Commands.Url" />

          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Pinnacle.Tab">
              <Group id="Pinnacle.Group">
                <Label resid="Pinnacle.Group.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>

                <Control xsi:type="Button" id="Pinnacle.TaskpaneButton">
                  <Label resid="Pinnacle.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Pinnacle.TaskpaneButton.Label" />
                    <Description resid="Pinnacle.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Pinnacle.RunButton">
                  <Label resid="Pinnacle.RunButton.Label" />
                  <Supertip>
                    <Title resid="Pinnacle.RunButton.Label" />
                    <Description resid="Pinnacle.RunButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>runCalculation</FunctionName>
                  </Action>
                </Control>
              </Group>
              <Label resid="Pinnacle.Tab.Label" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <!-- REPLACE WITH YOUR GITHUB PAGES URLs -->
        <bt:Image id="Icon.16x16" DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <!-- REPLACE WITH YOUR GITHUB PAGES URLs -->
        <bt:Url id="Commands.Url" DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/taskpane.html"/>
        <bt:Url id="Pinnacle.GetStarted.LearnMoreUrl" DefaultValue="https://YOUR-USERNAME.github.io/YOUR-REPO-NAME"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Pinnacle.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Pinnacle.RunButton.Label" DefaultValue="Run" />
        <bt:String id="Pinnacle.Group.Label" DefaultValue="Operating Expenses" />
        <bt:String id="Pinnacle.Tab.Label" DefaultValue="Pinnacle Real Estate" />
        <bt:String id="Pinnacle.GetStarted.Title" DefaultValue="Get started with Pinnacle Real Estate add-in!" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Pinnacle.TaskpaneButton.Tooltip" DefaultValue="Click to Show the taskpane" />
        <bt:String id="Pinnacle.RunButton.Tooltip" DefaultValue="Click to run operating expenses calculation" />
        <bt:String id="Pinnacle.GetStarted.Description" DefaultValue="Your sample add-in loaded successfully. Go to the HOME tab and click the 'Show Taskpane' button to get started." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;
}

function createDeploymentInstructions() {
    return `# Serverless Deployment Guide
## Pinnacle Real Estate Excel Add-in

### Overview
This guide shows how to deploy the Excel add-in without any server using GitHub Pages (free static hosting).

## Option 1: GitHub Pages Deployment (Recommended)

### Step 1: Create GitHub Repository
1. Create a new repository on GitHub (e.g., \`pinnacle-excel-addin\`)
2. Push this project to the repository:
   \`\`\`bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO-NAME.git
   git push -u origin main
   \`\`\`

### Step 2: Enable GitHub Pages
1. Go to your repository on GitHub
2. Click **Settings** tab
3. Scroll down to **Pages** section
4. Under **Source**, select **Deploy from a branch**
5. Choose **main** branch and **/docs** folder
6. Click **Save**
7. GitHub will provide your URL: \`https://YOUR-USERNAME.github.io/YOUR-REPO-NAME\`

### Step 3: Update Manifest URLs
1. Open \`docs/manifest-github-pages.xml\`
2. Replace all instances of:
   - \`YOUR-USERNAME\` with your GitHub username
   - \`YOUR-REPO-NAME\` with your repository name
3. Save the file

### Step 4: Test the Deployment
1. Visit \`https://YOUR-USERNAME.github.io/YOUR-REPO-NAME\`
2. You should see the add-in landing page
3. Test the manifest: \`https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/manifest-github-pages.xml\`

### Step 5: Distribute to Users
1. Provide users with the updated \`manifest-github-pages.xml\` file
2. Users install using Excel's **Insert > Get Add-ins > Upload My Add-in**

## Option 2: Other Static Hosting Services

### Netlify
1. Sign up at netlify.com
2. Drag and drop the \`docs\` folder to Netlify
3. Get your URL (e.g., \`https://amazing-name-123456.netlify.app\`)
4. Update manifest URLs accordingly

### Vercel
1. Sign up at vercel.com
2. Connect your GitHub repository
3. Set build output directory to \`docs\`
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
**Note**: This serverless deployment is production-ready and suitable for enterprise use.`;
}

function createIndexHtml() {
    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Pinnacle Real Estate Excel Add-in</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 40px 20px;
            background-color: #f5f5f5;
            color: #333;
        }
        .header {
            text-align: center;
            margin-bottom: 40px;
        }
        .logo {
            font-size: 48px;
            margin-bottom: 10px;
        }
        h1 {
            color: #0078d4;
            margin-bottom: 10px;
        }
        .subtitle {
            color: #666;
            font-size: 18px;
        }
        .card {
            background: white;
            border-radius: 8px;
            padding: 30px;
            margin: 20px 0;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .feature {
            display: flex;
            align-items: center;
            margin: 15px 0;
        }
        .feature-icon {
            font-size: 24px;
            margin-right: 15px;
            width: 30px;
        }
        .download-section {
            text-align: center;
            background: #0078d4;
            color: white;
            border-radius: 8px;
            padding: 30px;
            margin: 30px 0;
        }
        .download-button {
            display: inline-block;
            background: white;
            color: #0078d4;
            padding: 12px 24px;
            text-decoration: none;
            border-radius: 6px;
            font-weight: bold;
            margin: 10px;
        }
        .download-button:hover {
            background: #f0f0f0;
        }
        .instructions {
            background: #e8f4fd;
            border-left: 4px solid #0078d4;
            padding: 20px;
            margin: 20px 0;
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">🏢</div>
        <h1>Pinnacle Real Estate</h1>
        <p class="subtitle">Excel Add-in for Operating Expenses Calculation</p>
    </div>

    <div class="card">
        <h2>Features</h2>
        <div class="feature">
            <span class="feature-icon">⚡</span>
            <span>Automated operating expenses calculation and data transfer</span>
        </div>
        <div class="feature">
            <span class="feature-icon">🎯</span>
            <span>Professional ribbon interface with building-themed icons</span>
        </div>
        <div class="feature">
            <span class="feature-icon">🔄</span>
            <span>Forces Excel recalculation even in Manual mode</span>
        </div>
        <div class="feature">
            <span class="feature-icon">📊</span>
            <span>Clean taskpane interface with real-time status updates</span>
        </div>
        <div class="feature">
            <span class="feature-icon">🛡️</span>
            <span>Comprehensive error handling and user feedback</span>
        </div>
    </div>

    <div class="download-section">
        <h2>Download Add-in</h2>
        <p>Get the manifest file to install the add-in in Excel</p>
        <a href="manifest-github-pages.xml" class="download-button" download>
            📥 Download Manifest
        </a>
        <a href="Pinnacle_Real_Estate_Excel_Addin_Installation_Guide.pdf" class="download-button" download>
            📖 Installation Guide
        </a>
    </div>

    <div class="instructions">
        <h3>Quick Installation</h3>
        <ol>
            <li>Download the manifest file above</li>
            <li>Open Excel and go to <strong>Insert > Get Add-ins</strong></li>
            <li>Click <strong>Upload My Add-in</strong></li>
            <li>Select the downloaded manifest file</li>
            <li>Look for the "Pinnacle Real Estate" tab in Excel</li>
        </ol>
    </div>

    <div class="card">
        <h2>System Requirements</h2>
        <ul>
            <li>Excel 2016 or later (Windows/Mac) or Excel Online</li>
            <li>Internet connection for initial installation</li>
            <li>Excel model with "Outputs" and "Software Engineer Cash Flow" sheets</li>
        </ul>
    </div>

    <div class="card">
        <h2>Support</h2>
        <p>For technical support or questions about the add-in, contact your Pinnacle Real Estate IT administrator.</p>
        <p><strong>Version:</strong> 1.0.0<br>
        <strong>Last Updated:</strong> January 2025</p>
    </div>
</body>
</html>`;
}

// Run the deployment preparation
if (require.main === module) {
    createServerlessDeployment();
}

module.exports = { createServerlessDeployment };
