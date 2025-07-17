#!/bin/bash

# Manual deployment script for GitHub Pages
# This creates the docs folder locally with the correct manifest setup

echo "🚀 Building Pinnacle Real Estate Excel Add-in for GitHub Pages..."

# Build the project
echo "📦 Building production files..."
npm run build

# Create docs directory
echo "📁 Creating GitHub Pages structure..."
rm -rf docs
mkdir -p docs

# Copy built files to docs
echo "📋 Copying built files..."
cp -r dist/* docs/

# Copy GitHub-specific manifest as the main manifest.xml for client download
echo "📄 Setting up GitHub Pages manifest..."
cp manifest-github.xml docs/manifest.xml

# Create professional landing page
echo "🎨 Creating landing page..."
cat > docs/index.html << 'EOF'
<!DOCTYPE html>
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
        <a href="manifest.xml" class="download-button" download>
            📥 Download Manifest
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
        <strong>Last Updated:</strong> $(date +'%B %Y')</p>
    </div>
</body>
</html>
EOF

# Replace the date placeholder with actual date
sed -i.bak "s/\$(date +'%B %Y')/$(date +'%B %Y')/g" docs/index.html
rm -f docs/index.html.bak

# Copy installation guide if it exists
echo "📖 Adding installation guide..."
if [ -f "Pinnacle_Real_Estate_Excel_Addin_Installation_Guide.pdf" ]; then
    cp "Pinnacle_Real_Estate_Excel_Addin_Installation_Guide.pdf" docs/
    echo "✅ Installation guide added"
else
    echo "⚠️  Installation guide not found, skipping..."
fi

echo ""
echo "✅ GitHub Pages deployment ready!"
echo ""
echo "📋 Next steps:"
echo "1. git add docs/"
echo "2. git commit -m 'Deploy to GitHub Pages'"
echo "3. git push"
echo "4. Enable GitHub Pages in repository Settings > Pages"
echo "5. Choose 'Deploy from a branch' and select 'main' branch, '/docs' folder"
echo ""
echo "🌐 Your add-in will be available at:"
echo "https://solanki0494.github.io/pinnacle-excel-addin"
echo ""
echo "📥 Clients will download manifest.xml with GitHub Pages URLs"
