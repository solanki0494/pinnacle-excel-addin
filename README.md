# Pinnacle Real Estate Excel Add-in

Excel add-in for Pinnacle Real Estate to automate operating expenses calculations and update cash flow projections.

## Features

- **Custom Ribbon Tab**: "Pinnacle Real Estate" tab with professional building icons
- **Operating Expenses Calculation**: Automatically recalculates and copies operating expenses data
- **Cash Flow Update**: Updates the Software Engineer Cash Flow tab with calculated values
- **Taskpane Interface**: Clean, user-friendly interface with emoji icons
- **Service-Based Architecture**: Modular, maintainable code structure

## Project Structure

```
excel-addin/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Main UI interface
│   │   └── taskpane.ts      # Taskpane functionality
│   ├── commands/
│   │   ├── commands.html    # Ribbon commands page
│   │   └── commands.ts      # Ribbon button functionality
│   └── services/
│       └── operatingExpensesService.ts  # Core business logic
├── assets/                  # Professional building icons (16px, 32px, 64px, 80px)
├── manifest.xml            # Add-in manifest
├── webpack.config.js       # Build configuration
└── dist/                   # Built files
```

## Requirements

- Node.js (v16 or higher)
- Excel 2016 or later (Windows/Mac) or Excel Online
- HTTPS development server for testing

## Quick Start

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Start development server:**
   ```bash
   npm run dev-server
   ```
   Server runs at: `https://localhost:3000`

3. **Sideload the add-in:**
   ```bash
   npm run sideload
   ```

4. **Build for production:**
   ```bash
   npm run build
   ```

## Usage

### From Ribbon
1. Open Excel with your Pinnacle Real Estate model
2. Look for the "Pinnacle Real Estate" tab in the ribbon
3. Click the "Run" button (🏢 building icon) to execute calculations

### From Taskpane
1. Click "Show Taskpane" in the Pinnacle Real Estate ribbon
2. Use the "▶️ Run Calculation" button in the taskpane interface
3. View status messages and results

## Configuration

The add-in works with the following Excel worksheet structure:

- **Outputs Sheet**: Contains operating expenses data (range O32:CI35)
- **Software Engineer Cash Flow Sheet**: Target for updated values (range O32:CI35)

### Data Flow

1. Add-in forces Excel to recalculate (even in Manual mode)
2. Copies operating expenses from Outputs sheet (O32:CI35)
3. Pastes values to Software Engineer Cash Flow sheet (O32:CI35)
4. Preserves formatting and shows completion status

### Customizing Cell Ranges

To adapt to different Excel models, update the ranges in `src/services/operatingExpensesService.ts`:

```typescript
const sourceRange = outputsSheet.getRange("O32:CI35");
const targetRange = engineeringSheet.getRange("O32:CI35");
```

## Technical Details

- **Framework**: Office.js API with TypeScript
- **Build Tool**: Webpack 5 with hot reload
- **Architecture**: Service-based with clean separation of concerns
- **Icons**: Professional building icons (16px, 32px, 64px, 80px)
- **Supported Hosts**: Excel (Desktop and Online)
- **Permissions**: ReadWriteDocument

## Development Features

- **Hot Reload**: Automatic refresh during development
- **Professional Icons**: Building-themed icons for ribbon buttons
- **Error Handling**: Comprehensive error handling and user feedback
- **Office.js Integration**: Proper initialization and loading sequence
- **CSP Compliant**: Content Security Policy configured for development

## Deployment

### Option 1: Automated GitHub Pages Deployment (Recommended)

**Fully automated serverless deployment!** No servers, no manual steps:

1. **One-time setup:**
   - Go to repository Settings > Pages
   - Source: Select "GitHub Actions"
   - Save

2. **Deploy automatically:**
   ```bash
   git add .
   git commit -m "Update add-in"
   git push
   ```

3. **GitHub Actions automatically:**
   - ✅ Builds the project
   - ✅ Creates production files
   - ✅ Uses `manifest-github.xml` as `manifest.xml` for client download
   - ✅ Deploys to GitHub Pages
   - ✅ Makes available at: `https://solanki0494.github.io/pinnacle-excel-addin`

### Option 2: Manual Deployment (If GitHub Actions billing issue)

```bash
npm run deploy
git add docs/
git commit -m "Deploy to GitHub Pages"
git push
```

Then enable GitHub Pages: Settings > Pages > Deploy from branch > main > /docs

## Manifest Files

The project uses two manifest files:

- **`manifest.xml`** - For local development (localhost URLs)
- **`manifest-github.xml`** - For GitHub Pages deployment (GitHub URLs)

During deployment, `manifest-github.xml` is automatically copied as `manifest.xml` so clients download the correct version with GitHub Pages URLs.

### Option 3: Traditional Server Deployment

1. **Build for production:**
   ```bash
   npm run build
   ```

2. **Host the `dist/` folder** on a secure HTTPS server

3. **Update manifest.xml URLs** to point to your hosted location

4. **Distribute the manifest.xml** file to users

## Troubleshooting

- Verify worksheet names: "Outputs" and "Software Engineer Cash Flow"
- Check that cell ranges O32:CI35 contain the expected data
- Use browser developer tools to debug JavaScript issues
- Ensure HTTPS is used for hosting (required by Office.js)
