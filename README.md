# GrantMath Excel Add-in

Static deployment for GrantMath Excel Add-in.

## Files

- `index.html` - Task pane UI
- `addin.js` - JavaScript logic
- `manifest.xml` - Office Add-in manifest
- `icon-*.png` - Add-in icons

## Deployment

This is deployed to Vercel as a static site.

After deployment, download `manifest.xml` and upload to Excel.

## Usage

1. Download manifest.xml from deployed URL
2. Excel → Insert → Get Add-ins → Upload My Add-in
3. Upload manifest.xml
4. Click GrantMath button in ribbon
