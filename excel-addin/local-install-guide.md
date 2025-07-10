# Installing SheetMind in Desktop Excel

## 🎯 Two Types of Excel Add-ins

### Our Add-in (Office Add-in - Web-based)
- Uses `manifest.xml` 
- Runs in a web browser inside Excel
- Cross-platform (Windows, Mac, Online)
- **Best for:** Modern features, AI integration

### Traditional Excel Add-ins (.xlam)
- Excel-specific binary files
- VBA-based
- Windows-focused
- **Best for:** Legacy Excel features

## 🚀 Method 1: Excel Online (Recommended)

1. Go to [office.com/excel](https://office.com/excel)
2. Create a new workbook
3. **Insert** > **Office Add-ins** > **Upload My Add-in**
4. Enter manifest URL: `http://localhost:3000/manifest.xml`
5. ✅ Works immediately!

## 🖥️ Method 2: Desktop Excel Setup

### A. Enable Developer Mode (Windows)
1. **File** > **Options** > **Trust Center** > **Trust Center Settings**
2. **Trusted App Catalogs** > Check "Allow Office Add-ins to start"
3. Add `http://localhost:3000` to trusted locations

### B. Enable Add-in Development (Mac)
1. **Excel** > **Preferences** > **Authoring** > **General**
2. Check "Show Developer tab in ribbon"
3. **Developer** > **Add-ins** > **Office Add-ins**

### C. Sideload the Add-in
1. **Insert** > **Office Add-ins** > **Upload My Add-in**
2. Select manifest.xml file: `/path/to/excel-addin/manifest.xml`
3. Click **Upload**

## 🔧 Alternative: Create .xlam Version

If you prefer a traditional Excel add-in (.xlam), we can create a VBA version:

1. **Developer** > **Visual Basic** 
2. **Insert** > **Module**
3. Paste VBA code that calls our web API
4. **File** > **Save As** > **Excel Add-in (.xlam)**

## 🐛 Troubleshooting

**"File format not supported"**
- Use Excel Online instead
- Or enable Office Add-ins in Trust Center

**Add-in not loading**
- Check both servers are running (ports 8000 & 3000)
- Try Excel Online first
- Clear Excel cache

**Security warnings**
- Add localhost to trusted sites
- Enable Office Add-ins in Trust Center

## ✅ Quick Test

1. Open Excel Online: [office.com/excel](https://office.com/excel)
2. Insert > Office Add-ins > Upload My Add-in
3. URL: `http://localhost:3000/manifest.xml`
4. Look for 🧠 SheetMind button! 