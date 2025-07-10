# 🧪 SheetMind Excel Add-in Test Guide

## ✅ Current Status
- ✅ Backend server: `http://localhost:8000` 
- ✅ Add-in server: `http://localhost:3000`
- ✅ Fresh manifest: `manifest-http.xml`
- ⚠️ SSL errors: Excel trying HTTPS on old manifest

## 🔄 Step-by-Step Test Process

### 1. Remove Old Add-in (Important!)

**Excel Online:**
1. Go to [office.com/excel](https://office.com/excel)
2. **Insert** > **Office Add-ins**
3. **My Add-ins** tab
4. Find "SheetMind" → **...** → **Remove**
5. **Hard refresh**: Ctrl+Shift+R

**Desktop Excel:**
1. **File** > **Options** > **Add-ins**
2. Find "SheetMind" → **Remove**
3. Close Excel completely

### 2. Install Fresh Add-in

**Excel Online (Recommended):**
1. Open [office.com/excel](https://office.com/excel)
2. Create new workbook
3. **Insert** > **Office Add-ins** > **Upload My Add-in**
4. **Enter URL**: `http://localhost:3000/manifest-http.xml`
5. Click **Upload**

**Desktop Excel:**
1. Restart Excel
2. **Insert** > **Office Add-ins** > **Upload My Add-in**
3. **Browse**: Select `manifest-http.xml` file
4. Click **Upload**

### 3. Test the Add-in

1. **Look for button**: 🧠 SheetMind in Excel ribbon
2. **Create test data**:
   ```
   A1: Name    B1: Sales   C1: Region
   A2: John    B2: 1000    C2: North
   A3: Jane    B3: 1500    C3: South
   A4: Bob     B4: 800     C4: East
   ```
3. **Select range**: A1:C4
4. **Click SheetMind button** → Sidebar should open
5. **Test commands**:
   - "What's the total sales?"
   - "Create a chart"
   - "What's the average?"

### 4. Expected Results

**✅ Success Signs:**
- Sidebar opens with SheetMind interface
- Shows current selection context
- Chat interface loads
- No SSL errors in server logs

**❌ If Still Failing:**
- Check browser console (F12) for errors
- Verify both servers running
- Try incognito/private browser window
- Clear all browser cache

## 🐛 Troubleshooting

### Still See SSL Errors?
Excel might be caching old manifest. Try:
1. Use **completely different browser**
2. Use **manifest-http.xml** (new ID)
3. Change port: `python setup.py` on port 3001

### Add-in Won't Load?
1. Check network tab in browser dev tools
2. Verify `http://localhost:3000/taskpane.html` loads
3. Try Excel Online instead of desktop

### Backend Errors?
The OpenAI API version issue is separate. Add-in should still load and show interface.

## 📝 Test Checklist

- [ ] Removed old add-in
- [ ] Hard refreshed browser
- [ ] Used fresh manifest-http.xml
- [ ] Add-in button appears in ribbon
- [ ] Sidebar opens when clicked
- [ ] Interface shows current selection
- [ ] Chat input field is accessible
- [ ] No SSL errors in server logs

## 🎯 Quick URL Test

Direct test these URLs in browser:
- `http://localhost:3000/manifest-http.xml`
- `http://localhost:3000/taskpane.html`
- `http://localhost:8000/docs`

All should load without errors! 