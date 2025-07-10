# SheetMind Excel Add-in - Quick Start ⚡

Get SheetMind running as an Excel sidebar in 5 minutes!

## 🎯 What You're Building

A **Cursor-like AI sidebar** that appears directly in Excel, allowing you to chat with your spreadsheets using natural language.

## ⚡ Quick Setup (5 Steps)

### 1. Clone & Install
```bash
git clone https://github.com/yourusername/sheetmind.git
cd sheetmind
pip install -r requirements.txt
```

### 2. Configure AI
```bash
cp env.example .env
# Edit .env and add your OpenAI or Anthropic API key
```

### 3. Start Backend
```bash
python src/main.py web --port 8000
```

### 4. Serve Add-in (New Terminal)
```bash
cd excel-addin
python setup.py
```

### 5. Install in Excel
1. Open Excel
2. **Insert** > **Office Add-ins** > **Upload My Add-in**
3. Select `excel-addin/manifest.xml`
4. Click **Upload**
5. Look for **🧠 SheetMind** button in ribbon

## 🎉 Test It Out

1. **Open or create** an Excel file with some data
2. **Click SheetMind** button to open sidebar  
3. **Select some data** (e.g., A1:C10)
4. **Type**: "What's the average of these numbers?"
5. **Watch** as it creates a formula automatically!

## 🎨 Interface Preview

```
Excel → [🧠 SheetMind] → Sidebar Opens:

┌─────────────────────────┐
│ 🧠 SheetMind            │
├─────────────────────────┤
│ ✅ Connected!           │
├─────────────────────────┤
│ 📍 Sheet1: A1:C10      │
├─────────────────────────┤
│ 📊 📈 🧮               │ ← Quick actions
├─────────────────────────┤
│ You: "Sum column A"     │
│ 🧠: "Added SUM formula  │
│     in cell D1"         │
├─────────────────────────┤
│ [Type message here...] │
└─────────────────────────┘
```

## 🚀 Example Commands

**With data selected:**
- "Calculate the total"
- "Create a bar chart"  
- "What's the average?"
- "Sort by first column"
- "Analyze this data"

**The AI will:**
- ✅ Detect your selection automatically
- ✅ Create formulas in appropriate cells  
- ✅ Insert charts directly in Excel
- ✅ Provide data insights
- ✅ Remember conversation context

## 🐛 Troubleshooting

**Add-in not appearing?**
- Check both servers are running (ports 8000 & 3000)
- Try refreshing Excel

**AI not responding?**
- Verify API key in `.env` file
- Check backend console for errors

**Operations not working?**
- Ensure you have data selected in Excel
- Try simple commands first

## 🎯 Next Steps

- Try different chart types: "Create a pie chart"
- Experiment with analysis: "Find the top 5 values"
- Test with different data types
- Explore the quick action buttons

## 🔗 Full Documentation

- [Complete Excel Add-in Guide](EXCEL_ADDIN_README.md)
- [Original Web Version](QUICKSTART.md)

---

**You now have Cursor for Excel! 🎉**

Select data → Chat → Watch Excel magic happen! 