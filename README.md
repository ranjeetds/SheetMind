# 🧠 SheetMind AI - Excel's Cursor-like AI Assistant

A **local AI-powered Excel automation tool** that brings Cursor-like intelligence directly to Excel using **Ollama** for privacy and performance.

![SheetMind AI](https://img.shields.io/badge/Excel-AI%20Assistant-blue) ![Local AI](https://img.shields.io/badge/Ollama-Local%20AI-green) ![Privacy](https://img.shields.io/badge/Privacy-First-orange)

## ✨ What Is SheetMind?

SheetMind transforms Excel into an **AI-powered workspace** where you can:
- 🗣️ **Talk to your data** in natural language
- 🤖 **Automate complex tasks** with simple commands  
- 📊 **Analyze data instantly** using local AI
- 🔒 **Keep everything private** - no data leaves your machine
- ⚡ **Work efficiently** with Cursor-like AI assistance

## 🚀 Quick Start (3 Minutes)

### 1. Install Ollama (Local AI)
```bash
# macOS/Linux
curl -fsSL https://ollama.ai/install.sh | sh
ollama pull llama2

# Windows: Download from https://ollama.ai
```

### 2. Start SheetMind
```bash
# Clone and start
git clone <this-repo>
cd ExcelCursor
./start-sheetmind.sh    # Mac/Linux
# OR
start-sheetmind.bat     # Windows
```

### 3. Install in Excel
1. Open Excel → **Insert** → **Get Add-ins** 
2. Search "**Script Lab**" → Install
3. Script Lab → **Import** → `excel-addin/script-lab-proper.js`
4. Click **Run** → 🎉 **SheetMind AI is ready!**

## 💬 Example Commands

**Basic Operations:**
- "Sum column A"
- "Format as currency"
- "Create a chart from this data"
- "Make headers bold"

**AI-Powered Analysis:**
- "What insights can you find in this data?"
- "Create a summary table of sales by region"
- "Find outliers and highlight them"
- "Format this as a professional report"

**Advanced Automation:**
- "Calculate quarterly growth rates"
- "Generate pivot table for expense analysis"
- "Find duplicate entries and mark them"

## 🏗️ Architecture

```
Excel (Script Lab) ↔ FastAPI Backend ↔ Ollama AI
     Office.js      localhost:8000     Local LLM
```

**Why This Approach Works:**
- ✅ **No Upload Issues**: Uses Microsoft's trusted Script Lab
- ✅ **Full Excel Access**: Office.js provides complete Excel API
- ✅ **Local AI**: Ollama keeps data private and responses fast
- ✅ **Graceful Fallback**: Works even when AI is unavailable
- ✅ **Cursor-like UI**: Professional, familiar interface

## 📁 Project Structure

```
ExcelCursor/
├── 🧠 excel-addin/
│   ├── script-lab-proper.js      # Working Script Lab code
│   └── SCRIPT-LAB-FINAL.md       # Complete setup guide
├── ⚙️ src/
│   ├── agents/
│   │   ├── base_agent.py         # Ollama integration
│   │   └── excel_context_agent.py
│   └── ui/web/app.py             # FastAPI backend
├── 🚀 start-sheetmind.sh         # Mac/Linux startup
├── 🚀 start-sheetmind.bat        # Windows startup
└── 📋 requirements.txt           # Python dependencies
```

## 🔧 Configuration

Create a `.env` file to customize:
```bash
DEFAULT_AI_PROVIDER=ollama
OLLAMA_URL=http://localhost:11434
OLLAMA_MODEL=llama2              # or codellama, mistral, llama2:13b
```

**Model Recommendations:**
- `llama2` (3.8GB) - Fast, good for basic tasks
- `codellama` (3.8GB) - Best for Excel formulas and analysis  
- `llama2:13b` (7.4GB) - More capable reasoning
- `mistral` (4.1GB) - Good balance of speed and capability

## 🎯 Features

### ✅ Working Now
- Direct Excel manipulation via Office.js
- Real-time context awareness (knows your selection)
- Natural language command processing
- Local AI with Ollama integration
- Fallback to basic commands when AI unavailable
- Professional Cursor-like interface
- Quick action buttons for common tasks
- Conversation history and error handling

### 🚀 Capabilities
- **Data Analysis**: Instant insights and statistics
- **Chart Creation**: Automatic visualization
- **Formatting**: Professional styling and layouts
- **Formula Generation**: Complex Excel formulas
- **Data Cleaning**: Find and fix issues
- **Automation**: Multi-step operations
- **Context Awareness**: Understands your current selection

## 🛠️ Development

### Backend Development
```bash
# Edit files in src/ and restart
uvicorn src.ui.web.app:app --reload
```

### Frontend Development  
```bash
# Edit excel-addin/script-lab-proper.js
# Re-run in Script Lab to see changes
```

### Add New AI Models
```bash
ollama pull <model-name>
# Update OLLAMA_MODEL in .env
```

## 🔍 Troubleshooting

**AI Not Working?**
```bash
# Check Ollama
ollama list
ollama serve

# Check backend
curl http://localhost:8000/capabilities
```

**Excel Issues?**
1. Ensure latest Excel version
2. Check Script Lab is installed and updated
3. Verify internet connection (for Script Lab itself)
4. Try re-importing the script file

## 🎊 Why SheetMind?

1. **Privacy First**: Everything runs locally via Ollama
2. **No Upload Issues**: Uses Microsoft's trusted Script Lab
3. **Full Integration**: Complete Excel API access via Office.js
4. **AI-Powered**: Understands context and generates smart operations
5. **Cursor-like UX**: Familiar, professional interface
6. **Easy Setup**: Works in 3 minutes with simple scripts

## 📄 License

Open source project - see LICENSE file for details.

## 🤝 Contributing

Contributions welcome! This project aims to democratize AI-powered Excel automation.

---

**Transform your Excel workflow with the power of local AI.** 
**No cloud required. No data shared. Just intelligent automation.** 🚀 