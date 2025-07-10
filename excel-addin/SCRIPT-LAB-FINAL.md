# 🧠 SheetMind AI for Excel - Script Lab Solution (Ollama Powered)

**The working, production-ready solution using local AI via Ollama**

## ✅ What Works Now

This solution provides a **Cursor-like AI assistant directly within Excel** using:
- ✅ **Microsoft Script Lab** (trusted Microsoft add-in)
- ✅ **Office.js API** for direct Excel manipulation
- ✅ **Local FastAPI backend** with Ollama integration
- ✅ **Natural language commands** with AI understanding
- ✅ **Real-time Excel context** awareness
- ✅ **Privacy-focused** - everything runs locally

## 🚀 Quick Setup

### Step 1: Install Ollama (Local AI)

Download and install Ollama from https://ollama.ai

**macOS/Linux:**
```bash
# Install Ollama
curl -fsSL https://ollama.ai/install.sh | sh

# Download a model (choose one)
ollama pull llama2          # 3.8GB - Fast, good for basic tasks
ollama pull llama2:13b      # 7.4GB - Better reasoning
ollama pull codellama       # 3.8GB - Good for Excel/coding tasks
ollama pull mistral         # 4.1GB - Another good option

# Start Ollama (if not auto-started)
ollama serve
```

**Windows:**
1. Download installer from https://ollama.ai/download/windows
2. Run the installer
3. Open Command Prompt and run:
```cmd
ollama pull llama2
```

### Step 2: Start the Backend

```bash
# In the project directory
cd /Users/admin/Project/Fast/ExcelCursor

# Install dependencies
pip install -r requirements.txt

# Create .env file (optional - for configuration)
echo "DEFAULT_AI_PROVIDER=ollama" > .env
echo "OLLAMA_URL=http://localhost:11434" >> .env
echo "OLLAMA_MODEL=llama2" >> .env

# Start the backend
uvicorn src.ui.web.app:app --host 0.0.0.0 --port 8000 --reload
```

### Step 3: Install Script Lab in Excel

1. Open Excel
2. Go to **Insert** → **Get Add-ins** (or **Office Add-ins**)
3. Search for "**Script Lab**"
4. Install the **Script Lab** add-in by Microsoft
5. Click **Script Lab** in the ribbon

### Step 4: Load SheetMind AI

1. In Script Lab, click **Import**
2. Click **From File**
3. Select `excel-addin/script-lab-proper.js`
4. Click **Run** button

You'll see the SheetMind AI interface appear with:
- 🟢 **AI Status**: Connected to local Ollama
- 📊 **Quick Actions**: Sum, Currency, Chart, Bold
- 💬 **Chat Interface**: Natural language commands
- 📋 **Context Display**: Current Excel selection

## 🎯 Usage Examples

### Basic Commands (Always Work)
- "sum the selected column"
- "format as currency" 
- "make this bold"
- "create a chart"
- "analyze this data"
- "clear selection"
- "create table"

### AI-Powered Commands (When Ollama is Running)
- "Calculate the average of column B and put it in cell C1"
- "Find the highest value in this range and highlight it"
- "Create a summary table of this sales data"
- "Format this data as a professional report"
- "What insights can you find in this data?"

### Advanced Operations
- "sort by first column"
- "highlight important cells"
- "freeze panes here"
- "analyze data trends"

## 🔧 Configuration

### Environment Variables (.env file)
```bash
# AI Provider (always use ollama for local AI)
DEFAULT_AI_PROVIDER=ollama

# Ollama Configuration
OLLAMA_URL=http://localhost:11434
OLLAMA_MODEL=llama2

# Alternative models you can try:
# OLLAMA_MODEL=codellama    # Good for Excel/coding tasks
# OLLAMA_MODEL=mistral      # Alternative model
# OLLAMA_MODEL=llama2:13b   # Larger, more capable model
```

### Model Recommendations

**For Fast Performance:**
- `llama2` (3.8GB) - Quick responses, good for basic Excel tasks

**For Better AI Reasoning:**
- `codellama` (3.8GB) - Specialized for code/formulas
- `llama2:13b` (7.4GB) - More capable, better understanding
- `mistral` (4.1GB) - Good balance of speed and capability

## 🛠️ Architecture

```
Excel Script Lab ↔ Local FastAPI Backend ↔ Ollama AI
     (Office.js)     (http://localhost:8000)   (Local LLM)
```

- **Script Lab**: Runs in Excel, provides UI and Office.js integration
- **FastAPI Backend**: Processes requests, handles AI communication
- **Ollama**: Local AI model for natural language understanding

## 🔍 Features

### ✅ Working Features
- Direct Excel manipulation via Office.js
- Real-time context awareness (selection, worksheet)
- Natural language command processing
- Fallback to basic commands when AI unavailable
- Local AI processing (no data leaves your machine)
- Professional Cursor-like interface
- Quick action buttons for common tasks
- Conversation history
- Error handling and status monitoring

### 🎯 AI Capabilities
- Understands Excel context and current selection
- Generates appropriate Excel operations
- Provides intelligent suggestions
- Handles complex multi-step operations
- Explains what it's doing

## 🚨 Troubleshooting

### AI Status Shows Red (🔴 Backend not running)
```bash
# Check if backend is running
curl http://localhost:8000/capabilities

# If not, start it:
uvicorn src.ui.web.app:app --host 0.0.0.0 --port 8000 --reload
```

### Ollama Not Working
```bash
# Check if Ollama is running
ollama list

# Start Ollama service
ollama serve

# Test a model
ollama run llama2 "Hello"
```

### Script Lab Won't Load
1. Ensure you have the latest Excel
2. Check internet connection (needed for Script Lab itself)
3. Try restarting Excel
4. Re-import the script file

### Performance Issues
1. Use a smaller model: `ollama pull llama2` instead of `llama2:13b`
2. Close other applications to free RAM
3. Check if your system meets Ollama requirements

## 🏗️ Files Structure

```
excel-addin/
├── script-lab-proper.js     # Working Script Lab code
├── SCRIPT-LAB-FINAL.md      # This documentation
└── README-SIMPLE.md         # Simple overview

src/
├── agents/
│   ├── base_agent.py        # Ollama integration
│   └── excel_context_agent.py
├── ui/web/
│   └── app.py               # FastAPI backend
└── ...

requirements.txt             # Updated for Ollama (aiohttp)
```

## 🎉 Why This Solution Works

1. **No Upload Issues**: Uses Microsoft's trusted Script Lab
2. **Full Excel Integration**: Office.js provides complete Excel API access
3. **Local AI**: Ollama keeps everything private and fast
4. **Graceful Fallback**: Works even when AI is unavailable
5. **Professional UI**: Looks and feels like Cursor
6. **Real Context**: Knows exactly what you have selected
7. **Extensible**: Easy to add new commands and AI capabilities

## 🔄 Development

To modify or extend SheetMind:

1. **Backend Changes**: Edit files in `src/` and restart the server
2. **Frontend Changes**: Edit `script-lab-proper.js` and re-run in Script Lab
3. **New AI Models**: Change `OLLAMA_MODEL` in `.env` and restart
4. **New Commands**: Add to `processCommand()` function in the JavaScript

## 🎊 Success!

You now have a **Cursor-like AI assistant running directly in Excel** with:
- ✅ Local AI processing via Ollama
- ✅ Complete Excel integration 
- ✅ Natural language understanding
- ✅ Privacy protection (nothing leaves your machine)
- ✅ Professional interface

**The future of Excel automation is here, and it's running locally on your machine!** 🚀 