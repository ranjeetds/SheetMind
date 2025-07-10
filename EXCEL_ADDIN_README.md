# SheetMind Excel Add-in 🧠📊

> An AI-powered Excel add-in that brings natural language automation directly to your spreadsheets - just like Cursor for Excel!

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Excel Add-in](https://img.shields.io/badge/Excel-Add--in-green.svg)](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/)

## 🚀 What is SheetMind Excel Add-in?

SheetMind is an Excel add-in that appears as a **sidebar panel** in Excel (just like Cursor's interface), allowing you to:

- 🗣️ **Chat with your spreadsheet**: Use natural language to manipulate data
- 🎯 **Context-aware operations**: Works with your current selection automatically  
- 📊 **Direct Excel integration**: Creates charts, formulas, and calculations in real-time
- 🔄 **Live interaction**: See changes happen instantly in your spreadsheet
- 🌐 **Cross-platform**: Works on Excel for Windows, Mac, and Excel Online

## ✨ How It Works

1. **Open Excel** and load your data
2. **Click the SheetMind button** in the ribbon to open the sidebar
3. **Select data** you want to work with
4. **Chat naturally**: "Create a chart", "Calculate totals", "Analyze this data"
5. **Watch magic happen**: SheetMind executes operations directly in Excel

![SheetMind Interface Demo](assets/demo.gif)

## 🎯 Example Interactions

### With Data Selected
```
You: "What's the average of these numbers?"
SheetMind: Added AVERAGE formula in cell D1
→ Creates =AVERAGE(A1:C10) formula next to your selection

You: "Create a bar chart from this data"  
SheetMind: Created a bar chart from your selected data
→ Inserts chart directly into your worksheet

You: "Sort this by the first column"
SheetMind: Sorted the selected data in ascending order
→ Sorts your selection in-place
```

### Smart Context Awareness
- **Detects headers**: Knows when first row contains column names
- **Recognizes data types**: Handles numeric vs text data appropriately  
- **Suggests actions**: Quick buttons for common operations
- **Shows context**: Displays current selection info

## 🛠️ Installation & Setup

### Prerequisites
- **Excel 2016 or later** (Windows, Mac, or Excel Online)
- **Internet connection** for AI features
- **API key** for OpenAI or Anthropic

### Step 1: Set Up the Backend
```bash
# Clone the repository
git clone https://github.com/yourusername/sheetmind.git
cd sheetmind

# Install dependencies
pip install -r requirements.txt

# Configure API keys
cp env.example .env
# Edit .env with your OpenAI/Anthropic API key

# Start the backend server
python src/main.py web --port 8000
```

### Step 2: Serve the Add-in Files
```bash
# In a new terminal, serve the add-in files
cd excel-addin
python -m http.server 3000
```

### Step 3: Install in Excel

#### Option A: Sideload for Development
1. Open Excel
2. Go to **Insert** > **Office Add-ins**
3. Click **Upload My Add-in**
4. Select `excel-addin/manifest.xml`
5. Click **Upload**

#### Option B: Use Manifest URL (Recommended)
1. Upload the `excel-addin` folder to a web server
2. Update manifest URLs to point to your server
3. In Excel: **Insert** > **Office Add-ins** > **Upload My Add-in**
4. Enter the manifest URL

### Step 4: Start Using SheetMind
1. Look for the **🧠 SheetMind** button in the Excel ribbon
2. Click it to open the sidebar
3. Select some data and start chatting!

## 🎨 Interface Features

### Sidebar Layout (Like Cursor)
```
┌─────────────────────────┐
│ 🧠 SheetMind            │ ← Header
├─────────────────────────┤
│ ✅ Connected and ready! │ ← Status
├─────────────────────────┤
│ 📍 Sheet1: A1:C10      │ ← Current Context
│    (10 rows × 3 cols)   │
├─────────────────────────┤
│ 📊 📈 🧮               │ ← Quick Actions
├─────────────────────────┤
│                         │
│ Chat Messages...        │ ← Chat Area
│                         │
├─────────────────────────┤
│ Try these:              │ ← Suggestions
│ • Sum column A          │
│ • Create bar chart      │
├─────────────────────────┤
│ [Type message here...] │ ← Input
└─────────────────────────┘
```

### Smart Features
- **Context Display**: Shows current worksheet and selection
- **Quick Actions**: One-click buttons for common tasks
- **Live Suggestions**: Context-aware command suggestions
- **Status Updates**: Real-time connection and operation status
- **Auto-resize**: Input area adapts to message length

## 🔧 Supported Operations

### ✅ Currently Supported
- **Calculations**: SUM, AVERAGE, COUNT, MIN, MAX formulas
- **Charts**: Bar, line, and pie charts from selected data
- **Analysis**: Data summaries and statistical insights
- **Sorting**: Sort selected ranges in ascending/descending order
- **Context**: Works with current selection automatically

### 🚧 Coming Soon
- **Filtering**: Advanced data filtering
- **Formatting**: Cell styling and conditional formatting
- **Pivot Tables**: Interactive data summarization
- **Custom Functions**: AI-generated Excel functions
- **Data Import**: Smart data import and cleaning

## 🏗️ Architecture

```
Excel Desktop/Online
│
├── SheetMind Add-in (Sidebar)
│   ├── Office.js (Excel integration)
│   ├── HTML/CSS/JS (User interface)
│   └── Real-time communication
│
└── Backend Server (localhost:8000)
    ├── FastAPI web server
    ├── Excel Context Agent (AI)
    ├── NLP Command Processor
    └── AI Provider (OpenAI/Anthropic)
```

### How It Works Internally
1. **User types** natural language command
2. **Add-in captures** current Excel context (selection, worksheet, data)
3. **Sends to backend** with context and message
4. **AI processes** command with Excel context awareness
5. **Returns operations** (formulas, charts, etc.)
6. **Add-in executes** operations directly in Excel using Office.js

## 🔒 Security & Privacy

- **Local processing**: Your data stays in Excel
- **Secure communication**: HTTPS connections to AI services
- **No data storage**: We don't store your spreadsheet data
- **API key control**: You control your own AI service keys

## 🚀 Development

### Project Structure
```
excel-addin/
├── manifest.xml         # Add-in definition
├── taskpane.html       # Main sidebar interface  
├── commands.html       # Function file
└── assets/             # Icons and images

src/
├── agents/
│   └── excel_context_agent.py  # Context-aware agent
├── ui/web/
│   └── app.py          # Backend API with /chat-excel endpoint
└── ...
```

### Adding New Commands
1. **Update NLP patterns** in `command_processor.py`
2. **Add handler method** in `excel_context_agent.py`
3. **Test with Excel** using the add-in interface

### Custom AI Providers
The system supports multiple AI providers. Add new ones by extending `BaseAgent`.

## 🐛 Troubleshooting

### Add-in Not Loading
- Check that both servers are running (port 8000 and 3000)
- Verify manifest.xml URLs are correct
- Try refreshing Excel or reloading the add-in

### AI Not Responding
- Check API keys in `.env` file
- Verify backend server is running
- Check browser console for errors (F12 in task pane)

### Operations Not Executing
- Ensure you have data selected in Excel
- Check that the selection contains valid data
- Try simpler commands first

## 🤝 Contributing

We welcome contributions! Here's how to get started:

1. **Fork the repository**
2. **Set up development environment** (see Installation)
3. **Test with Excel** to ensure everything works
4. **Make your changes** and test thoroughly
5. **Submit a pull request**

### Development Guidelines
- Test all changes with real Excel data
- Ensure cross-platform compatibility (Windows/Mac/Online)
- Follow Office Add-ins best practices
- Update documentation for new features

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Inspired by **Cursor** and other AI-powered development tools
- Built with **Microsoft Office Add-ins** platform
- Uses **Office.js** for Excel integration
- Powered by **OpenAI** and **Anthropic** AI models

## 🔗 Links

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Report Issues](https://github.com/yourusername/sheetmind/issues)
- [Discussions](https://github.com/yourusername/sheetmind/discussions)

---

**Ready to supercharge your Excel experience with AI?** 🚀

Install SheetMind and start chatting with your spreadsheets today! 