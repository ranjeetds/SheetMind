# SheetMind Quick Start Guide 🚀

Welcome to SheetMind! This guide will help you get up and running in minutes.

## 🔧 Installation

### Prerequisites

1. **Python 3.8+** installed on your system
2. **Excel** (local installation) or access to **Office 365**
3. **API Key** for OpenAI or Anthropic (for AI features)

### Install SheetMind

```bash
# Clone the repository
git clone https://github.com/yourusername/sheetmind.git
cd sheetmind

# Install dependencies
pip install -r requirements.txt

# Or install in development mode
pip install -e .
```

## 🔑 Configuration

1. **Set up your API keys** (required for AI features):

```bash
# Copy the environment template
cp .env.example .env

# Edit .env with your API keys
nano .env
```

Add your API keys to the `.env` file:

```env
OPENAI_API_KEY=your_openai_api_key_here
# OR
ANTHROPIC_API_KEY=your_anthropic_api_key_here

DEFAULT_AI_PROVIDER=openai  # or anthropic
```

## 🚀 Quick Start

### Option 1: Web Interface (Recommended)

Start the web server:

```bash
python src/main.py web
```

Then open your browser to: `http://localhost:8000`

### Option 2: Command Line Interface

Start interactive chat mode:

```bash
python src/main.py chat
```

Or with a specific file:

```bash
python src/main.py chat --file examples/sample_sales_data.csv
```

## 📊 Try It Out

1. **Upload the sample data**: Use `examples/sample_sales_data.csv` (or convert it to Excel format)

2. **Try these commands**:

   - `"Show me a summary of the data"`
   - `"Create a bar chart from the revenue data"`
   - `"Calculate the total revenue"`
   - `"Sort data by revenue descending"`
   - `"Show me the top 5 customers by revenue"`
   - `"Filter products where quantity > 20"`

## 🎯 Example Commands

### Data Analysis
```
"What's the average revenue per sale?"
"Show me statistics for the revenue column"
"Find correlations in the data"
```

### Charts & Visualizations
```
"Create a line chart showing revenue trends"
"Make a pie chart of sales by region"
"Generate a bar chart of top products"
```

### Data Manipulation
```
"Sort by date ascending"
"Filter electronics products only"
"Show rows where revenue > 5000"
```

### Calculations
```
"Sum all revenue"
"Average quantity sold"
"Maximum price in the dataset"
"Count total orders"
```

## 🌐 Web Interface Features

- **📁 File Upload**: Drag & drop Excel files
- **💬 Chat Interface**: Natural language commands
- **📊 Data Info Panel**: Live data statistics
- **💡 Example Commands**: Click to try
- **🔄 Real-time Processing**: Instant responses

## 🛠️ Troubleshooting

### Common Issues

**1. "No module named 'win32com'"**
- On Windows: `pip install pywin32`
- On Mac/Linux: This is expected (Windows-only feature)

**2. "API key not found"**
- Make sure your `.env` file has the correct API key
- Check that the key is valid and has credits

**3. "File upload failed"**
- Ensure the file is a valid Excel (.xlsx) or CSV format
- Check file permissions and size (max 50MB)

**4. "Excel operations not working"**
- Try using file-based operations (works cross-platform)
- For Windows COM API features, ensure Excel is installed

### Need Help?

- 📚 Check the [full documentation](README.md)
- 🐛 Report issues on [GitHub](https://github.com/yourusername/sheetmind/issues)
- 💬 Join the [discussions](https://github.com/yourusername/sheetmind/discussions)

## 🎉 What's Next?

Once you're comfortable with the basics:

1. **Explore Advanced Features**: Try complex data analysis commands
2. **Create Custom Workflows**: Chain multiple operations together
3. **Integrate with Your Tools**: Use the API endpoints for automation
4. **Contribute**: Help improve SheetMind by contributing code or feedback

Happy spreadsheeting with AI! 🧠📊 