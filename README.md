# SheetMind 🧠📊

> An open-source Cursor-like tool for Excel that lets natural language agents execute commands on your spreadsheets.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

## 🚀 What is SheetMind?

SheetMind brings the power of AI-driven natural language processing to Excel, allowing you to:

- **Talk to your spreadsheets**: Use natural language to manipulate data, create formulas, and generate insights
- **Automate complex tasks**: Let AI agents handle repetitive Excel operations
- **Smart data analysis**: Get intelligent suggestions and automated data processing
- **Cross-platform compatibility**: Works with Excel on Windows, macOS, and web-based Excel

## ✨ Features

- 🗣️ **Natural Language Interface**: "Add a column for profit margins" → Automatically creates formulas
- 🤖 **Intelligent Agents**: AI-powered agents that understand Excel operations
- 🔧 **Excel Integration**: Seamless integration with local Excel installations and web Excel
- 📊 **Smart Analysis**: Automated data insights and visualization suggestions
- 🌐 **Web Interface**: Clean, modern web UI for easy interaction
- 🔓 **Open Source**: MIT licensed, community-driven development

## 🛠️ Installation

### Prerequisites

- Python 3.8 or higher
- Excel (local installation or Office 365 access)
- API key for OpenAI or similar LLM service

### Quick Start

```bash
# Clone the repository
git clone https://github.com/yourusername/sheetmind.git
cd sheetmind

# Install dependencies
pip install -r requirements.txt

# Set up environment variables
cp .env.example .env
# Edit .env with your API keys

# Run SheetMind
python src/main.py
```

## 🎯 Usage Examples

### Basic Commands

```
"Create a pivot table from the sales data"
"Add a formula to calculate 20% tax on column C"
"Sort the data by date in descending order"
"Find all rows where revenue is greater than $10,000"
"Create a chart showing monthly trends"
```

### Advanced Operations

```
"Analyze the correlation between marketing spend and sales"
"Generate a summary report of Q4 performance"
"Clean the data by removing duplicates and fixing formatting"
"Create a dashboard with key metrics"
```

## 🏗️ Architecture

```
SheetMind/
├── src/
│   ├── agents/          # AI agent implementations
│   ├── integrations/    # Excel integration layers
│   ├── nlp/            # Natural language processing
│   ├── ui/             # User interfaces
│   └── main.py         # Application entry point
├── tests/              # Test suites
└── docs/               # Documentation
```

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Inspired by Cursor and other AI-powered development tools
- Built with love for the Excel community
- Special thanks to all contributors

## 🔗 Links

- [Documentation](https://sheetmind.readthedocs.io)
- [Issues](https://github.com/yourusername/sheetmind/issues)
- [Discussions](https://github.com/yourusername/sheetmind/discussions) 