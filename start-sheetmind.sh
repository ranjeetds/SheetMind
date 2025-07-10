#!/bin/bash

# 🧠 SheetMind AI Startup Script
# Start the complete SheetMind system with local Ollama AI

echo "🧠 Starting SheetMind AI for Excel..."

# Check if Ollama is installed
if ! command -v ollama &> /dev/null; then
    echo "❌ Ollama is not installed!"
    echo "Please install from: https://ollama.ai"
    echo "Then run: ollama pull llama2"
    exit 1
fi

# Check if a model is available
if ! ollama list | grep -q "llama2\|codellama\|mistral"; then
    echo "📥 No AI model found. Downloading llama2..."
    ollama pull llama2
fi

# Start Ollama in background if not running
if ! pgrep -f "ollama serve" > /dev/null; then
    echo "🚀 Starting Ollama AI service..."
    ollama serve &
    sleep 3
fi

# Create .env file if it doesn't exist
if [ ! -f .env ]; then
    echo "⚙️ Creating configuration..."
    cat > .env << EOF
DEFAULT_AI_PROVIDER=ollama
OLLAMA_URL=http://localhost:11434
OLLAMA_MODEL=llama2
EOF
fi

# Install Python dependencies if needed
if [ ! -d "venv" ] && [ ! -f ".venv/bin/activate" ]; then
    echo "📦 Installing Python dependencies..."
    pip install -r requirements.txt
fi

echo "🌐 Starting SheetMind backend..."
echo "📍 Backend will be available at: http://localhost:8000"
echo ""
echo "📋 Next steps:"
echo "1. Open Excel"
echo "2. Install Script Lab add-in (Insert → Get Add-ins → Search 'Script Lab')"
echo "3. In Script Lab, Import → From File → Select 'excel-addin/script-lab-proper.js'"
echo "4. Click Run to start SheetMind AI!"
echo ""
echo "🔄 Press Ctrl+C to stop the backend when done"
echo "───────────────────────────────────────────────"

# Start the FastAPI backend
uvicorn src.ui.web.app:app --host 0.0.0.0 --port 8000 --reload 