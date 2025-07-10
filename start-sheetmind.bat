@echo off
echo 🧠 Starting SheetMind AI for Excel...

REM Check if Ollama is installed
ollama --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Ollama is not installed!
    echo Please install from: https://ollama.ai/download/windows
    echo Then run: ollama pull llama2
    pause
    exit /b 1
)

REM Check if a model is available
ollama list | findstr /R "llama2 codellama mistral" >nul
if errorlevel 1 (
    echo 📥 No AI model found. Downloading llama2...
    ollama pull llama2
)

REM Start Ollama service if not running
tasklist /FI "IMAGENAME eq ollama.exe" 2>NUL | find /I /N "ollama.exe">NUL
if errorlevel 1 (
    echo 🚀 Starting Ollama AI service...
    start "" ollama serve
    timeout /t 3 /nobreak >nul
)

REM Create .env file if it doesn't exist
if not exist ".env" (
    echo ⚙️ Creating configuration...
    (
        echo DEFAULT_AI_PROVIDER=ollama
        echo OLLAMA_URL=http://localhost:11434
        echo OLLAMA_MODEL=llama2
    ) > .env
)

REM Install Python dependencies if needed
if not exist "venv" if not exist ".venv" (
    echo 📦 Installing Python dependencies...
    pip install -r requirements.txt
)

echo 🌐 Starting SheetMind backend...
echo 📍 Backend will be available at: http://localhost:8000
echo.
echo 📋 Next steps:
echo 1. Open Excel
echo 2. Install Script Lab add-in (Insert → Get Add-ins → Search 'Script Lab')
echo 3. In Script Lab, Import → From File → Select 'excel-addin/script-lab-proper.js'
echo 4. Click Run to start SheetMind AI!
echo.
echo 🔄 Press Ctrl+C to stop the backend when done
echo ───────────────────────────────────────────────

REM Start the FastAPI backend
uvicorn src.ui.web.app:app --host 0.0.0.0 --port 8000 --reload 