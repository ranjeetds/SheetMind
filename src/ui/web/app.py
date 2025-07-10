"""
FastAPI web application for SheetMind.

Provides a modern web interface for interacting with Excel through natural language.
"""

import os
import sys
from pathlib import Path
from typing import Dict, List, Optional

import aiofiles
from fastapi import FastAPI, File, HTTPException, Request, UploadFile, WebSocket
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

# Add the src directory to the Python path to enable imports
current_dir = Path(__file__).parent
src_dir = current_dir.parent.parent
sys.path.insert(0, str(src_dir))

from agents.excel_agent import ExcelAgent
from agents.excel_context_agent import ExcelContextAgent

# Get the directory containing this file
current_dir = Path(__file__).parent
templates_dir = current_dir / "templates"
static_dir = current_dir / "static"

# Ensure directories exist
templates_dir.mkdir(exist_ok=True)
static_dir.mkdir(exist_ok=True)

templates = Jinja2Templates(directory=str(templates_dir))


class ChatMessage(BaseModel):
    """Model for chat messages."""
    message: str
    file_path: Optional[str] = None


class ChatResponse(BaseModel):
    """Model for chat responses."""
    response: str
    data_info: Optional[Dict] = None


class ExcelContextMessage(BaseModel):
    """Model for Excel add-in messages with context."""
    message: str
    context: Optional[Dict] = None


class ExcelContextResponse(BaseModel):
    """Model for Excel add-in responses."""
    response: str
    operations: List[Dict] = []
    context_analysis: Optional[Dict] = None


def create_app() -> FastAPI:
    """Create and configure the FastAPI application."""
    app = FastAPI(
        title="SheetMind",
        description="AI-powered Excel automation tool",
        version="0.1.0"
    )
    
    # Add CORS middleware to allow Script Lab and other origins
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],  # Allow all origins for development
        allow_credentials=True,
        allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
        allow_headers=["*"],
    )
    
    # Mount static files
    if static_dir.exists():
        app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")
    
    # Global agent instances
    excel_agent = ExcelAgent()
    excel_context_agent = ExcelContextAgent()
    
    @app.get("/", response_class=HTMLResponse)
    async def home(request: Request):
        """Serve the main page."""
        return templates.TemplateResponse("index.html", {"request": request})
    
    @app.post("/chat", response_model=ChatResponse)
    async def chat(message: ChatMessage):
        """Process chat messages."""
        try:
            # If a file path is provided, load it first
            if message.file_path:
                success = excel_agent.load_file(message.file_path)
                if not success:
                    return ChatResponse(
                        response=f"Failed to load file: {message.file_path}",
                        data_info=None
                    )
            
            # Process the message
            response = await excel_agent.process_query(message.message)
            
            # Get current data info
            data_info = excel_agent.get_data_info()
            
            return ChatResponse(response=response, data_info=data_info)
            
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))
    
    @app.post("/upload")
    async def upload_file(file: UploadFile = File(...)):
        """Handle file uploads."""
        try:
            # Create uploads directory if it doesn't exist
            upload_dir = Path("uploads")
            upload_dir.mkdir(exist_ok=True)
            
            # Save the uploaded file
            file_path = upload_dir / file.filename
            
            async with aiofiles.open(file_path, 'wb') as f:
                content = await file.read()
                await f.write(content)
            
            # Load the file in the agent
            success = excel_agent.load_file(str(file_path))
            
            if success:
                data_info = excel_agent.get_data_info()
                return {
                    "success": True,
                    "message": f"Successfully loaded {file.filename}",
                    "file_path": str(file_path),
                    "data_info": data_info
                }
            else:
                return {
                    "success": False,
                    "message": f"Failed to load {file.filename}"
                }
                
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))
    
    @app.get("/data-info")
    async def get_data_info():
        """Get information about the currently loaded data."""
        return excel_agent.get_data_info()
    
    @app.get("/capabilities")
    async def get_capabilities():
        """Get list of agent capabilities."""
        return {"capabilities": excel_agent.get_capabilities()}
    
    @app.post("/chat-excel", response_model=ExcelContextResponse)
    async def chat_excel(message: ExcelContextMessage):
        """Process chat messages from Excel add-in with context."""
        try:
            # Process the message with Excel context
            result = await excel_context_agent.process_query_with_context(
                message.message, 
                message.context
            )
            
            return ExcelContextResponse(**result)
            
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))
    
    @app.websocket("/ws")
    async def websocket_endpoint(websocket: WebSocket):
        """WebSocket endpoint for real-time chat."""
        await websocket.accept()
        
        try:
            while True:
                # Receive message from client
                data = await websocket.receive_text()
                
                # Process the message
                response = await excel_agent.process_query(data)
                
                # Send response back
                await websocket.send_text(response)
                
        except Exception as e:
            await websocket.send_text(f"Error: {e}")
            await websocket.close()
    
    return app


# Create the application instance
app = create_app() 