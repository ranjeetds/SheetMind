"""
Base agent class for SheetMind.

Provides core functionality for AI-powered agents that can understand
and execute natural language commands using local Ollama models.
"""

import os
import aiohttp
import json
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv

load_dotenv()


class BaseAgent(ABC):
    """
    Abstract base class for all SheetMind agents.
    
    Provides common functionality for AI integration using local Ollama models.
    """
    
    def __init__(self, ai_provider: str = None):
        """Initialize the base agent."""
        self.ai_provider = ai_provider or os.getenv("DEFAULT_AI_PROVIDER", "ollama")
        self.conversation_history: List[Dict[str, str]] = []
        
        # Initialize AI clients
        self._setup_ai_clients()
    
    def _setup_ai_clients(self):
        """Set up AI service clients."""
        if self.ai_provider == "ollama":
            self.ollama_url = os.getenv("OLLAMA_URL", "http://localhost:11434")
            self.model = os.getenv("OLLAMA_MODEL", "llama2")
        else:
            # Fallback to simple responses if no AI available
            self.ollama_url = None
            self.model = None
    
    async def generate_response(self, prompt: str, system_prompt: str = None) -> str:
        """
        Generate a response using the configured AI provider.
        
        Args:
            prompt: The user's input prompt
            system_prompt: Optional system prompt for context
            
        Returns:
            AI-generated response
        """
        try:
            if self.ai_provider == "ollama":
                return await self._ollama_generate(prompt, system_prompt)
            else:
                return await self._fallback_response(prompt)
        except Exception as e:
            return f"AI service unavailable: {e}. Using basic command processing."
    
    async def _ollama_generate(self, prompt: str, system_prompt: str = None) -> str:
        """Generate response using Ollama."""
        if not self.ollama_url:
            return await self._fallback_response(prompt)
        
        try:
            # Build context from conversation history and system prompt
            context = ""
            if system_prompt:
                context += f"System: {system_prompt}\n\n"
            
            # Add recent conversation history
            for msg in self.conversation_history[-5:]:  # Keep last 5 messages
                context += f"{msg['role'].capitalize()}: {msg['content']}\n"
            
            # Add current prompt
            full_prompt = f"{context}User: {prompt}\nAssistant:"
            
            # Call Ollama API
            async with aiohttp.ClientSession() as session:
                payload = {
                    "model": self.model,
                    "prompt": full_prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.2,
                        "top_p": 0.9,
                        "max_tokens": 1000
                    }
                }
                
                async with session.post(f"{self.ollama_url}/api/generate", json=payload) as response:
                    if response.status == 200:
                        result = await response.json()
                        return result.get("response", "").strip()
                    else:
                        return await self._fallback_response(prompt)
                        
        except Exception as e:
            return await self._fallback_response(prompt)
    
    async def _fallback_response(self, prompt: str) -> str:
        """Provide fallback responses when AI is not available."""
        prompt_lower = prompt.lower()
        
        # Simple pattern matching for common Excel operations
        if "sum" in prompt_lower:
            return "I can help you sum data. Select the range you want to sum and I'll add the SUM formula."
        elif "chart" in prompt_lower:
            return "I can create charts from your data. Select the data range and I'll create a chart for you."
        elif "format" in prompt_lower and "currency" in prompt_lower:
            return "I can format cells as currency. Select the cells and I'll apply currency formatting."
        elif "bold" in prompt_lower:
            return "I can make text bold. Select the cells and I'll apply bold formatting."
        elif "clear" in prompt_lower:
            return "I can clear cell contents. Select the cells you want to clear."
        elif "table" in prompt_lower:
            return "I can create formatted tables. Select your data range and I'll convert it to a table."
        elif "analyze" in prompt_lower:
            return "I can analyze your data. Select the range and I'll provide basic statistics."
        else:
            return f"I understand you want to: {prompt}. Available commands: sum, chart, format currency, bold, clear, table, analyze"
    
    def add_to_conversation(self, role: str, content: str):
        """Add a message to the conversation history."""
        self.conversation_history.append({"role": role, "content": content})
        
        # Keep conversation history manageable
        if len(self.conversation_history) > 10:
            self.conversation_history = self.conversation_history[-10:]
    
    @abstractmethod
    async def process_query(self, query: str) -> str:
        """
        Process a natural language query and execute the corresponding action.
        
        Args:
            query: Natural language query from the user
            
        Returns:
            Result of the query execution
        """
        pass
    
    @abstractmethod
    def get_capabilities(self) -> List[str]:
        """
        Return a list of capabilities this agent supports.
        
        Returns:
            List of capability descriptions
        """
        pass
    
    def get_system_prompt(self) -> str:
        """
        Get the system prompt for this agent.
        
        Returns:
            System prompt that defines the agent's role and capabilities
        """
        capabilities = self.get_capabilities()
        capabilities_text = "\n".join(f"- {cap}" for cap in capabilities)
        
        return f"""You are SheetMind, an AI assistant specialized in Excel operations.

Your capabilities include:
{capabilities_text}

Guidelines:
1. Always provide clear, actionable responses
2. When performing operations, explain what you're doing
3. If you need clarification, ask specific questions
4. Focus on practical Excel solutions
5. Be concise but thorough in explanations
6. You are running locally via Ollama for privacy and speed

Respond in a helpful, professional manner while being conversational.""" 