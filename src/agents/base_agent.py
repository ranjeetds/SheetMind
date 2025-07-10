"""
Base agent class for SheetMind.

Provides core functionality for AI-powered agents that can understand
and execute natural language commands.
"""

import os
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional

from openai import AsyncOpenAI
from anthropic import Anthropic
from dotenv import load_dotenv

load_dotenv()


class BaseAgent(ABC):
    """
    Abstract base class for all SheetMind agents.
    
    Provides common functionality for AI integration and command processing.
    """
    
    def __init__(self, ai_provider: str = None):
        """Initialize the base agent."""
        self.ai_provider = ai_provider or os.getenv("DEFAULT_AI_PROVIDER", "openai")
        self.conversation_history: List[Dict[str, str]] = []
        
        # Initialize AI clients
        self._setup_ai_clients()
    
    def _setup_ai_clients(self):
        """Set up AI service clients."""
        if self.ai_provider == "openai":
            self.openai_client = AsyncOpenAI(api_key=os.getenv("OPENAI_API_KEY"))
            self.model = os.getenv("OPENAI_MODEL", "gpt-4o")
        elif self.ai_provider == "anthropic":
            self.anthropic_client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
            self.model = os.getenv("ANTHROPIC_MODEL", "claude-3-sonnet-20240229")
        else:
            raise ValueError(f"Unsupported AI provider: {self.ai_provider}")
    
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
            if self.ai_provider == "openai":
                return await self._openai_generate(prompt, system_prompt)
            elif self.ai_provider == "anthropic":
                return await self._anthropic_generate(prompt, system_prompt)
        except Exception as e:
            raise Exception(f"AI generation failed: {e}")
    
    async def _openai_generate(self, prompt: str, system_prompt: str = None) -> str:
        """Generate response using OpenAI."""
        messages = []
        
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        
        # Add conversation history
        messages.extend(self.conversation_history[-10:])  # Keep last 10 messages
        
        # Add current prompt
        messages.append({"role": "user", "content": prompt})
        
        response = await self.openai_client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=0.1,
            max_tokens=2000
        )
        
        return response.choices[0].message.content
    
    async def _anthropic_generate(self, prompt: str, system_prompt: str = None) -> str:
        """Generate response using Anthropic."""
        # Build context from conversation history
        context = ""
        for msg in self.conversation_history[-10:]:
            context += f"{msg['role']}: {msg['content']}\n"
        
        full_prompt = f"{system_prompt}\n\n{context}\nuser: {prompt}\nassistant:"
        
        response = await self.anthropic_client.completions.create(
            model=self.model,
            prompt=full_prompt,
            max_tokens_to_sample=2000,
            temperature=0.1
        )
        
        return response.completion
    
    def add_to_conversation(self, role: str, content: str):
        """Add a message to the conversation history."""
        self.conversation_history.append({"role": role, "content": content})
        
        # Keep conversation history manageable
        if len(self.conversation_history) > 20:
            self.conversation_history = self.conversation_history[-20:]
    
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

Respond in a helpful, professional manner while being conversational.""" 