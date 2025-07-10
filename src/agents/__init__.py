"""
Agents package for SheetMind.

Contains AI agents that can understand and execute Excel operations.
"""

import sys
from pathlib import Path

# Add src to path for imports
current_dir = Path(__file__).parent
src_dir = current_dir.parent
sys.path.insert(0, str(src_dir))

from agents.base_agent import BaseAgent
from agents.excel_agent import ExcelAgent

__all__ = ["BaseAgent", "ExcelAgent"] 