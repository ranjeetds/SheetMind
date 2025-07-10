"""
Command processor for SheetMind.

Analyzes natural language commands and converts them into structured
Excel operations that can be executed by the Excel agent.
"""

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Union

import pandas as pd


@dataclass
class ExcelCommand:
    """Represents a structured Excel command."""
    action: str  # The type of action to perform
    target: str  # What to act on (e.g., "column A", "sheet1", "data")
    parameters: Dict[str, Any]  # Additional parameters for the action
    confidence: float  # Confidence in the parsing (0.0 to 1.0)
    original_query: str  # Original natural language query


class CommandProcessor:
    """
    Processes natural language commands and converts them to Excel operations.
    
    Uses pattern matching and keyword analysis to understand user intent.
    """
    
    def __init__(self):
        """Initialize the command processor with pattern definitions."""
        self.action_patterns = self._define_action_patterns()
        self.data_patterns = self._define_data_patterns()
        self.location_patterns = self._define_location_patterns()
    
    def _define_action_patterns(self) -> Dict[str, List[str]]:
        """Define patterns for different types of actions."""
        return {
            "create": [
                r"create|make|add|generate|build|new",
                r"pivot table|chart|graph|plot|table|column|row|sheet|worksheet"
            ],
            "calculate": [
                r"calculate|compute|sum|average|count|total|find",
                r"formula|function|equation|operation"
            ],
            "sort": [
                r"sort|order|arrange|organize",
                r"ascending|descending|asc|desc|alphabetical|numerical"
            ],
            "filter": [
                r"filter|find|search|show|display|where",
                r"greater than|less than|equal|contains|matches|>|<|="
            ],
            "format": [
                r"format|style|color|font|bold|italic",
                r"currency|percentage|date|number|text"
            ],
            "analyze": [
                r"analyze|analysis|correlation|trend|pattern|insight",
                r"statistics|stats|summary|report"
            ],
            "export": [
                r"export|save|download|output",
                r"csv|pdf|image|file"
            ],
            "import": [
                r"import|load|open|read",
                r"file|data|csv|excel"
            ]
        }
    
    def _define_data_patterns(self) -> Dict[str, List[str]]:
        """Define patterns for data references."""
        return {
            "column": [
                r"column [A-Z]+|col [A-Z]+",
                r"column \w+|col \w+"
            ],
            "row": [
                r"row \d+",
                r"rows? \d+-\d+"
            ],
            "range": [
                r"[A-Z]+\d+:[A-Z]+\d+",
                r"range [A-Z]+\d+:[A-Z]+\d+"
            ],
            "sheet": [
                r"sheet \w+|worksheet \w+",
                r"tab \w+"
            ],
            "data": [
                r"data|dataset|table|spreadsheet",
                r"all data|entire data|whole data"
            ]
        }
    
    def _define_location_patterns(self) -> Dict[str, List[str]]:
        """Define patterns for location references."""
        return {
            "cell": [
                r"[A-Z]+\d+",
                r"cell [A-Z]+\d+"
            ],
            "column_ref": [
                r"column [A-Z]",
                r"col [A-Z]"
            ],
            "row_ref": [
                r"row \d+",
                r"line \d+"
            ]
        }
    
    def process_command(self, query: str) -> ExcelCommand:
        """
        Process a natural language command into a structured Excel command.
        
        Args:
            query: Natural language query from the user
            
        Returns:
            ExcelCommand object with parsed information
        """
        query = query.lower().strip()
        
        # Detect the primary action
        action = self._detect_action(query)
        
        # Detect the target (what to act on)
        target = self._detect_target(query)
        
        # Extract parameters based on action
        parameters = self._extract_parameters(query, action)
        
        # Calculate confidence based on pattern matches
        confidence = self._calculate_confidence(query, action, target, parameters)
        
        return ExcelCommand(
            action=action,
            target=target,
            parameters=parameters,
            confidence=confidence,
            original_query=query
        )
    
    def _detect_action(self, query: str) -> str:
        """Detect the primary action from the query."""
        scores = {}
        
        for action, patterns in self.action_patterns.items():
            score = 0
            for pattern in patterns:
                matches = len(re.findall(pattern, query, re.IGNORECASE))
                score += matches
            scores[action] = score
        
        # Return action with highest score, default to "analyze"
        if not scores or max(scores.values()) == 0:
            return "analyze"
        
        return max(scores, key=scores.get)
    
    def _detect_target(self, query: str) -> str:
        """Detect what the action should target."""
        for target_type, patterns in self.data_patterns.items():
            for pattern in patterns:
                match = re.search(pattern, query, re.IGNORECASE)
                if match:
                    return match.group(0)
        
        # Default to "data" if no specific target found
        return "data"
    
    def _extract_parameters(self, query: str, action: str) -> Dict[str, Any]:
        """Extract action-specific parameters from the query."""
        params = {}
        
        if action == "create":
            params.update(self._extract_create_params(query))
        elif action == "calculate":
            params.update(self._extract_calculate_params(query))
        elif action == "sort":
            params.update(self._extract_sort_params(query))
        elif action == "filter":
            params.update(self._extract_filter_params(query))
        elif action == "format":
            params.update(self._extract_format_params(query))
        elif action == "analyze":
            params.update(self._extract_analyze_params(query))
        
        # Common parameters
        params.update(self._extract_common_params(query))
        
        return params
    
    def _extract_create_params(self, query: str) -> Dict[str, Any]:
        """Extract parameters for create operations."""
        params = {}
        
        if "pivot table" in query:
            params["chart_type"] = "pivot"
        elif any(word in query for word in ["chart", "graph", "plot"]):
            # Detect chart type
            if "bar" in query:
                params["chart_type"] = "bar"
            elif "line" in query:
                params["chart_type"] = "line"
            elif "pie" in query:
                params["chart_type"] = "pie"
            else:
                params["chart_type"] = "bar"  # Default
        
        # Extract title if mentioned
        title_match = re.search(r"(?:title|name|call it) ['\"]([^'\"]+)['\"]", query)
        if title_match:
            params["title"] = title_match.group(1)
        
        return params
    
    def _extract_calculate_params(self, query: str) -> Dict[str, Any]:
        """Extract parameters for calculation operations."""
        params = {}
        
        # Detect operation type
        if any(word in query for word in ["sum", "total", "add"]):
            params["operation"] = "sum"
        elif any(word in query for word in ["average", "mean", "avg"]):
            params["operation"] = "average"
        elif "count" in query:
            params["operation"] = "count"
        elif any(word in query for word in ["max", "maximum", "highest"]):
            params["operation"] = "max"
        elif any(word in query for word in ["min", "minimum", "lowest"]):
            params["operation"] = "min"
        
        # Extract percentage if mentioned
        percentage_match = re.search(r"(\d+(?:\.\d+)?)%", query)
        if percentage_match:
            params["percentage"] = float(percentage_match.group(1))
        
        return params
    
    def _extract_sort_params(self, query: str) -> Dict[str, Any]:
        """Extract parameters for sort operations."""
        params = {}
        
        if any(word in query for word in ["descending", "desc", "high to low", "largest first"]):
            params["order"] = "desc"
        else:
            params["order"] = "asc"  # Default to ascending
        
        return params
    
    def _extract_filter_params(self, query: str) -> Dict[str, Any]:
        """Extract parameters for filter operations."""
        params = {}
        
        # Extract comparison operators and values
        gt_match = re.search(r"(?:greater than|>)\s*(\d+(?:\.\d+)?)", query)
        if gt_match:
            params["operator"] = ">"
            params["value"] = float(gt_match.group(1))
        
        lt_match = re.search(r"(?:less than|<)\s*(\d+(?:\.\d+)?)", query)
        if lt_match:
            params["operator"] = "<"
            params["value"] = float(lt_match.group(1))
        
        eq_match = re.search(r"(?:equal to?|=)\s*(['\"]?)([^'\"]+)\1", query)
        if eq_match:
            params["operator"] = "="
            params["value"] = eq_match.group(2)
        
        contains_match = re.search(r"contains?\s+['\"]([^'\"]+)['\"]", query)
        if contains_match:
            params["operator"] = "contains"
            params["value"] = contains_match.group(1)
        
        return params
    
    def _extract_format_params(self, query: str) -> Dict[str, Any]:
        """Extract parameters for format operations."""
        params = {}
        
        if "currency" in query:
            params["format_type"] = "currency"
        elif "percentage" in query:
            params["format_type"] = "percentage"
        elif "date" in query:
            params["format_type"] = "date"
        elif "bold" in query:
            params["format_type"] = "bold"
        elif "italic" in query:
            params["format_type"] = "italic"
        
        # Extract color if mentioned
        color_match = re.search(r"color[:\s]+(\w+)", query)
        if color_match:
            params["color"] = color_match.group(1)
        
        return params
    
    def _extract_analyze_params(self, query: str) -> Dict[str, Any]:
        """Extract parameters for analysis operations."""
        params = {}
        
        if "correlation" in query:
            params["analysis_type"] = "correlation"
        elif any(word in query for word in ["trend", "pattern"]):
            params["analysis_type"] = "trend"
        elif any(word in query for word in ["summary", "overview"]):
            params["analysis_type"] = "summary"
        elif "statistics" in query or "stats" in query:
            params["analysis_type"] = "statistics"
        
        return params
    
    def _extract_common_params(self, query: str) -> Dict[str, Any]:
        """Extract common parameters that apply to multiple actions."""
        params = {}
        
        # Extract column references
        col_match = re.search(r"column ([A-Z])", query, re.IGNORECASE)
        if col_match:
            params["column"] = col_match.group(1).upper()
        
        # Extract range references
        range_match = re.search(r"([A-Z]+\d+:[A-Z]+\d+)", query, re.IGNORECASE)
        if range_match:
            params["range"] = range_match.group(1).upper()
        
        # Extract sheet references
        sheet_match = re.search(r"(?:sheet|worksheet)\s+(\w+)", query, re.IGNORECASE)
        if sheet_match:
            params["sheet"] = sheet_match.group(1)
        
        return params
    
    def _calculate_confidence(self, query: str, action: str, target: str, parameters: Dict[str, Any]) -> float:
        """Calculate confidence score for the parsed command."""
        confidence = 0.0
        
        # Base confidence from action detection
        if action != "analyze":  # Default action
            confidence += 0.3
        
        # Confidence from target detection
        if target != "data":  # Default target
            confidence += 0.2
        
        # Confidence from parameter extraction
        if parameters:
            confidence += min(0.4, len(parameters) * 0.1)
        
        # Bonus for specific Excel terminology
        excel_terms = ["formula", "cell", "range", "chart", "pivot", "sheet"]
        for term in excel_terms:
            if term in query:
                confidence += 0.1
        
        return min(1.0, confidence)
    
    def get_command_suggestions(self, partial_query: str) -> List[str]:
        """
        Get command suggestions based on partial input.
        
        Args:
            partial_query: Partial user input
            
        Returns:
            List of suggested completions
        """
        suggestions = []
        
        common_commands = [
            "Create a pivot table from the data",
            "Add a formula to calculate tax in column D",
            "Sort the data by date ascending",
            "Filter rows where revenue > 1000",
            "Format column C as currency",
            "Create a bar chart from sales data",
            "Calculate the average of column B",
            "Find the top 10 customers by revenue",
            "Export data as CSV",
            "Create a new worksheet"
        ]
        
        # Filter suggestions based on partial input
        if partial_query:
            query_lower = partial_query.lower()
            suggestions = [cmd for cmd in common_commands 
                          if query_lower in cmd.lower()]
        else:
            suggestions = common_commands[:5]  # Show top 5 if no input
        
        return suggestions 