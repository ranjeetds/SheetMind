"""
Excel Context Agent for SheetMind Add-in.

Specialized agent that works with the current Excel context (selected range, worksheet)
and generates Office.js operations that can be executed directly in the Excel add-in.
"""

import json
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

import sys
from pathlib import Path

# Add src to path for imports
current_dir = Path(__file__).parent
src_dir = current_dir.parent
sys.path.insert(0, str(src_dir))

from agents.base_agent import BaseAgent
from nlp.command_processor import CommandProcessor, ExcelCommand


class ExcelContextAgent(BaseAgent):
    """
    Specialized agent for Excel add-in that works with current Excel context.
    
    Generates Office.js operations that can be executed directly in Excel.
    """
    
    def __init__(self, ai_provider: str = None):
        """Initialize the Excel context agent."""
        super().__init__(ai_provider)
        self.command_processor = CommandProcessor()
    
    async def process_query_with_context(self, query: str, excel_context: Dict) -> Dict[str, Any]:
        """
        Process a query with Excel context and return response with operations.
        
        Args:
            query: Natural language query from the user
            excel_context: Current Excel context (worksheet, selection, etc.)
            
        Returns:
            Dictionary with response and operations to execute
        """
        try:
            # Add to conversation history
            self.add_to_conversation("user", query)
            
            # Parse the command
            command = self.command_processor.process_command(query)
            
            # Analyze Excel context
            context_analysis = self._analyze_excel_context(excel_context)
            
            # Generate response and operations
            if command.confidence < 0.6:
                # Use AI to better understand the query
                ai_response = await self._get_ai_interpretation_with_context(
                    query, command, excel_context, context_analysis
                )
                response_text = ai_response
                operations = []
            else:
                # Execute the command with context
                result = await self._execute_command_with_context(command, excel_context, context_analysis)
                response_text = result["response"]
                operations = result.get("operations", [])
            
            # Add result to conversation
            self.add_to_conversation("assistant", response_text)
            
            return {
                "response": response_text,
                "operations": operations,
                "context_analysis": context_analysis
            }
            
        except Exception as e:
            error_msg = f"Sorry, I encountered an error while processing your request: {e}"
            self.add_to_conversation("assistant", error_msg)
            return {
                "response": error_msg,
                "operations": [],
                "context_analysis": {}
            }
    
    def _analyze_excel_context(self, excel_context: Dict) -> Dict[str, Any]:
        """Analyze the current Excel context."""
        if not excel_context:
            return {"has_selection": False, "message": "No Excel context available"}
        
        worksheet = excel_context.get("worksheet", {})
        selection = excel_context.get("selection", {})
        
        analysis = {
            "has_selection": True,
            "worksheet_name": worksheet.get("name", "Unknown"),
            "selection_address": selection.get("address", ""),
            "row_count": selection.get("rowCount", 0),
            "column_count": selection.get("columnCount", 0),
            "has_data": False,
            "is_numeric": False,
            "has_headers": False
        }
        
        # Analyze selection data
        values = selection.get("values", [])
        if values and len(values) > 0:
            analysis["has_data"] = True
            
            # Check if first row looks like headers
            if len(values) > 1:
                first_row = values[0] if values[0] else []
                second_row = values[1] if len(values) > 1 and values[1] else []
                
                # Simple heuristic: if first row is text and second row has numbers
                if first_row and second_row:
                    first_row_text = any(isinstance(cell, str) and cell.strip() for cell in first_row)
                    second_row_numeric = any(isinstance(cell, (int, float)) for cell in second_row)
                    analysis["has_headers"] = first_row_text and second_row_numeric
            
            # Check if data is primarily numeric
            flat_values = [cell for row in values for cell in row if cell is not None]
            numeric_count = sum(1 for cell in flat_values if isinstance(cell, (int, float)))
            analysis["is_numeric"] = numeric_count > len(flat_values) * 0.5 if flat_values else False
            analysis["numeric_percentage"] = numeric_count / len(flat_values) if flat_values else 0
        
        return analysis
    
    async def _get_ai_interpretation_with_context(
        self, 
        query: str, 
        command: ExcelCommand, 
        excel_context: Dict,
        context_analysis: Dict
    ) -> str:
        """Get AI interpretation with Excel context."""
        system_prompt = self.get_system_prompt()
        
        context_info = f"""
Current Excel Context:
- Worksheet: {context_analysis.get('worksheet_name', 'Unknown')}
- Selection: {context_analysis.get('selection_address', 'None')}
- Data size: {context_analysis.get('row_count', 0)} rows × {context_analysis.get('column_count', 0)} columns
- Has data: {context_analysis.get('has_data', False)}
- Numeric data: {context_analysis.get('is_numeric', False)}
- Has headers: {context_analysis.get('has_headers', False)}

User query: "{query}"

My initial parsing detected:
- Action: {command.action}
- Target: {command.target}  
- Parameters: {command.parameters}
- Confidence: {command.confidence:.2f}

Please provide a helpful response about what I can do with the current Excel selection.
Focus on practical operations that make sense for the current data.
"""
        
        return await self.generate_response(context_info, system_prompt)
    
    async def _execute_command_with_context(
        self, 
        command: ExcelCommand, 
        excel_context: Dict,
        context_analysis: Dict
    ) -> Dict[str, Any]:
        """Execute a command with Excel context."""
        try:
            if command.action == "calculate":
                return await self._handle_calculate_with_context(command, excel_context, context_analysis)
            elif command.action == "create":
                return await self._handle_create_with_context(command, excel_context, context_analysis)
            elif command.action == "analyze":
                return await self._handle_analyze_with_context(command, excel_context, context_analysis)
            elif command.action == "sort":
                return await self._handle_sort_with_context(command, excel_context, context_analysis)
            elif command.action == "filter":
                return await self._handle_filter_with_context(command, excel_context, context_analysis)
            else:
                return {
                    "response": f"I understand you want to {command.action}, but I need more specific instructions for this operation.",
                    "operations": []
                }
                
        except Exception as e:
            return {
                "response": f"Error executing command: {e}",
                "operations": []
            }
    
    async def _handle_calculate_with_context(
        self, 
        command: ExcelCommand, 
        excel_context: Dict,
        context_analysis: Dict
    ) -> Dict[str, Any]:
        """Handle calculation with current Excel context."""
        if not context_analysis.get("has_data"):
            return {
                "response": "Please select a range with data to perform calculations.",
                "operations": []
            }
        
        selection = excel_context.get("selection", {})
        address = selection.get("address", "")
        values = selection.get("values", [])
        
        operation = command.parameters.get("operation", "sum")
        
        # Find a good location for the result (next column or below data)
        row_count = selection.get("rowCount", 0)
        column_count = selection.get("columnCount", 0)
        
        # Parse the address to get the range
        try:
            # Simple address parsing (e.g., "A1:C10")
            start_cell, end_cell = address.split(":")
            start_col = start_cell[0]
            start_row = int(start_cell[1:])
            
            # Place result in next column
            result_col_index = ord(start_col) - ord('A') + column_count
            result_col = chr(ord('A') + result_col_index)
            result_address = f"{result_col}{start_row}"
            
            # Generate formula
            if operation == "sum":
                formula = f"=SUM({address})"
                response = f"Added SUM formula in cell {result_address}"
            elif operation == "average":
                formula = f"=AVERAGE({address})"
                response = f"Added AVERAGE formula in cell {result_address}"
            elif operation == "count":
                formula = f"=COUNT({address})"
                response = f"Added COUNT formula in cell {result_address}"
            elif operation == "max":
                formula = f"=MAX({address})"
                response = f"Added MAX formula in cell {result_address}"
            elif operation == "min":
                formula = f"=MIN({address})"
                response = f"Added MIN formula in cell {result_address}"
            else:
                return {
                    "response": f"Operation '{operation}' is not supported yet.",
                    "operations": []
                }
            
            operations = [{
                "type": "setFormula",
                "range": result_address,
                "formula": [[formula]]
            }]
            
            return {
                "response": response,
                "operations": operations
            }
            
        except Exception as e:
            return {
                "response": f"Could not create formula: {e}",
                "operations": []
            }
    
    async def _handle_create_with_context(
        self, 
        command: ExcelCommand, 
        excel_context: Dict,
        context_analysis: Dict
    ) -> Dict[str, Any]:
        """Handle create operations with Excel context."""
        if not context_analysis.get("has_data"):
            return {
                "response": "Please select a range with data to create charts or visualizations.",
                "operations": []
            }
        
        selection = excel_context.get("selection", {})
        address = selection.get("address", "")
        
        chart_type = command.parameters.get("chart_type", "bar")
        title = command.parameters.get("title", f"{chart_type.title()} Chart")
        
        # Map chart types to Excel chart types
        excel_chart_types = {
            "bar": "ColumnClustered",
            "line": "Line", 
            "pie": "Pie"
        }
        
        excel_chart_type = excel_chart_types.get(chart_type, "ColumnClustered")
        
        operations = [{
            "type": "insertChart",
            "chartType": excel_chart_type,
            "dataRange": address,
            "title": title
        }]
        
        return {
            "response": f"Created a {chart_type} chart from your selected data.",
            "operations": operations
        }
    
    async def _handle_analyze_with_context(
        self, 
        command: ExcelCommand, 
        excel_context: Dict,
        context_analysis: Dict
    ) -> Dict[str, Any]:
        """Handle analysis with Excel context."""
        if not context_analysis.get("has_data"):
            return {
                "response": "Please select a range with data to analyze.",
                "operations": []
            }
        
        selection = excel_context.get("selection", {})
        values = selection.get("values", [])
        
        # Convert to pandas DataFrame for analysis
        try:
            if context_analysis.get("has_headers"):
                headers = values[0]
                data_rows = values[1:]
                df = pd.DataFrame(data_rows, columns=headers)
            else:
                df = pd.DataFrame(values)
            
            # Perform basic analysis
            numeric_df = df.select_dtypes(include=['number'])
            
            analysis_text = f"📊 Data Analysis Summary:\n\n"
            analysis_text += f"• Data size: {len(df)} rows × {len(df.columns)} columns\n"
            
            if len(numeric_df.columns) > 0:
                analysis_text += f"• Numeric columns: {len(numeric_df.columns)}\n"
                analysis_text += f"• Average values:\n"
                
                for col in numeric_df.columns:
                    avg_val = numeric_df[col].mean()
                    analysis_text += f"  - {col}: {avg_val:.2f}\n"
                
                analysis_text += f"\n• Data ranges:\n"
                for col in numeric_df.columns:
                    min_val = numeric_df[col].min()
                    max_val = numeric_df[col].max()
                    analysis_text += f"  - {col}: {min_val:.2f} to {max_val:.2f}\n"
            else:
                analysis_text += "• No numeric data found for statistical analysis\n"
            
            return {
                "response": analysis_text,
                "operations": []
            }
            
        except Exception as e:
            return {
                "response": f"Could not analyze data: {e}",
                "operations": []
            }
    
    async def _handle_sort_with_context(
        self, 
        command: ExcelCommand, 
        excel_context: Dict,
        context_analysis: Dict
    ) -> Dict[str, Any]:
        """Handle sort operations with Excel context."""
        if not context_analysis.get("has_data"):
            return {
                "response": "Please select a range with data to sort.",
                "operations": []
            }
        
        selection = excel_context.get("selection", {})
        address = selection.get("address", "")
        
        order = command.parameters.get("order", "asc")
        ascending = order == "asc"
        
        operations = [{
            "type": "sort",
            "range": address,
            "key": 0,  # Sort by first column
            "ascending": ascending
        }]
        
        direction = "ascending" if ascending else "descending"
        return {
            "response": f"Sorted the selected data in {direction} order.",
            "operations": operations
        }
    
    async def _handle_filter_with_context(
        self, 
        command: ExcelCommand, 
        excel_context: Dict,
        context_analysis: Dict
    ) -> Dict[str, Any]:
        """Handle filter operations with Excel context."""
        return {
            "response": "Filter operations are not yet implemented for the Excel add-in, but I understand you want to filter your data.",
            "operations": []
        }
    
    async def process_query(self, query: str) -> str:
        """
        Process a query without Excel context (fallback method).
        
        Args:
            query: Natural language query from the user
            
        Returns:
            String response
        """
        # This is a fallback method - the Excel add-in should use process_query_with_context instead
        return "I need Excel context to work properly. Please use the Excel add-in interface."
    
    def get_capabilities(self) -> List[str]:
        """Return list of capabilities this agent supports."""
        return [
            "Analyze selected Excel data",
            "Create charts from current selection",
            "Perform calculations on selected ranges",
            "Sort data in current selection",
            "Generate formulas for calculations",
            "Provide data insights and summaries",
            "Work with current Excel context"
        ] 