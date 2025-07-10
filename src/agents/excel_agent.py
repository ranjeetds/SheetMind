"""
Excel Agent for SheetMind.

The main agent that combines natural language processing with Excel operations
to execute user commands on spreadsheets.
"""

import asyncio
import os
from typing import Any, Dict, List, Optional

import pandas as pd

import sys
from pathlib import Path

# Add src to path for imports
current_dir = Path(__file__).parent
src_dir = current_dir.parent
sys.path.insert(0, str(src_dir))

from agents.base_agent import BaseAgent
from integrations.excel_handler import ExcelHandler
from nlp.command_processor import CommandProcessor, ExcelCommand


class ExcelAgent(BaseAgent):
    """
    Main Excel agent that processes natural language commands and executes them.
    
    Combines AI-powered language understanding with Excel integration capabilities.
    """
    
    def __init__(self, ai_provider: str = None):
        """Initialize the Excel agent."""
        super().__init__(ai_provider)
        self.excel_handler = None
        self.command_processor = CommandProcessor()
        self.current_file = None
    
    def load_file(self, file_path: str) -> bool:
        """
        Load an Excel file for operations.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if successful
        """
        try:
            self.excel_handler = ExcelHandler(file_path)
            self.current_file = file_path
            return True
        except Exception as e:
            print(f"Error loading file: {e}")
            return False
    
    async def process_query(self, query: str) -> str:
        """
        Process a natural language query and execute the corresponding Excel operation.
        
        Args:
            query: Natural language query from the user
            
        Returns:
            Result description of the operation
        """
        try:
            # Add to conversation history
            self.add_to_conversation("user", query)
            
            # If no file is loaded and this isn't a file operation, suggest loading one
            if not self.excel_handler and not any(word in query.lower() for word in ["load", "open", "create", "new"]):
                response = "I'd be happy to help with Excel operations! However, no Excel file is currently loaded. Please either:\n1. Load an existing file: 'Load file.xlsx'\n2. Create a new file: 'Create a new workbook'"
                self.add_to_conversation("assistant", response)
                return response
            
            # Parse the command
            command = self.command_processor.process_command(query)
            
            # If confidence is low, use AI to better understand the query
            if command.confidence < 0.6:
                ai_response = await self._get_ai_interpretation(query, command)
                # Try to extract a better command from AI response
                if "I'll" in ai_response or "I will" in ai_response:
                    command = self._refine_command_from_ai(ai_response, command)
            
            # Execute the command
            result = await self._execute_command(command)
            
            # Add result to conversation
            self.add_to_conversation("assistant", result)
            
            return result
            
        except Exception as e:
            error_msg = f"Sorry, I encountered an error while processing your request: {e}"
            self.add_to_conversation("assistant", error_msg)
            return error_msg
    
    async def _get_ai_interpretation(self, query: str, command: ExcelCommand) -> str:
        """Get AI interpretation of the query for better understanding."""
        system_prompt = self.get_system_prompt()
        
        context_prompt = f"""
        User query: "{query}"
        
        My initial parsing detected:
        - Action: {command.action}
        - Target: {command.target}
        - Parameters: {command.parameters}
        - Confidence: {command.confidence:.2f}
        
        Please provide a clear interpretation of what the user wants to do with their Excel data.
        Focus on the specific Excel operation they're requesting.
        """
        
        return await self.generate_response(context_prompt, system_prompt)
    
    def _refine_command_from_ai(self, ai_response: str, original_command: ExcelCommand) -> ExcelCommand:
        """Refine the command based on AI interpretation."""
        # Try to extract a better command from the AI response
        refined_command = self.command_processor.process_command(ai_response)
        
        # If the refined command has higher confidence, use it
        if refined_command.confidence > original_command.confidence:
            return refined_command
        
        return original_command
    
    async def _execute_command(self, command: ExcelCommand) -> str:
        """Execute a parsed Excel command."""
        try:
            if command.action == "create":
                return await self._handle_create(command)
            elif command.action == "calculate":
                return await self._handle_calculate(command)
            elif command.action == "sort":
                return await self._handle_sort(command)
            elif command.action == "filter":
                return await self._handle_filter(command)
            elif command.action == "format":
                return await self._handle_format(command)
            elif command.action == "analyze":
                return await self._handle_analyze(command)
            elif command.action == "export":
                return await self._handle_export(command)
            elif command.action == "import":
                return await self._handle_import(command)
            else:
                return f"I understand you want to {command.action}, but I'm not sure how to handle that operation yet. Could you provide more specific details?"
                
        except Exception as e:
            return f"Error executing command: {e}"
    
    async def _handle_create(self, command: ExcelCommand) -> str:
        """Handle create operations (charts, pivot tables, etc.)."""
        if not self.excel_handler:
            return "Please load an Excel file first."
        
        params = command.parameters
        
        if params.get("chart_type") == "pivot":
            return "Pivot table creation is not yet implemented, but I understand you want to create a pivot table from your data."
        
        elif params.get("chart_type") in ["bar", "line", "pie"]:
            chart_type = params["chart_type"]
            title = params.get("title", f"{chart_type.title()} Chart")
            
            # Get data range - default to all data
            data = self.excel_handler.get_data()
            if data.empty:
                return "No data found to create a chart from."
            
            # Create a simple range - this is a basic implementation
            rows, cols = data.shape
            data_range = f"A1:{chr(65 + cols - 1)}{rows + 1}"
            
            success = self.excel_handler.create_chart(chart_type, data_range, title)
            
            if success:
                self.excel_handler.save()
                return f"Successfully created a {chart_type} chart titled '{title}' from your data."
            else:
                return f"Failed to create the {chart_type} chart."
        
        elif "column" in command.target or "column" in command.original_query:
            return "I understand you want to create a new column. Could you specify what data or formula should go in this column?"
        
        elif "sheet" in command.target or "worksheet" in command.original_query:
            sheet_name = params.get("sheet", "NewSheet")
            success = self.excel_handler.create_sheet(sheet_name)
            
            if success:
                self.excel_handler.save()
                return f"Successfully created a new worksheet named '{sheet_name}'."
            else:
                return f"Failed to create the worksheet '{sheet_name}'."
        
        return f"I understand you want to create something, but I need more specific details about what to create."
    
    async def _handle_calculate(self, command: ExcelCommand) -> str:
        """Handle calculation operations."""
        if not self.excel_handler:
            return "Please load an Excel file first."
        
        params = command.parameters
        operation = params.get("operation", "sum")
        
        # Get the data
        data = self.excel_handler.get_data()
        if data.empty:
            return "No data found to perform calculations on."
        
        # If a specific column is mentioned
        if "column" in params:
            col = params["column"]
            if col in data.columns:
                column_data = data[col]
            else:
                # Try to find column by index
                try:
                    col_index = ord(col.upper()) - ord('A')
                    if col_index < len(data.columns):
                        column_data = data.iloc[:, col_index]
                        col = data.columns[col_index]
                    else:
                        return f"Column {col} not found in the data."
                except:
                    return f"Invalid column reference: {col}"
        else:
            # Use the first numeric column
            numeric_cols = data.select_dtypes(include=['number']).columns
            if len(numeric_cols) == 0:
                return "No numeric columns found for calculation."
            column_data = data[numeric_cols[0]]
            col = numeric_cols[0]
        
        # Perform the calculation
        try:
            if operation == "sum":
                result = column_data.sum()
                return f"The sum of column '{col}' is {result:,.2f}"
            elif operation == "average":
                result = column_data.mean()
                return f"The average of column '{col}' is {result:,.2f}"
            elif operation == "count":
                result = column_data.count()
                return f"Column '{col}' has {result} non-empty values"
            elif operation == "max":
                result = column_data.max()
                return f"The maximum value in column '{col}' is {result:,.2f}"
            elif operation == "min":
                result = column_data.min()
                return f"The minimum value in column '{col}' is {result:,.2f}"
            else:
                return f"Calculation operation '{operation}' is not supported yet."
        except Exception as e:
            return f"Error performing calculation: {e}"
    
    async def _handle_sort(self, command: ExcelCommand) -> str:
        """Handle sort operations."""
        if not self.excel_handler:
            return "Please load an Excel file first."
        
        params = command.parameters
        order = params.get("order", "asc")
        ascending = order == "asc"
        
        # Get the data
        data = self.excel_handler.get_data()
        if data.empty:
            return "No data found to sort."
        
        # Determine sort column
        sort_column = None
        if "column" in params:
            col = params["column"]
            if col in data.columns:
                sort_column = col
            else:
                try:
                    col_index = ord(col.upper()) - ord('A')
                    if col_index < len(data.columns):
                        sort_column = data.columns[col_index]
                except:
                    pass
        
        if not sort_column:
            # Use first column by default
            sort_column = data.columns[0]
        
        try:
            # Sort the data
            sorted_data = data.sort_values(by=sort_column, ascending=ascending)
            
            # Write back to Excel
            self.excel_handler.write_data(sorted_data)
            self.excel_handler.save()
            
            direction = "ascending" if ascending else "descending"
            return f"Successfully sorted data by column '{sort_column}' in {direction} order."
            
        except Exception as e:
            return f"Error sorting data: {e}"
    
    async def _handle_filter(self, command: ExcelCommand) -> str:
        """Handle filter operations."""
        if not self.excel_handler:
            return "Please load an Excel file first."
        
        params = command.parameters
        operator = params.get("operator")
        value = params.get("value")
        
        if not operator or value is None:
            return "I need more specific filter criteria. For example: 'show rows where revenue > 1000'"
        
        # Get the data
        data = self.excel_handler.get_data()
        if data.empty:
            return "No data found to filter."
        
        # Determine filter column
        filter_column = None
        if "column" in params:
            col = params["column"]
            if col in data.columns:
                filter_column = col
        else:
            # Try to guess from common column names
            for col in data.columns:
                if any(keyword in col.lower() for keyword in ["revenue", "sales", "amount", "price", "value"]):
                    filter_column = col
                    break
        
        if not filter_column:
            filter_column = data.columns[0]  # Default to first column
        
        try:
            # Apply filter
            if operator == ">":
                filtered_data = data[data[filter_column] > value]
            elif operator == "<":
                filtered_data = data[data[filter_column] < value]
            elif operator == "=":
                filtered_data = data[data[filter_column] == value]
            elif operator == "contains":
                filtered_data = data[data[filter_column].astype(str).str.contains(str(value), case=False)]
            else:
                return f"Filter operator '{operator}' is not supported."
            
            if filtered_data.empty:
                return f"No rows found matching the filter criteria: {filter_column} {operator} {value}"
            
            # Create a summary
            count = len(filtered_data)
            return f"Found {count} rows where {filter_column} {operator} {value}. Here are the first few:\n\n{filtered_data.head().to_string()}"
            
        except Exception as e:
            return f"Error filtering data: {e}"
    
    async def _handle_format(self, command: ExcelCommand) -> str:
        """Handle format operations."""
        return "Formatting operations are not yet fully implemented, but I understand you want to format your data."
    
    async def _handle_analyze(self, command: ExcelCommand) -> str:
        """Handle analysis operations."""
        if not self.excel_handler:
            return "Please load an Excel file first."
        
        # Get the data
        data = self.excel_handler.get_data()
        if data.empty:
            return "No data found to analyze."
        
        params = command.parameters
        analysis_type = params.get("analysis_type", "summary")
        
        try:
            if analysis_type == "summary":
                # Basic data summary
                summary = data.describe()
                numeric_cols = data.select_dtypes(include=['number']).columns
                
                result = f"Data Summary:\n"
                result += f"- Total rows: {len(data)}\n"
                result += f"- Total columns: {len(data.columns)}\n"
                result += f"- Numeric columns: {len(numeric_cols)}\n\n"
                
                if len(numeric_cols) > 0:
                    result += "Statistical Summary:\n"
                    result += summary.to_string()
                
                return result
            
            elif analysis_type == "correlation":
                numeric_data = data.select_dtypes(include=['number'])
                if len(numeric_data.columns) < 2:
                    return "Need at least 2 numeric columns to calculate correlations."
                
                corr_matrix = numeric_data.corr()
                return f"Correlation Analysis:\n\n{corr_matrix.to_string()}"
            
            else:
                return f"Analysis type '{analysis_type}' is not yet implemented."
                
        except Exception as e:
            return f"Error analyzing data: {e}"
    
    async def _handle_export(self, command: ExcelCommand) -> str:
        """Handle export operations."""
        return "Export operations are not yet implemented, but I understand you want to export your data."
    
    async def _handle_import(self, command: ExcelCommand) -> str:
        """Handle import operations."""
        return "Import operations are not yet implemented, but I understand you want to import data."
    
    def get_capabilities(self) -> List[str]:
        """Return list of capabilities this agent supports."""
        return [
            "Load and analyze Excel files",
            "Create charts (bar, line, pie)",
            "Perform calculations (sum, average, count, min, max)",
            "Sort data by columns",
            "Filter data with conditions",
            "Analyze data patterns and statistics",
            "Create new worksheets",
            "Process natural language commands"
        ]
    
    def get_data_info(self) -> Dict[str, Any]:
        """Get information about the currently loaded data."""
        if not self.excel_handler:
            return {"status": "No file loaded"}
        
        data = self.excel_handler.get_data()
        if data.empty:
            return {"status": "File loaded but no data found"}
        
        return {
            "status": "Data loaded",
            "file": self.current_file,
            "rows": len(data),
            "columns": len(data.columns),
            "column_names": list(data.columns),
            "numeric_columns": list(data.select_dtypes(include=['number']).columns),
            "sheets": self.excel_handler.get_sheet_names()
        } 