#!/usr/bin/env python3
"""
SheetMind - Main application entry point.

Provides both web interface and CLI for interacting with Excel through natural language.
"""

import asyncio
import os
import sys
from pathlib import Path

import click
import uvicorn
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from ui.web.app import create_app
from agents.excel_agent import ExcelAgent

console = Console()

# Load environment variables
load_dotenv()


@click.group()
@click.version_option(version="0.1.0")
def cli():
    """SheetMind - AI-powered Excel automation tool."""
    console.print(Panel.fit(
        "[bold blue]🧠 SheetMind[/bold blue]\n"
        "[dim]AI-powered Excel automation[/dim]",
        border_style="blue"
    ))


@cli.command()
@click.option("--host", default="0.0.0.0", help="Host address")
@click.option("--port", default=8000, help="Port number")
@click.option("--reload", is_flag=True, help="Enable auto-reload")
def web(host: str, port: int, reload: bool):
    """Start the web interface."""
    console.print(f"🚀 Starting SheetMind web interface at http://{host}:{port}")
    
    app = create_app()
    uvicorn.run(
        app,
        host=host,
        port=port,
        reload=reload
    )


@cli.command()
@click.option("--file", "-f", help="Excel file to work with")
def chat(file: str):
    """Start interactive chat mode with Excel."""
    console.print("💬 Starting SheetMind chat mode...")
    
    if file and not os.path.exists(file):
        console.print(f"[red]Error: File '{file}' not found[/red]")
        return
    
    agent = ExcelAgent()
    
    if file:
        console.print(f"📊 Loaded Excel file: {file}")
        agent.load_file(file)
    
    console.print("\n[dim]Type 'exit' to quit, 'help' for commands[/dim]\n")
    
    while True:
        try:
            query = console.input("[bold green]SheetMind>[/bold green] ")
            
            if query.lower() in ['exit', 'quit']:
                break
            elif query.lower() == 'help':
                show_help()
                continue
            elif query.strip() == '':
                continue
            
            # Process the query
            with console.status("[bold blue]Processing...[/bold blue]"):
                result = asyncio.run(agent.process_query(query))
            
            console.print(f"[bold cyan]Result:[/bold cyan] {result}")
            
        except KeyboardInterrupt:
            break
        except Exception as e:
            console.print(f"[red]Error: {e}[/red]")
    
    console.print("\n👋 Goodbye!")


def show_help():
    """Show help information."""
    help_text = """
[bold]Available Commands:[/bold]

[cyan]Basic Operations:[/cyan]
• "Create a pivot table from sales data"
• "Add a formula to calculate tax in column D"
• "Sort data by date ascending"
• "Find rows where revenue > 1000"

[cyan]Data Analysis:[/cyan]
• "Show me the top 10 customers by revenue"
• "Calculate the average sales per month"
• "Find correlations in the data"

[cyan]Formatting:[/cyan]
• "Format currency columns"
• "Add conditional formatting for negative values"
• "Create a chart from this data"

[cyan]File Operations:[/cyan]
• "Save as CSV"
• "Export chart as image"
• "Create a new worksheet"

[dim]Type 'exit' to quit[/dim]
    """
    console.print(Panel(help_text, title="SheetMind Help", border_style="cyan"))


@cli.command()
def version():
    """Show version information."""
    console.print("SheetMind v0.1.0")


if __name__ == "__main__":
    cli() 