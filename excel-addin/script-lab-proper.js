document.getElementById("run").addEventListener("click", () => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
    // Initialize SheetMind AI interface
    initializeSheetMindUI();
    
    console.log("🧠 SheetMind AI initialized with Ollama support!");
    await context.sync();
  });
}

function initializeSheetMindUI() {
  // Remove existing SheetMind container if it exists
  const existingContainer = document.getElementById('sheetmind-container');
  if (existingContainer) {
    existingContainer.remove();
  }
  
  // Create our UI container
  const container = document.createElement('div');
  container.id = 'sheetmind-container';
  container.innerHTML = `
    <div style="font-family: 'Segoe UI', sans-serif; padding: 10px;">
      <h2 style="color: #0078d4; margin-bottom: 15px;">🧠 SheetMind AI (Local)</h2>
      
      <div id="status" style="background: #d4edda; border: 1px solid #c3e6cb; border-radius: 5px; padding: 10px; margin-bottom: 15px; text-align: center; color: #155724;">
        ✅ Connected to Excel! Using local Ollama AI
      </div>
      
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 15px;">
        <button id="sum-btn" class="ms-Button ms-Button--primary" style="padding: 8px; font-size: 12px;">📊 Sum</button>
        <button id="currency-btn" class="ms-Button ms-Button--primary" style="padding: 8px; font-size: 12px;">💰 Currency</button>
        <button id="chart-btn" class="ms-Button ms-Button--primary" style="padding: 8px; font-size: 12px;">📈 Chart</button>
        <button id="bold-btn" class="ms-Button ms-Button--primary" style="padding: 8px; font-size: 12px;">🔸 Bold</button>
      </div>
      
      <div id="chat-container" style="background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 5px; padding: 10px; height: 200px; overflow-y: auto; margin-bottom: 10px; font-size: 14px;">
        <div style="background: #e3f2fd; padding: 8px; border-radius: 4px; margin-bottom: 8px;">
          👋 Welcome to SheetMind AI with local Ollama!<br>
          <small>Try: "sum column A", "analyze this data", "create a chart", "format as table", or any natural language command</small>
        </div>
      </div>
      
      <div style="display: flex; gap: 8px; margin-bottom: 10px;">
        <input type="text" id="user-input" placeholder="Ask me anything about your Excel data..." 
               style="flex: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px;">
        <button id="send-btn" class="ms-Button ms-Button--primary" style="padding: 8px 15px;">Send</button>
      </div>
      
      <div id="excel-info" style="background: #f1f3f4; border-radius: 4px; padding: 8px; font-size: 12px;">
        <strong>Context:</strong> <span id="context-info">Loading...</span>
      </div>
      
      <div style="margin-top: 10px; font-size: 11px; color: #666;">
        <strong>AI Status:</strong> <span id="ai-status">Checking Ollama connection...</span>
      </div>
    </div>
  `;
  
  // Find where to insert our UI (after the existing content)
  const existingContent = document.querySelector('#content-main') || document.body;
  existingContent.appendChild(container);
  
  // Set up event listeners
  setupEventListeners();
  
  // Update Excel context
  updateExcelContext();
  
  // Check AI backend status
  checkAIStatus();
  
  // Update context every 3 seconds
  setInterval(updateExcelContext, 3000);
}

function setupEventListeners() {
  document.getElementById('send-btn').addEventListener('click', () => {
    const input = document.getElementById('user-input');
    const message = input.value.trim();
    if (message) {
      sendMessage(message);
      input.value = '';
    }
  });
  
  document.getElementById('user-input').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      const message = e.target.value.trim();
      if (message) {
        sendMessage(message);
        e.target.value = '';
      }
    }
  });
  
  // Quick action buttons
  document.getElementById('sum-btn').addEventListener('click', () => sendMessage('sum the selected column'));
  document.getElementById('currency-btn').addEventListener('click', () => sendMessage('format selected cells as currency'));
  document.getElementById('chart-btn').addEventListener('click', () => sendMessage('create a chart from this data'));
  document.getElementById('bold-btn').addEventListener('click', () => sendMessage('make the selected cells bold'));
}

async function checkAIStatus() {
  const statusElement = document.getElementById('ai-status');
  try {
    // Try to reach the backend
    const response = await fetch('http://localhost:8000/capabilities');
    if (response.ok) {
      statusElement.innerHTML = '🟢 Local AI Backend Connected';
      statusElement.style.color = '#28a745';
    } else {
      statusElement.innerHTML = '🟡 Backend available, AI may be limited';
      statusElement.style.color = '#ffc107';
    }
  } catch (error) {
    statusElement.innerHTML = '🔴 Backend not running (using basic commands only)';
    statusElement.style.color = '#dc3545';
  }
}

function addMessage(text, type) {
  const chatContainer = document.getElementById('chat-container');
  const messageDiv = document.createElement('div');
  messageDiv.style.cssText = `
    margin-bottom: 8px; 
    padding: 8px; 
    border-radius: 4px; 
    font-size: 13px;
    ${type === 'user' ? 'background: #bbdefb; text-align: right;' : 'background: #f3e5f5;'}
  `;
  messageDiv.innerHTML = text;
  chatContainer.appendChild(messageDiv);
  chatContainer.scrollTop = chatContainer.scrollHeight;
}

async function sendMessage(message) {
  addMessage(message, 'user');
  addMessage('🤔 Processing...', 'ai');
  
  try {
    // First try AI-powered processing
    const aiResponse = await tryAIProcessing(message);
    if (aiResponse) {
      addMessage(`🤖 ${aiResponse}`, 'ai');
      return;
    }
    
    // Fall back to basic command processing
    await processCommand(message);
  } catch (error) {
    addMessage(`❌ Error: ${error.message}`, 'ai');
  }
  
  updateExcelContext();
}

async function tryAIProcessing(message) {
  try {
    // Get current Excel context with better error handling
    let excelContext = {};
    
    await Excel.run(async (context) => {
      try {
        const range = context.workbook.getSelectedRange();
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Load all necessary properties
        range.load("address, values, rowCount, columnCount, text, formulas");
        worksheet.load("name");
        
        await context.sync();
        
        // Get the actual values (not just formulas)
        const values = range.values || [];
        const text = range.text || [];
        
        // Limit data size for API call but ensure we have meaningful data
        const limitedValues = values.slice(0, 10).map(row => 
          Array.isArray(row) ? row.slice(0, 10) : [row]
        );
        
        const limitedText = text.slice(0, 10).map(row => 
          Array.isArray(row) ? row.slice(0, 10) : [row]
        );
        
        excelContext = {
          worksheet: {
            name: worksheet.name
          },
          selection: {
            address: range.address,
            values: limitedValues,
            text: limitedText,
            rowCount: range.rowCount,
            columnCount: range.columnCount
          }
        };
        
        // Debug log to see what we're capturing
        console.log('Excel context captured:', {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          hasValues: values.length > 0,
          firstValue: values[0] ? values[0][0] : 'no data'
        });
        
      } catch (error) {
        console.error('Error reading Excel context:', error);
        // Fallback context
        excelContext = {
          worksheet: { name: 'Unknown' },
          selection: {
            address: 'A1',
            values: [],
            text: [],
            rowCount: 0,
            columnCount: 0
          }
        };
      }
    });
    
    // Only proceed if we have valid context
    if (!excelContext.selection || excelContext.selection.rowCount === 0) {
      addMessage('📍 Please select some cells in Excel first, then try your command again.', 'ai');
      return 'Please select cells in Excel to analyze or work with.';
    }
    
    // Call the AI backend with proper context
    const response = await fetch('http://localhost:8000/chat-excel', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        message: message,
        context: excelContext
      })
    });
    
    if (response.ok) {
      const result = await response.json();
      
      // Execute any Excel operations returned by the AI
      if (result.operations && result.operations.length > 0) {
        await executeAIOperations(result.operations);
      }
      
      return result.response;
    } else {
      console.error('API response not ok:', response.status, response.statusText);
      return null;
    }
    
  } catch (error) {
    console.log('AI processing failed, falling back to basic commands:', error);
    return null;
  }
}

async function executeAIOperations(operations) {
  for (const operation of operations) {
    try {
      await Excel.run(async (context) => {
        // Execute operations based on type
        // This would be expanded based on the AI's operation format
        console.log('Executing AI operation:', operation);
      });
    } catch (error) {
      addMessage(`❌ Error executing AI operation: ${error.message}`, 'ai');
    }
  }
}

async function processCommand(command) {
  const lowerCommand = command.toLowerCase();
  
  if (lowerCommand.includes('sum')) {
    await sumColumn();
  } else if (lowerCommand.includes('currency')) {
    await formatAsCurrency();
  } else if (lowerCommand.includes('clear')) {
    await clearSelection();
  } else if (lowerCommand.includes('bold')) {
    await formatBold();
  } else if (lowerCommand.includes('chart')) {
    await createChart();
  } else if (lowerCommand.includes('table')) {
    await createTable();
  } else if (lowerCommand.includes('analyze')) {
    await analyzeData();
  } else if (lowerCommand.includes('sort')) {
    await sortData();
  } else if (lowerCommand.includes('highlight') || lowerCommand.includes('color')) {
    await highlightCells();
  } else if (lowerCommand.includes('freeze')) {
    await freezePanes();
  } else {
    addMessage('✅ Available commands: sum, currency, chart, bold, clear, table, analyze, sort, highlight, freeze panes', 'ai');
  }
}

async function sumColumn() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, rowCount");
    await context.sync();
    
    const sumCell = range.getCell(range.rowCount, 0);
    sumCell.formulas = [["=SUM(" + range.address + ")"]];
    
    await context.sync();
    addMessage(`✅ Added SUM formula for ${range.address}`, 'ai');
  });
}

async function formatAsCurrency() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.numberFormat = [["$#,##0.00"]];
    
    await context.sync();
    addMessage('✅ Formatted selection as currency', 'ai');
  });
}

async function clearSelection() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.clear();
    
    await context.sync();
    addMessage('✅ Cleared selected cells', 'ai');
  });
}

async function formatBold() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.font.bold = true;
    
    await context.sync();
    addMessage('✅ Made selection bold', 'ai');
  });
}

async function createChart() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const chart = context.workbook.worksheets.getActiveWorksheet().charts.add("ColumnClustered", range, "Auto");
    chart.title.text = "SheetMind Chart";
    
    await context.sync();
    addMessage('✅ Created chart from selected data', 'ai');
  });
}

async function createTable() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const table = context.workbook.tables.add(range, true);
    table.name = "SheetMindTable";
    table.getHeaderRowRange().format.fill.color = "#0078d4";
    table.getHeaderRowRange().format.font.color = "white";
    
    await context.sync();
    addMessage('✅ Created formatted table', 'ai');
  });
}

async function sortData() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("columnCount");
    await context.sync();
    
    // Sort by first column
    const sortFields = [{
      key: 0,
      ascending: true
    }];
    
    range.sort.apply(sortFields);
    
    await context.sync();
    addMessage('✅ Sorted data by first column', 'ai');
  });
}

async function highlightCells() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.fill.color = "#FFFF00"; // Yellow highlight
    
    await context.sync();
    addMessage('✅ Highlighted selected cells', 'ai');
  });
}

async function freezePanes() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const worksheet = range.worksheet;
    worksheet.freezePanes.freezeAt(range);
    
    await context.sync();
    addMessage('✅ Froze panes at selected cell', 'ai');
  });
}

async function analyzeData() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values, rowCount, columnCount");
    await context.sync();
    
    const rows = range.rowCount;
    const cols = range.columnCount;
    const values = range.values;
    
    // Basic analysis
    let numericCount = 0;
    let nonEmptyCount = 0;
    
    for (let row of values) {
      for (let cell of row) {
        if (cell !== null && cell !== '') {
          nonEmptyCount++;
          if (typeof cell === 'number') {
            numericCount++;
          }
        }
      }
    }
    
    const analysis = `📊 Data Analysis:<br>
• ${rows} rows × ${cols} columns<br>
• ${nonEmptyCount} non-empty cells<br>
• ${numericCount} numeric values (${Math.round(numericCount/nonEmptyCount*100)}%)<br>
• Suggestion: ${numericCount > nonEmptyCount * 0.5 ? 'Try creating a chart or calculating sums' : 'Consider formatting as a table'}`;
    
    addMessage(analysis, 'ai');
  });
}

async function updateExcelContext() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      
      range.load("address, rowCount, columnCount, values");
      worksheet.load("name");
      await context.sync();
      
      const contextInfo = document.getElementById('context-info');
      if (contextInfo) {
        const hasData = range.values && range.values.length > 0 && 
                       range.values.some(row => row.some(cell => cell !== null && cell !== ''));
        
        contextInfo.innerHTML = `
          📊 ${worksheet.name} | ${range.address} (${range.rowCount}×${range.columnCount})
          ${hasData ? '✅ Has data' : '⚪ Empty selection'}
        `;
      }
    });
  } catch (error) {
    const contextInfo = document.getElementById('context-info');
    if (contextInfo) {
      contextInfo.innerHTML = '❌ Unable to read Excel selection';
    }
    console.error('Error updating context:', error);
  }
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
    // Show error in UI if possible
    const chatContainer = document.getElementById('chat-container');
    if (chatContainer) {
      const errorDiv = document.createElement('div');
      errorDiv.style.cssText = 'background: #f8d7da; color: #721c24; padding: 8px; border-radius: 4px; margin-bottom: 8px;';
      errorDiv.innerHTML = `❌ Error: ${error.message}`;
      chatContainer.appendChild(errorDiv);
    }
  }
} 