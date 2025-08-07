<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server Logo" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-mcp-server.svg)](https://pypi.org/project/excel-mcp-server/)
[![Total Downloads](https://static.pepy.tech/badge/excel-mcp-server)](https://pepy.tech/project/excel-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![smithery badge](https://smithery.ai/badge/@haris-musa/excel-mcp-server)](https://smithery.ai/server/@haris-musa/excel-mcp-server)
[![Install MCP Server](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-mcp-server&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLW1jcC1zZXJ2ZXIgc3RkaW8ifQ%3D%3D)

A Model Context Protocol (MCP) server that lets you manipulate Excel files without needing Microsoft Excel installed. Create, read, and modify Excel workbooks with your AI agent.

## Features

- 📊 **Excel Operations**: Create, read, update workbooks and worksheets
- 📈 **Data Manipulation**: Formulas, formatting, charts, pivot tables, and Excel tables
- 🔍 **Data Validation**: Built-in validation for ranges, formulas, and data integrity
- 🎨 **Formatting**: Font styling, colors, borders, alignment, and conditional formatting
- 📋 **Table Operations**: Create and manage Excel tables with custom styling
- 📊 **Chart Creation**: Generate various chart types (line, bar, pie, scatter, etc.)
- 🔄 **Pivot Tables**: Create dynamic pivot tables for data analysis
- 🔧 **Sheet Management**: Copy, rename, delete worksheets with ease
- 🔌 **Triple transport support**: stdio, SSE (deprecated), and streamable HTTP
- 🌐 **Remote & Local**: Works both locally and as a remote service

## Installation

### Prerequisites

- **Python 3.10+** or **uv** (recommended)
- **Windows**: Microsoft Excel or xlwings-compatible environment
- **Claude Code**: MCP server support enabled

### Quick Setup with uv (Recommended)

1. **Install uv** (if not already installed):
   ```bash
   # Windows PowerShell
   irm https://astral.sh/uv/install.ps1 | iex
   
   # macOS/Linux
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

2. **Configure Claude Code**:
   
   Edit your Claude Code config file:
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
   - **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Linux**: `~/.config/Claude/claude_desktop_config.json`

   Add the following configuration:
   ```json
   {
      "mcpServers": {
         "xlwings-mcp-server": {
            "command": "uv",
            "args": [
               "run",
               "--directory",
               "[YOUR_PROJECT_PATH]/mcp-servers/xlwings-mcp-server",
               "python",
               "-m",
               "excel_mcp",
               "stdio"
            ]
         }
      }
   }
   ```
   
   Replace `[YOUR_PROJECT_PATH]` with your actual project directory.

3. **Restart Claude Code** to apply the configuration.

### Alternative: Python Setup

If you prefer using Python directly:

```bash
cd [YOUR_PROJECT_PATH]/mcp-servers/xlwings-mcp-server
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux
pip install -e .
```

Then configure Claude Code with:
```json
{
   "mcpServers": {
      "xlwings-mcp-server": {
         "command": "[YOUR_PROJECT_PATH]/mcp-servers/xlwings-mcp-server/.venv/Scripts/python",
         "args": ["-m", "excel_mcp", "stdio"]
      }
   }
}
```

## Usage

The server supports three transport methods:

### 1. Stdio Transport (for local use)

```bash
# Using uv
uv run --directory [YOUR_PROJECT_PATH]/mcp-servers/xlwings-mcp-server python -m excel_mcp stdio

# Using Python directly
python -m excel_mcp stdio
```

### 2. SSE Transport (Server-Sent Events - Deprecated)

```bash
python -m excel_mcp sse
```

Configure with `"url": "http://localhost:8000/sse"`

### 3. Streamable HTTP Transport (Recommended for remote connections)

```bash
python -m excel_mcp streamable-http
```

Configure with `"url": "http://localhost:8000/mcp"`

## Environment Variables & File Path Handling

### SSE and Streamable HTTP Transports

When running the server with the **SSE or Streamable HTTP protocols**, you **must set the `EXCEL_FILES_PATH` environment variable on the server side**. This variable tells the server where to read and write Excel files.
- If not set, it defaults to `./excel_files`.

You can also set the `FASTMCP_PORT` environment variable to control the port the server listens on (default is `8000` if not set).
- Example (Windows PowerShell):
  ```powershell
  $env:EXCEL_FILES_PATH="E:\MyExcelFiles"
  $env:FASTMCP_PORT="8007"
  uvx excel-mcp-server streamable-http
  ```
- Example (Linux/macOS):
  ```bash
  EXCEL_FILES_PATH=/path/to/excel_files FASTMCP_PORT=8007 uvx excel-mcp-server streamable-http
  ```

### Stdio Transport

When using the **stdio protocol**, the file path is provided with each tool call, so you do **not** need to set `EXCEL_FILES_PATH` on the server. The server will use the path sent by the client for each operation.

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. See [TOOLS.md](TOOLS.md) for complete documentation of all available tools.

### Quick Verification

Test your installation:
```bash
# Test with uv
uv run --directory [YOUR_PROJECT_PATH]/mcp-servers/xlwings-mcp-server python -m excel_mcp --help

# Test with Python
python -m excel_mcp --help
```

In Claude Code, try:
```
"Create an Excel file at C:\test.xlsx"
"Write 'Hello World' to cell A1"
"Read the data from Sheet1"
```

## Troubleshooting

- **"MCP server not found"**: Check your path in Claude Code config
- **"xlwings import error"**: Run `pip install xlwings`
- **Excel connection issues on Windows**: Ensure Microsoft Excel is installed

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=haris-musa/excel-mcp-server&type=Date)](https://www.star-history.com/#haris-musa/excel-mcp-server&Date)

## License

MIT License - see [LICENSE](LICENSE) for details.
