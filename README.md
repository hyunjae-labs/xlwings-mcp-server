# xlwings-mcp-server

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/yourusername/xlwings-mcp-server)
[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-Compatible-orange.svg)](https://modelcontextprotocol.io)

A robust Model Context Protocol (MCP) server for Excel automation using xlwings. This server provides comprehensive Excel file manipulation capabilities through a session-based architecture, designed for high-performance and reliable Excel operations.

## 🚀 Features

### Core Capabilities
- **Session-based Architecture**: Persistent Excel workbook sessions for optimal performance
- **Comprehensive Excel Operations**: Full support for data manipulation, formulas, formatting, and visualization
- **Thread-safe Operations**: Concurrent access with per-session locking
- **Automatic Resource Management**: TTL-based session cleanup and LRU eviction policies
- **Zero-Error Design**: Katherine Johnson principle compliance with comprehensive error handling

### Excel Operations
- **Workbook Management**: Open, create, list, and close Excel workbooks
- **Worksheet Operations**: Create, copy, rename, and delete worksheets
- **Data Manipulation**: Read, write, and modify Excel data with full type support
- **Formula Support**: Apply and validate Excel formulas with syntax checking
- **Advanced Formatting**: Cell styling, conditional formatting, and range formatting
- **Visualization**: Chart creation with multiple chart types
- **Table Operations**: Native Excel table creation and management
- **Range Operations**: Cell merging, copying, and deletion

## 🛠️ Installation

### Prerequisites
- Python 3.10 or higher
- Windows OS (required for xlwings COM integration)
- Microsoft Excel installed

### Using pip
```bash
pip install xlwings-mcp-server
```

### From Source
```bash
git clone https://github.com/yourusername/xlwings-mcp-server.git
cd xlwings-mcp-server
pip install -e .
```

### Using uv (Recommended)
```bash
uv add xlwings-mcp-server
```

## ⚡ Quick Start

### 1. Basic Usage

Start the MCP server:
```bash
xlwings-mcp-server
```

Or run directly:
```bash
python -m xlwings_mcp
```

### 2. Session-based Workflow

```python
# Example using MCP client
import mcp

# Open a workbook session
session_result = client.call_tool("mcp__xlwings-mcp-server__open_workbook", {
    "filepath": "C:/path/to/your/file.xlsx",
    "visible": False,
    "read_only": False
})

session_id = session_result["session_id"]

# Write data
client.call_tool("mcp__xlwings-mcp-server__write_data_to_excel", {
    "session_id": session_id,
    "sheet_name": "Sheet1",
    "data": [["Name", "Age", "Score"], ["Alice", 25, 95], ["Bob", 30, 87]]
})

# Apply formulas
client.call_tool("mcp__xlwings-mcp-server__apply_formula", {
    "session_id": session_id,
    "sheet_name": "Sheet1",
    "cell": "D2",
    "formula": "=B2+C2"
})

# Create chart
client.call_tool("mcp__xlwings-mcp-server__create_chart", {
    "session_id": session_id,
    "sheet_name": "Sheet1",
    "data_range": "A1:C3",
    "chart_type": "column",
    "target_cell": "E1"
})

# Close session
client.call_tool("mcp__xlwings-mcp-server__close_workbook", {
    "session_id": session_id
})
```

## 🔧 Configuration

### Environment Variables
```bash
# Session management
EXCEL_MCP_SESSION_TTL=600          # Session TTL in seconds (default: 600)
EXCEL_MCP_MAX_SESSIONS=8           # Maximum concurrent sessions (default: 8)
EXCEL_MCP_DEBUG_LOG=1              # Enable debug logging (default: 0)

# Excel settings
EXCEL_MCP_VISIBLE=false            # Show Excel windows (default: false)
EXCEL_MCP_CALC_MODE=automatic      # Calculation mode (default: automatic)
```

### MCP Configuration (.mcp.json)
```json
{
  "name": "xlwings-mcp-server",
  "version": "1.0.0",
  "transport": {
    "type": "stdio"
  },
  "tools": {
    "prefix": "mcp__xlwings-mcp-server__"
  }
}
```

## 📚 API Reference

### Session Management
- `open_workbook(filepath, visible=False, read_only=False)`: Create new session
- `close_workbook(session_id)`: Close session and save workbook  
- `list_workbooks()`: List active sessions
- `force_close_workbook_by_path(filepath)`: Force close by file path

### Data Operations
- `write_data_to_excel(session_id, sheet_name, data, start_cell=None)`
- `read_data_from_excel(session_id, sheet_name, start_cell=None, end_cell=None)`
- `apply_formula(session_id, sheet_name, cell, formula)`
- `validate_formula_syntax(session_id, sheet_name, cell, formula)`

### Worksheet Management
- `create_worksheet(session_id, sheet_name)`
- `copy_worksheet(session_id, source_sheet, target_sheet)`
- `rename_worksheet(session_id, old_name, new_name)`
- `delete_worksheet(session_id, sheet_name)`

### Formatting & Visualization
- `format_range(session_id, sheet_name, start_cell, **formatting_options)`
- `create_chart(session_id, sheet_name, data_range, chart_type, target_cell)`
- `create_table(session_id, sheet_name, data_range, table_name=None)`

### Range Operations
- `merge_cells(session_id, sheet_name, start_cell, end_cell)`
- `unmerge_cells(session_id, sheet_name, start_cell, end_cell)`
- `copy_range(session_id, sheet_name, source_start, source_end, target_start)`
- `delete_range(session_id, sheet_name, start_cell, end_cell)`

## 🏗️ Architecture

### Session-based Design
The server implements a sophisticated session management system:

- **ExcelSessionManager**: Singleton pattern managing all Excel sessions
- **Per-session Isolation**: Each session has independent Excel Application instance
- **Thread Safety**: RLock per session preventing concurrent access issues
- **Resource Management**: Automatic cleanup with TTL and LRU policies
- **Error Recovery**: Comprehensive error handling and session recovery

### Performance Optimizations
- **Session Reuse**: Eliminates Excel restart overhead between operations
- **Connection Pooling**: Efficient COM object management
- **Batch Operations**: Optimized for multiple operations on same workbook
- **Memory Management**: Proactive cleanup of Excel processes

## 🧪 Testing

### Run Tests
```bash
# Run all tests
python -m pytest test/

# Run specific test categories  
python -m pytest test/test_session.py      # Session management
python -m pytest test/test_functions.py   # MCP function tests
python -m pytest test/test_integration.py # Integration tests
```

### Test Coverage
The project maintains 100% test coverage for:
- All MCP tool functions (17 functions tested)
- Session lifecycle management
- Error handling and recovery
- Performance benchmarks

## 🔒 Security Considerations

- **File System Access**: Server operates within specified directory permissions
- **Excel Process Isolation**: Each session runs in separate Excel instance
- **Resource Limits**: Configurable session limits prevent resource exhaustion
- **Input Validation**: All inputs validated before Excel API calls
- **Safe Defaults**: Read-only mode available, invisible Excel instances by default

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Setup
```bash
git clone https://github.com/yourusername/xlwings-mcp-server.git
cd xlwings-mcp-server
uv venv
uv sync
uv run python -m xlwings_mcp
```

## 📝 Changelog

See [CHANGELOG.md](CHANGELOG.md) for detailed version history.

## 🐛 Troubleshooting

### Common Issues

**Excel COM Error**: Ensure Excel is properly installed and not running in safe mode
```bash
# Check Excel installation
excel --version
```

**Session Not Found**: Verify session hasn't expired (default TTL: 10 minutes)
```bash
# List active sessions
client.call_tool("mcp__xlwings-mcp-server__list_workbooks")
```

**Permission Denied**: Run with appropriate file system permissions
```bash
# Windows: Run as administrator if needed
```

### Debug Mode
Enable detailed logging:
```bash
export EXCEL_MCP_DEBUG_LOG=1
xlwings-mcp-server
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [xlwings](https://www.xlwings.org/) - Excellent Python-Excel integration library
- [Model Context Protocol](https://modelcontextprotocol.io) - Standardized AI-tool communication
- [Claude Code](https://claude.ai/code) - Development assistance
- Katherine Johnson - Inspiration for zero-error engineering principles

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/xlwings-mcp-server/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/xlwings-mcp-server/discussions)
- **Email**: haris.musa@outlook.com

---

**Made with ❤️ for the Excel automation community**