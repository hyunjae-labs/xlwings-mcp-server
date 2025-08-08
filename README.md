# xlwings-mcp-server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/python-%3E%3D3.10-blue)](https://www.python.org/)
[![xlwings](https://img.shields.io/badge/xlwings-%3E%3D0.30.0-green)](https://www.xlwings.org/)

A Model Context Protocol (MCP) server that manipulates Excel files using **xlwings** - providing native Excel integration through COM automation.

<a href="https://glama.ai/mcp/servers/@hyunjae-labs/xlwings-mcp-server">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@hyunjae-labs/xlwings-mcp-server/badge" alt="xlwings Excel Server MCP server" />
</a>

## 🎯 Why xlwings Instead of openpyxl?

**This MCP server is specifically designed for corporate environments where:**
- 🔒 Document security policies prevent direct file access
- 🏢 Excel files are managed by enterprise document management systems
- 📊 You need to work with Excel through official Microsoft APIs
- ✅ IT compliance requires using approved COM automation

**Key difference**: While openpyxl directly reads/writes Excel files (which may be blocked by security policies), xlwings controls Excel through Microsoft's official COM interface - the same way VBA macros work. This means if you can run Excel macros, you can use this MCP server.

## 🙏 Acknowledgments

**This project is based on [excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) by Haris Musa.**

The original excel-mcp-server uses openpyxl for Excel manipulation. This fork has been modified to use xlwings instead, which provides:
- Native Excel COM automation
- Better compatibility with complex Excel features
- Real-time Excel interaction
- Support for Excel-specific features like native pivot tables and charts

## 📋 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

The original excel-mcp-server is also MIT licensed. Copyright (c) 2025 Haris.

## 🚀 Features

All 25 tools from the original excel-mcp-server are fully functional:

### Core Excel Operations
- ✅ Create, open, save workbooks
- ✅ Manage worksheets (create, copy, rename, delete)
- ✅ Read and write data with validation
- ✅ Apply formulas and validate syntax
- ✅ Format cells and ranges

### Advanced Features
- ✅ Native Excel charts through COM
- ✅ Real pivot tables (not just data summaries)
- ✅ Excel tables (ListObjects)
- ✅ Cell merging and unmerging
- ✅ Row and column operations
- ✅ Range operations (copy, delete)
- ✅ Data validation info

## 📦 Installation

### Prerequisites
- **Python 3.10+**
- **Microsoft Excel** (required for xlwings)
- **Windows** (recommended) or macOS with Excel

### Option 1: Install from PyPI (Recommended)
```bash
pip install xlwings-mcp-server
```

### Option 2: Install from Source
1. Clone the repository:
```bash
git clone https://github.com/hyunjae-labs/xlwings-mcp-server.git
cd xlwings-mcp-server
```

2. Create virtual environment:
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# or
source .venv/bin/activate  # macOS/Linux
```

3. Install in development mode:
```bash
pip install -e .
```

## 🔧 Configuration

Add to your Claude Code MCP configuration:

```json
{
  "mcpServers": {
    "xlwings-mcp-server": {
      "type": "stdio",
      "command": "C:\\path\\to\\xlwings-mcp-server\\.venv\\Scripts\\python.exe",
      "args": ["-m", "xlwings_mcp", "stdio"]
    }
  }
}
```

## 📚 Available Tools

The server provides 25 tools for Excel manipulation:

### Workbook Operations (3)
- `create_workbook` - Create new Excel file
- `create_worksheet` - Add new worksheet
- `get_workbook_metadata` - Get workbook information

### Data Operations (5)
- `write_data_to_excel` - Write data to cells
- `read_data_from_excel` - Read cell data
- `apply_formula` - Apply Excel formulas
- `validate_formula_syntax` - Validate formula syntax
- `validate_excel_range` - Validate cell ranges

### Formatting & Visual (5)
- `format_range` - Apply cell formatting
- `create_chart` - Create Excel charts
- `create_pivot_table` - Create pivot tables
- `create_table` - Create Excel tables
- `merge_cells` - Merge cell ranges

### Sheet Management (6)
- `copy_worksheet` - Copy worksheets
- `delete_worksheet` - Delete worksheets
- `rename_worksheet` - Rename worksheets
- `unmerge_cells` - Unmerge cells
- `get_merged_cells` - Get merged cell info
- `copy_range` - Copy cell ranges

### Row/Column Operations (6)
- `delete_range` - Delete cell ranges
- `get_data_validation_info` - Get validation rules
- `insert_rows` - Insert rows
- `insert_columns` - Insert columns
- `delete_sheet_rows` - Delete rows
- `delete_sheet_columns` - Delete columns

## 🔄 When to Use Which?

### Use **excel-mcp-server** (Original) when:
- ✅ You don't have Excel installed
- ✅ You need cross-platform support
- ✅ You want faster performance
- ✅ Simple Excel operations are sufficient

### Use **xlwings-mcp-server** (This Fork) when:
- ✅ Corporate security blocks direct file access
- ✅ You need to work with protected/encrypted Excel files
- ✅ You require native Excel features (real pivot tables, complex charts)
- ✅ Your organization mandates using official Microsoft APIs
- ✅ You need real-time Excel integration

| Feature | excel-mcp-server (Original) | xlwings-mcp-server (This Fork) |
|---------|----------------------------|--------------------------------|
| **How it works** | Direct file manipulation | Controls Excel application |
| **Security Policy** | May be blocked | Works if macros are allowed |
| **Excel Required** | No | Yes |
| **Best for** | Personal use, servers | Corporate environments |

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## 📄 Citation

If you use this project, please acknowledge both:
1. The original [excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) by Haris Musa
2. This xlwings modification

## 🔗 Links

- **Original Project**: [excel-mcp-server](https://github.com/haris-musa/excel-mcp-server)
- **xlwings Documentation**: [xlwings.org](https://www.xlwings.org/)
- **MCP Protocol**: [Model Context Protocol](https://modelcontextprotocol.io/)