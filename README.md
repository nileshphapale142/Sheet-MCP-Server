# Google Sheets MCP Server

A Model Context Protocol (MCP) server that provides tools for reading and interacting with Google Sheets data. This server enables AI assistants like Claude Desktop to seamlessly access, search, and analyze Google Spreadsheets.

Built with [uv](https://docs.astral.sh/uv/) for fast and reliable dependency management.

## 🚀 Features

- **📋 List Spreadsheets**: Discover all accessible Google Spreadsheets
- **🔍 Search by Name**: Find spreadsheets by name (exact or partial match)
- **📊 Read Sheet Data**: Extract data from specific ranges or entire sheets
- **📝 Sheet Metadata**: Get detailed information about spreadsheet structure
- **📑 List Sheets/Tabs**: View all sheets within a spreadsheet
- **🔎 Search Content**: Find specific data within sheets
- **🎯 Range Data**: Get formatted data from specific cell ranges
- **🔐 Dual Authentication**: Supports both OAuth 2.0 (private sheets) and API key (public sheets)

## 📦 Installation

### Prerequisites

- Python 3.8+
- [uv](https://docs.astral.sh/uv/) package manager
- Google Cloud Project with Sheets API enabled
- Claude Desktop or other MCP-compatible client

### Dependencies

This project uses `uv` for fast and reliable dependency management:

```bash
# Install dependencies using uv
uv sync

# Or if you don't have uv installed yet:
pip install uv
uv sync
```

**Note**: All dependencies are managed through `pyproject.toml` and `uv.lock` files.

## ⚙️ Setup

### Option 1: OAuth 2.0 Authentication (Recommended for Private Sheets)

1. **Create Google Cloud Project**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing one
   - Enable Google Sheets API and Google Drive API

2. **Create OAuth 2.0 Credentials**:
   - Go to "Credentials" → "Create Credentials" → "OAuth 2.0 Client ID"
   - Choose "Desktop application"
   - Download the `credentials.json` file
   - Place it in the project root directory

3. **Generate Authentication Token**:
   ```bash
   # Run the token generation script using uv
   uv run token_gen.py
   ```
   - A browser window will open for Google authentication
   - Grant necessary permissions
   - This will create `token.json` file for future use

4. **Environment Setup (Optional)**:
   ```bash
   # Create .env file if you want to use API key fallback
   echo "SHEET_API_KEY=your_api_key_here" > .env
   ```

### Option 2: API Key Authentication (Public Sheets Only)

1. **Create API Key**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Enable Google Sheets API
   - Go to "Credentials" → "Create Credentials" → "API Key"
   - Copy the API key

2. **Environment Configuration**:
   ```bash
   # Create .env file
   echo "SHEET_API_KEY=your_api_key_here" > .env
   ```

**Note**: API key authentication only works with publicly shared spreadsheets.

## 🏃‍♂️ Running the Server

### Prerequisites
**Important**: You must generate the authentication token before adding the MCP server to Claude Desktop.

1. **Generate Token First**:
   ```bash
   uv run token_gen.py
   ```
   Complete the OAuth flow in your browser.

2. **Then Run Server**:

### Standalone Mode
```bash
uv run main.py
```

### With Claude Desktop

**Only after completing token generation**, add to your Claude Desktop configuration:

```json
{
  "mcpServers": {
    "sheet-mcp-server": {
        "command": "uv",
        "args": [
            "--directory",
            "C:\\path\\to\\your\\SheetMCP",
            "run",
            "main.py"
        ]
    }
  }
}
```

**Replace `C:\\path\\to\\your\\SheetMCP` with your actual project directory path.**

## 🛠️ Available Tools

### 1. List Spreadsheets
Lists all Google Spreadsheets accessible to the authenticated user.

```json
{
  "name": "list_spreadsheets",
  "arguments": {
    "limit": 20,
    "order_by": "modifiedTime desc"
  }
}
```

### 2. Search Spreadsheets by Name
Find spreadsheets by name with exact or partial matching.

```json
{
  "name": "search_spreadsheets_by_name",
  "arguments": {
    "name": "Sales Report",
    "exact_match": false
  }
}
```

### 3. Read Sheet Data
Extract data from a specific range or entire sheet.

```json
{
  "name": "read_sheet_data",
  "arguments": {
    "spreadsheet_id": "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
    "range": "Sheet1!A1:C10"
  }
}
```

### 4. Get Sheet Metadata
Retrieve detailed information about a spreadsheet.

```json
{
  "name": "get_sheet_metadata",
  "arguments": {
    "spreadsheet_id": "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
  }
}
```

### 5. List Sheets
Get all sheet/tab names within a spreadsheet.

```json
{
  "name": "list_sheets",
  "arguments": {
    "spreadsheet_id": "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
  }
}
```

### 6. Search Sheet Data
Find specific content within a sheet.

```json
{
  "name": "search_sheet_data",
  "arguments": {
    "spreadsheet_id": "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
    "search_term": "Revenue",
    "sheet_name": "Q4 Data"
  }
}
```

### 7. Get Range Data
Extract data with specific formatting options.

```json
{
  "name": "get_range_data",
  "arguments": {
    "spreadsheet_id": "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
    "range": "Sheet1!A1:E20",
    "value_render_option": "FORMATTED_VALUE"
  }
}
```

## 🔐 Authentication Scopes

The server uses these Google API scopes:

- `https://www.googleapis.com/auth/spreadsheets.readonly` - Read spreadsheet content
- `https://www.googleapis.com/auth/drive.metadata.readonly` - List and discover spreadsheets

## 📁 Project Structure

```
SheetMCP/
├── main.py                 # Main MCP server implementation
├── token_gen.py           # OAuth token generation script
├── pyproject.toml         # Project configuration and dependencies
├── uv.lock                # Locked dependency versions
├── requirements.txt        # Legacy pip requirements (optional)
├── .env                   # Your environment variables (optional)
├── credentials.json       # OAuth credentials (download from Google)
├── token.json            # Generated OAuth tokens (created by token_gen.py)
└── README.md             # This file
```

## 🔧 Configuration

### Environment Variables

Create a `.env` file with:

```env
# For API key authentication (public sheets only) - Optional fallback
SHEET_API_KEY=your_api_key_here

# Server Configuration
SERVER_NAME=google-sheets-mcp-server
SERVER_VERSION=1.0.0
LOG_LEVEL=INFO
```

**Note**: The `.env` file is optional. OAuth authentication via `token_gen.py` is the primary method.

## 🐛 Troubleshooting

### Common Issues

1. **OAuth Browser Doesn't Open**:
   - Check firewall settings
   - Ensure Python has internet access
   - Try running authentication manually first

2. **Permission Denied Errors**:
   - Verify spreadsheet is accessible with your Google account
   - Check if spreadsheet is shared appropriately
   - Ensure correct API scopes are enabled

3. **API Key Limitations**:
   - API keys only work with public spreadsheets
   - Use OAuth 2.0 for private spreadsheets
   - Verify spreadsheet sharing settings

4. **Token Expiration**:
   - Delete `token.json` to force re-authentication
   - Run `uv run token_gen.py` again to regenerate
   - Check if refresh token is still valid

5. **uv Installation Issues**:
   - Install uv: `curl -LsSf https://astral.sh/uv/install.sh | sh` (Unix) or `powershell -c "irm https://astral.sh/uv/install.ps1 | iex"` (Windows)
   - Or use pip: `pip install uv`
   - Verify installation: `uv --version`

### Pre-Authentication

**Recommended workflow** to avoid OAuth interruption during use:

```bash
# Step 1: Install dependencies
uv sync

# Step 2: Generate authentication token first
uv run token_gen.py

# Step 3: Then start the MCP server or add to Claude Desktop
uv run main.py
```

This ensures authentication is completed before Claude Desktop tries to use the server.

## 📋 Example Usage with Claude

```
User: "List my Google Spreadsheets"
Claude: Uses list_spreadsheets tool to show available spreadsheets

User: "Find spreadsheets with 'Budget' in the name"
Claude: Uses search_spreadsheets_by_name tool with name="Budget"

User: "Read data from the Sales Q4 spreadsheet, range A1:E20"
Claude: Uses search + read_sheet_data tools to find and extract data

User: "What sheets are in this spreadsheet?"
Claude: Uses list_sheets tool to show all tabs/sheets
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🔗 Related Links

- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/)
- [Google Sheets API Documentation](https://developers.google.com/sheets/api)
- [Google Drive API Documentation](https://developers.google.com/drive/api)
- [Claude Desktop](https://claude.ai/desktop)

## ⚠️ Security Notes

- Keep your `credentials.json` and `token.json` files secure
- Never commit authentication files to version control
- Use environment variables for API keys
- Regularly review and rotate credentials
- Follow the principle of least privilege for API scopes