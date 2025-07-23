from mcp.server.stdio import stdio_server
import asyncio
import logging
import os
from dotenv import load_dotenv
from mcp import ServerCapabilities, ToolsCapability, ResourcesCapability
from mcp.server.models import InitializationOptions
from mcp.server import Server
from mcp.types import (
    Resource,
    Tool
)
import mcp.types as types

from googleapiclient.discovery import build

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import json
import pickle
import os.path

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("sheets-mcp-server")

# OAuth 2.0 scopes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

class GoogleSheetsServer:
    def __init__(self):
        self.app = Server("google-sheets-mcp-server")
        self.sheets_service = None
        self.drive_service = None
        self._setup_handlers()
        
    async def authenticate_google_services(self):
        """Authenticate with Google APIs using OAuth 2.0"""
        creds = None
        
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                # Check for credentials file
                if os.path.exists('credentials.json'):
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                else:
                    # Fallback to API key if no OAuth credentials
                    api_key = os.getenv("GOOGLE_SHEETS_API_KEY")
                    if api_key:
                        logger.info("Using API key authentication (limited to public sheets)")
                        self.sheets_service = build('sheets', 'v4', developerKey=api_key)
                        return
                    else:
                        raise ValueError("Either credentials.json file or GOOGLE_SHEETS_API_KEY environment variable is required")
            
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        
        try:
            # Build services
            self.sheets_service = build('sheets', 'v4', credentials=creds)
            self.drive_service = build('drive', 'v3', credentials=creds)
            logger.info("Successfully authenticated with Google APIs using OAuth 2.0")
        except Exception as error:
            logger.error(f"An error occurred during service building: {error}")
            raise

    def _setup_handlers(self):
        @self.app.list_resources()
        async def handle_list_resources() -> list[Resource]:
            """List available Google Sheets resources"""
            return [
                Resource(
                    uri="sheets://",
                    name="Google Sheets Reader",
                    description="Access and read Google Sheets data",
                    mimeType="application/json",
                )
            ]

        @self.app.read_resource()
        async def handle_read_resource(uri: str) -> str:
            """Read a specific Google Sheets resource"""
            return json.dumps({
                "type": "google_sheets_resource",
                "description": "Google Sheets MCP Server for reading spreadsheet data",
                "capabilities": [
                    "list_spreadsheets",
                    "search_spreadsheets_by_name",
                    "read_sheet_data",
                    "get_sheet_metadata",
                    "list_sheets",
                    "get_range_data",
                    "search_sheet_data"
                ]
            })

        @self.app.list_tools()
        async def handle_list_tools() -> list[Tool]:
            """List available Google Sheets tools"""
            tools = [
                Tool(
                    name="read_sheet_data",
                    description="Read data from a Google Sheet by spreadsheet ID and range",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "The ID of the Google Spreadsheet"
                            },
                            "range": {
                                "type": "string",
                                "description": "The range to read (e.g., 'Sheet1!A1:C10')",
                                "default": "Sheet1"
                            }
                        },
                        "required": ["spreadsheet_id"]
                    },
                ),
                Tool(
                    name="get_sheet_metadata",
                    description="Get metadata about a Google Spreadsheet",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "The ID of the Google Spreadsheet"
                            }
                        },
                        "required": ["spreadsheet_id"]
                    },
                ),
                Tool(
                    name="list_sheets",
                    description="List all sheets/tabs in a Google Spreadsheet",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "The ID of the Google Spreadsheet"
                            }
                        },
                        "required": ["spreadsheet_id"]
                    },
                ),
                Tool(
                    name="search_sheet_data",
                    description="Search for specific data in a Google Sheet",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "The ID of the Google Spreadsheet"
                            },
                            "search_term": {
                                "type": "string",
                                "description": "The term to search for"
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the specific sheet to search in",
                                "default": "Sheet1"
                            }
                        },
                        "required": ["spreadsheet_id", "search_term"]
                    },
                ),
                Tool(
                    name="get_range_data",
                    description="Get data from a specific range with formatting options",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "The ID of the Google Spreadsheet"
                            },
                            "range": {
                                "type": "string",
                                "description": "The range to read (e.g., 'Sheet1!A1:C10')"
                            },
                            "value_render_option": {
                                "type": "string",
                                "description": "How to render values",
                                "enum": ["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"],
                                "default": "FORMATTED_VALUE"
                            }
                        },
                        "required": ["spreadsheet_id", "range"]
                    },
                )
            ]
            
            # Add Drive API tools only if we have OAuth credentials
            if self.drive_service:
                tools.extend([
                    Tool(
                        name="list_spreadsheets",
                        description="List all Google Spreadsheets accessible to the authenticated user",
                        inputSchema={
                            "type": "object",
                            "properties": {
                                "limit": {
                                    "type": "integer",
                                    "description": "Maximum number of spreadsheets to return",
                                    "default": 20
                                },
                                "order_by": {
                                    "type": "string",
                                    "description": "How to order results",
                                    "enum": ["name", "modifiedTime", "createdTime"],
                                    "default": "modifiedTime desc"
                                }
                            },
                            "required": []
                        },
                    ),
                    Tool(
                        name="search_spreadsheets_by_name",
                        description="Search for Google Spreadsheets by name",
                        inputSchema={
                            "type": "object",
                            "properties": {
                                "name": {
                                    "type": "string",
                                    "description": "Name or partial name of the spreadsheet to search for"
                                },
                                "exact_match": {
                                    "type": "boolean",
                                    "description": "Whether to search for exact name match",
                                    "default": False
                                }
                            },
                            "required": ["name"]
                        },
                    )
                ])
            
            return tools

        @self.app.call_tool()
        async def handle_call_tool(name: str, arguments: dict | None) -> list[types.TextContent]:
            """Handle tool calls for Google Sheets operations"""
            if not self.sheets_service:
                await self.authenticate_google_services()

            try:
                if name == "list_spreadsheets":
                    return await self._list_spreadsheets(arguments)
                elif name == "search_spreadsheets_by_name":
                    return await self._search_spreadsheets_by_name(arguments)
                elif name == "read_sheet_data":
                    return await self._read_sheet_data(arguments)
                elif name == "get_sheet_metadata":
                    return await self._get_sheet_metadata(arguments)
                elif name == "list_sheets":
                    return await self._list_sheets(arguments)
                elif name == "search_sheet_data":
                    return await self._search_sheet_data(arguments)
                elif name == "get_range_data":
                    return await self._get_range_data(arguments)
                else:
                    raise ValueError(f"Unknown tool: {name}")
            except Exception as e:
                logger.error(f"Error in {name}: {str(e)}")
                return [types.TextContent(type="text", text=f"Error: {str(e)}")]

    async def _list_spreadsheets(self, arguments: dict) -> list[types.TextContent]:
        """List all Google Spreadsheets accessible to the user"""
        if not self.drive_service:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": "Drive API not available. OAuth authentication required to list spreadsheets."
                }, indent=2)
            )]
        
        try:
            limit = arguments.get("limit", 20)
            order_by = arguments.get("order_by", "modifiedTime desc")
            
            # Call the Drive v3 API - based on Google documentation pattern
            results = (
                self.drive_service.files()
                .list(
                    q="mimeType='application/vnd.google-apps.spreadsheet'",
                    pageSize=limit,
                    orderBy=order_by,
                    fields="nextPageToken, files(id, name, modifiedTime, createdTime, owners, shared)"
                )
                .execute()
            )
            
            items = results.get('files', [])
            
            if not items:
                return [types.TextContent(
                    type="text",
                    text=json.dumps({
                        "status": "success",
                        "message": "No spreadsheets found.",
                        "count": 0,
                        "spreadsheets": []
                    }, indent=2)
                )]
            
            spreadsheets = [
                {
                    "id": item['id'],
                    "name": item['name'],
                    "modified_time": item.get('modifiedTime'),
                    "created_time": item.get('createdTime'),
                    "owners": [owner.get('displayName', owner.get('emailAddress')) for owner in item.get('owners', [])],
                    "shared": item.get('shared', False)
                }
                for item in items
            ]
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "success",
                    "count": len(spreadsheets),
                    "spreadsheets": spreadsheets,
                    "next_page_token": results.get('nextPageToken')  # For pagination
                }, indent=2)
            )]
            
        except Exception as error:
            logger.error(f"An error occurred while listing spreadsheets: {error}")
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": f"An error occurred: {error}"
                }, indent=2)
            )]

    
    
    async def _search_spreadsheets_by_name(self, arguments: dict) -> list[types.TextContent]:
        """Search for Google Spreadsheets by name"""
        if not self.drive_service:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": "Drive API not available. OAuth authentication required to search spreadsheets."
                }, indent=2)
            )]
        
        try:
            name = arguments.get("name")
            exact_match = arguments.get("exact_match", False)
            
            if exact_match:
                query = f"mimeType='application/vnd.google-apps.spreadsheet' and name='{name}'"
            else:
                query = f"mimeType='application/vnd.google-apps.spreadsheet' and name contains '{name}'"
            
            # Call the Drive v3 API - based on Google documentation pattern
            results = (
                self.drive_service.files()
                .list(
                    q=query,
                    pageSize=50,
                    fields="nextPageToken, files(id, name, modifiedTime, createdTime, owners)"
                )
                .execute()
            )
            
            items = results.get('files', [])
            
            if not items:
                return [types.TextContent(
                    type="text",
                    text=json.dumps({
                        "status": "success",
                        "search_term": name,
                        "exact_match": exact_match,
                        "message": f"No spreadsheets found matching '{name}'.",
                        "matches_found": 0,
                        "spreadsheets": []
                    }, indent=2)
                )]
            
            spreadsheets = [
                {
                    "id": item['id'],
                    "name": item['name'],
                    "modified_time": item.get('modifiedTime'),
                    "created_time": item.get('createdTime'),
                    "owners": [owner.get('displayName', owner.get('emailAddress')) for owner in item.get('owners', [])]
                }
                for item in items
            ]
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "success",
                    "search_term": name,
                    "exact_match": exact_match,
                    "matches_found": len(spreadsheets),
                    "spreadsheets": spreadsheets,
                    "next_page_token": results.get('nextPageToken')  # For pagination
                }, indent=2)
            )]
            
        except Exception as error:
            logger.error(f"An error occurred while searching spreadsheets: {error}")
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": f"An error occurred: {error}"
                }, indent=2)
            )]

    # ... (keep all the existing _read_sheet_data, _get_sheet_metadata, etc. methods)
    async def _read_sheet_data(self, arguments: dict) -> list[types.TextContent]:
        """Read data from a Google Sheet"""
        spreadsheet_id = arguments.get("spreadsheet_id")
        range_name = arguments.get("range", "Sheet1")
        
        try:
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            if not values:
                return [types.TextContent(
                    type="text",
                    text=json.dumps({
                        "status": "success",
                        "message": "No data found in the specified range",
                        "data": []
                    }, indent=2)
                )]
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "success",
                    "spreadsheet_id": spreadsheet_id,
                    "range": range_name,
                    "row_count": len(values),
                    "column_count": len(values[0]) if values else 0,
                    "data": values
                }, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": str(e)
                }, indent=2)
            )]

    async def _get_sheet_metadata(self, arguments: dict) -> list[types.TextContent]:
        """Get metadata about a Google Spreadsheet"""
        spreadsheet_id = arguments.get("spreadsheet_id")
        
        try:
            result = self.sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id
            ).execute()
            
            metadata = {
                "status": "success",
                "title": result.get('properties', {}).get('title'),
                "locale": result.get('properties', {}).get('locale'),
                "time_zone": result.get('properties', {}).get('timeZone'),
                "sheet_count": len(result.get('sheets', [])),
                "sheets": [
                    {
                        "sheet_id": sheet['properties']['sheetId'],
                        "title": sheet['properties']['title'],
                        "sheet_type": sheet['properties'].get('sheetType', 'GRID'),
                        "row_count": sheet['properties'].get('gridProperties', {}).get('rowCount'),
                        "column_count": sheet['properties'].get('gridProperties', {}).get('columnCount')
                    }
                    for sheet in result.get('sheets', [])
                ]
            }
            
            return [types.TextContent(
                type="text",
                text=json.dumps(metadata, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": str(e)
                }, indent=2)
            )]

    async def _list_sheets(self, arguments: dict) -> list[types.TextContent]:
        """List all sheets in a Google Spreadsheet"""
        spreadsheet_id = arguments.get("spreadsheet_id")
        
        try:
            result = self.sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                fields="sheets.properties"
            ).execute()
            
            sheets = [
                {
                    "sheet_id": sheet['properties']['sheetId'],
                    "title": sheet['properties']['title'],
                    "index": sheet['properties']['index'],
                    "sheet_type": sheet['properties'].get('sheetType', 'GRID')
                }
                for sheet in result.get('sheets', [])
            ]
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "success",
                    "spreadsheet_id": spreadsheet_id,
                    "sheet_count": len(sheets),
                    "sheets": sheets
                }, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": str(e)
                }, indent=2)
            )]

    async def _search_sheet_data(self, arguments: dict) -> list[types.TextContent]:
        """Search for specific data in a Google Sheet"""
        spreadsheet_id = arguments.get("spreadsheet_id")
        search_term = arguments.get("search_term")
        sheet_name = arguments.get("sheet_name", "Sheet1")
        
        try:
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=sheet_name
            ).execute()
            
            values = result.get('values', [])
            matches = []
            
            for row_idx, row in enumerate(values):
                for col_idx, cell in enumerate(row):
                    if search_term.lower() in str(cell).lower():
                        matches.append({
                            "row": row_idx + 1,
                            "column": col_idx + 1,
                            "cell_value": cell,
                            "full_row": row
                        })
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "success",
                    "search_term": search_term,
                    "sheet_name": sheet_name,
                    "matches_found": len(matches),
                    "matches": matches
                }, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": str(e)
                }, indent=2)
            )]

    async def _get_range_data(self, arguments: dict) -> list[types.TextContent]:
        """Get data from a specific range with formatting options"""
        spreadsheet_id = arguments.get("spreadsheet_id")
        range_name = arguments.get("range")
        value_render_option = arguments.get("value_render_option", "FORMATTED_VALUE")
        
        try:
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueRenderOption=value_render_option
            ).execute()
            
            values = result.get('values', [])
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "success",
                    "spreadsheet_id": spreadsheet_id,
                    "range": range_name,
                    "value_render_option": value_render_option,
                    "row_count": len(values),
                    "column_count": len(values[0]) if values else 0,
                    "data": values
                }, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "status": "error",
                    "error": str(e)
                }, indent=2)
            )]



capabilities = ServerCapabilities(
    tools=ToolsCapability(  # Use ToolsCapability instance
        listChanged=True  # Example: Supports dynamic tool list changes
        # Add other tool options as needed, e.g., subscribe=True for updates
    ),
    resources=ResourcesCapability(  # Add this to advertise resources
        listChanged=True,  # Supports dynamic resource list changes
        readable=True  # Indicates resources can be read (e.g., via handle_read_resource)
        # Add more if needed, e.g., writable=True for future expansions
    ),
    sessions=True  # Sessions can remain boolean if your SDK allows; adjust if needed
    # Add more capabilities, e.g., resources=ResourcesCapability(...)
)

async def main():
    """Run the MCP server"""
    server = GoogleSheetsServer()

    # Pre-authenticate before starting the server
    try:
        await server.authenticate_google_services()
        logger.info("Pre-authentication completed successfully")
    except Exception as e:
        logger.warning(f"Pre-authentication failed: {e}. Will authenticate on first tool use.")
    

    async with stdio_server() as (read_stream, write_stream):
        await server.app.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="powerbi-model-server",
                server_version="0.1.0",
                capabilities=capabilities
            ),
        )

if __name__ == "__main__":
    asyncio.run(main())