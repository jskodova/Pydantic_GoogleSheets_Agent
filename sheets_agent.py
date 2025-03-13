import ssl
import time
from typing import Any, List, Dict
from dataclasses import dataclass
from googleapiclient.discovery import Resource
from pydantic import BaseModel, Field
from pydantic_ai import Agent, RunContext
from pydantic_ai.settings import ModelSettings
from pydantic_ai.exceptions import ModelRetry, UnexpectedModelBehavior
from google_apis import create_service
from load_models import OLLAMA_MODEL

@dataclass
class SheetsDependencies:
    sheets_service: Resource
    spreadsheet_id: str

class SheetsResult(BaseModel):
    request_status: bool = Field(description='Status of request')
    result_details: str = Field(description='Details of the request result')

def init_google_sheets_client() -> Resource:
    client_secret = 'credentials.json'
    API_NAME = 'sheets'
    API_VERSION = 'v4'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    service = create_service(client_secret, API_NAME, API_VERSION, SCOPES)
    return service

sheets_agent = Agent(
    model=OLLAMA_MODEL,
    deps_type=SheetsDependencies,
    result_type=SheetsResult,
    system_prompt=("""  
    You are a Google Sheets agent to help me manage my Google Sheets' tasks.
                   
    When making API calls, wait 1 second between each request to avoid rate limits. 
    """),
    model_settings=ModelSettings(timeout=10),
    retries=3
)

@sheets_agent.tool(retries=2)
def add_sheet(ctx: RunContext[SheetsDependencies], sheet_name: str) -> Any:
    """
    Adds a new sheet to an existing Google Spreadsheet.
    
    Args:
        sheet_name: Name for the new sheet
    
    Returns:
        Response from the API after adding the sheet.
    """
    try:
        print(f'Calling add_sheet to add sheet "{sheet_name}"')
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name
                    }
                }
            }]
        }
        response = ctx.deps.sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=ctx.deps.spreadsheet_id,
            body=request_body
        ).execute()
        time.sleep(1)
        return response

    except (Exception, UnexpectedModelBehavior) as e:
        return f'An error occurred: {str(e)}'
    
    except ssl.SSL_ERROR_SSL as e:
        raise ModelRetry(f'An error occurred: {str(e)}. Please try again.')

@sheets_agent.tool(retries=2)
def delete_sheet(ctx: RunContext[SheetsDependencies], sheet_name: str) -> Any:
    """
    Deletes a sheet from an existing Google Spreadsheet.
    
    Args:
        sheet_name: Name of the sheet to delete.
    
    Returns:
        Response from the API after deletion.
    """
    try:
        print(f'Calling delete_sheet to delete sheet "{sheet_name}"')
        sheet_metadata = ctx.deps.sheets_service.spreadsheets().get(
            spreadsheetId=ctx.deps.spreadsheet_id
        ).execute()

        sheet_id = None
        # Iterate over sheets and find the one with the matching title.
        for sheet in sheet_metadata.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                sheet_id = sheet['properties']['sheetId']
                break

        if sheet_id is None:
            print(f'Sheet {sheet_name} not found')
            return f"Sheet {sheet_name} is already deleted or does not exist"
          
        request_body = {
            'requests': [{
                'deleteSheet': {
                    'sheetId': sheet_id
                }
            }]
        }
        response = ctx.deps.sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=ctx.deps.spreadsheet_id,
            body=request_body
        ).execute()
        return response

    except (Exception, UnexpectedModelBehavior) as e:
        return f'An error occurred: {str(e)}'
    
    except ssl.SSL_ERROR_SSL as e:
        raise ModelRetry(f'An error occurred: {str(e)}. Please try again.')
    
@sheets_agent.tool(retries=2)
def list_sheets(ctx: RunContext[SheetsDependencies]) -> List[Dict[str, Any]]:
    try:
        print('Calling list_sheets')
        sheet_metadata = ctx.deps.sheets_service.spreadsheets().get(
            spreadsheetId=ctx.deps.spreadsheet_id
        ).execute()
        sheets = sheet_metadata.get('sheets', [])
        # Return an empty list if no sheets are found
        if not sheets:
            return []
        sheet_list = [{'id': sheet['properties']['sheetId'], 'name': sheet['properties']['title']} for sheet in sheets]
        return sheet_list
    except (Exception, ssl.SSL_ERROR_SSL) as e:
        raise ModelRetry(f'An error occurred: {str(e)}. Please try again.')
    
if __name__ == '__main__':
    SPREADSHEET_ID = '1PS0vkTA3spWI1AxepP3zhfztt1ps_xw2SeZVA_gjyL0'
    service = init_google_sheets_client()
    deps = SheetsDependencies(service, SPREADSHEET_ID)

    import sys

    prompt = input('User: ')
    if prompt.lower() == 'exit':
        sys.exit('See you next time')

    response = sheets_agent.run_sync(prompt.strip(), deps=deps)
    print(f'Sheet agent: {response.data}')

    while True:
        prompt = input('User: ')
        if prompt.lower() == 'exit':
            sys.exit('See you next time')
        response = sheets_agent.run_sync(prompt.strip(), deps=deps, message_history=response.all_messages())
        print(f'Sheets agent: {response.data.result_details}')
