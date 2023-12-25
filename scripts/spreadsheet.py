from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from scripts.spreadsheet_utils import get_header_formats, get_page_id, get_entry_format, delete_sheet, summarize_data, clear_formatting_first_page, HEADERS, SERVICE_ACCOUNT_FILE, SCOPES, SPREADSHEET_ID


def get_data_array(plays: dict) -> list:
    games = []
    for _, game in plays.items():
        game_list = [
            int(game['GAME_NUMBER']),
            game['home_team']['Team'],
            game['home_team']['Rk'],
            float(game['home_team_score']),
            game['away_team']['Team'],
            game['away_team']['Rk'],
            float(game['away_team_score']),
            float(game['spread']),
            game['book']
        ]
        games.append(game_list)

    return games


def format_headers(sheet_id: str) -> None:
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=credentials)

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': get_header_formats(sheet_id)}
    ).execute()


def add_entries(date: str, sheet_id: str, plays: dict) -> int:
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=credentials)
    sheet_name = date
    data = get_data_array(plays)

    values = [HEADERS] + data
    if data is not None:
        body = {
            'values': values
        }
        range = f'{sheet_name}!A1'
        request = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                         range=range,
                                                         valueInputOption='RAW',
                                                         body=body)
        request.execute()
        format_requests = get_entry_format(sheet_id, len(values))
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'requests': format_requests}
        ).execute()

        return len(values)
    return None


def create_page(sheet_name: str) -> None:
    # Load credentials
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    # Request body to create a new sheet
    body = {
        'requests': [{
            'addSheet': {
                'properties': {
                    'title': sheet_name
                }
            }
        }]
    }
    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body=body).execute()
    except Exception as e:
        print(f"An error occurred: {e}")
        return None


def add_formulas(page_id: str, num_rows: int) -> None:
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=credentials)

    requests = []
    for row in range(2, num_rows + 1):  # Assuming row 1 is the header row
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': page_id,
                    'startRowIndex': row - 1,
                    'endRowIndex': row,
                    'startColumnIndex': 9,  # Column J
                    'endColumnIndex': 10
                },
                'cell': {
                    'userEnteredValue': {
                        'formulaValue': f'=C{row}-F{row}'
                    }
                },
                'fields': 'userEnteredValue'
            }
        })
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': page_id,
                    'startRowIndex': row - 1,
                    'endRowIndex': row,
                    'startColumnIndex': 10,  # Column K
                    'endColumnIndex': 11
                },
                'cell': {
                    'userEnteredValue': {
                        'formulaValue': f'=IF(AND(D{row}=0, G{row}=0), "N/A", D{row}>G{row})'
                    }

                },
                'fields': 'userEnteredValue'
            }
        })
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': page_id,
                    'startRowIndex': row - 1,
                    'endRowIndex': row,
                    'startColumnIndex': 11,  # Column L
                    'endColumnIndex': 12
                },
                'cell': {
                    'userEnteredValue': {
                        'formulaValue': f'=IF(AND(D{row}=0, G{row}=0), "N/A", D{row}-G{row})'
                    }
                },
                'fields': 'userEnteredValue'
            }
        })
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': page_id,
                    'startRowIndex': row - 1,
                    'endRowIndex': row,
                    'startColumnIndex': 12,  # Column M
                    'endColumnIndex': 13
                },
                'cell': {
                    'userEnteredValue': {
                        'formulaValue': (f'=IF(AND(D{row}=0, G{row}=0), "N/A", '
                                         f'IF(L{row} > 0, L{row} > ABS(H{row}), FALSE))')
                    }
                },
                'fields': 'userEnteredValue'
            }
        })
    try:
        # Apply the batch update with all the formula requests
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'requests': requests}
        ).execute()
    except:
        pass


def populate_spreadsheet(date: str, plays: dict) -> None:
    if get_page_id(date) is not None:
        delete_sheet(get_page_id(date))
    create_page(date)
    clear_formatting_first_page()
    page_id = get_page_id(date)
    num_games = add_entries(date, page_id, plays)
    if num_games is not None:
        format_headers(page_id)
        add_formulas(page_id, num_games)
    summarize_data()
