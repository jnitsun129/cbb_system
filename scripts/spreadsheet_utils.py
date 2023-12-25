from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from scripts.utils import num_pages, get_spreadsheet_id

HEADERS = ['GameNo', 'Home Team', 'Rk',
           'Points', 'Away Team', 'Rk', 'Points', 'Spread', 'Book', 'Rank Diff', 'ML', 'Margin', 'Cover']

SERVICE_ACCOUNT_FILE = './utility_files/google_account.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

CURR_PAGE_ID = num_pages()

SPREADSHEET_ID = get_spreadsheet_id()


def get_page_id(sheet_name):
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=credentials)

    spreadsheet_metadata = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID).execute()

    for sheet in spreadsheet_metadata.get('sheets', []):
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']

    return None


def delete_sheet(sheet_id):
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=credentials)

    requests = [{
        'deleteSheet': {
            'sheetId': sheet_id
        }
    }]

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body={'requests': requests}).execute()


def get_header_formats(sheet_id):
    return [
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': len(HEADERS)
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'RIGHT',
                        'textFormat': {
                            'bold': True,
                            'foregroundColor': {
                                'red': 1.0,
                                'green': 0,
                                'blue': 0
                            }
                        }
                    }
                },
                'fields': 'userEnteredFormat(horizontalAlignment,textFormat)'
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'stringValue': 'AVG Rank Diff ML:'
                        },
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True
                            },
                            'horizontalAlignment': 'LEFT'
                        }
                    }]
                }],
                'fields': 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment)',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 14,
                    'endColumnIndex': 15
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'formulaValue': '=IFERROR(AVERAGEIF(K:K, TRUE, J:J), "N/A")'
                        }
                    }]
                }],
                'fields': 'userEnteredValue',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 15,
                    'endColumnIndex': 16
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'stringValue': 'AVG Rank Diff SPREAD:'
                        },
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True
                            },
                            'horizontalAlignment': 'LEFT'
                        }
                    }]
                }],
                'fields': 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment)',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,
                    'endRowIndex': 2,
                    'startColumnIndex': 14,
                    'endColumnIndex': 15
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'formulaValue': '=IFERROR(AVERAGEIF(M:M, TRUE, J:J), "N/A")'
                        }
                    }]
                }],
                'fields': 'userEnteredValue',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,
                    'endRowIndex': 2,
                    'startColumnIndex': 15,
                    'endColumnIndex': 16
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'stringValue': 'ML Wins:'
                        },
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True
                            },
                            'horizontalAlignment': 'LEFT'
                        }
                    }]
                }],
                'fields': 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment)',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 2,
                    'endRowIndex': 3,
                    'startColumnIndex': 14,
                    'endColumnIndex': 15
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'formulaValue': '=COUNTIF(K:K, TRUE)'
                        },
                        'userEnteredFormat': {
                            'horizontalAlignment': 'LEFT'
                        }
                    }]
                }],
                'fields': 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment)',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 2,
                    'endRowIndex': 3,
                    'startColumnIndex': 15,
                    'endColumnIndex': 16
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'stringValue': 'Spread Wins:'
                        },
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True
                            },
                            'horizontalAlignment': 'LEFT'
                        }
                    }]
                }],
                'fields': 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment)',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 3,
                    'endRowIndex': 4,
                    'startColumnIndex': 14,
                    'endColumnIndex': 15
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'formulaValue': '=COUNTIF(M:M, TRUE)'
                        },
                        'userEnteredFormat': {
                            'horizontalAlignment': 'LEFT'
                        }
                    }]
                }],
                'fields': 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment)',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 3,
                    'endRowIndex': 4,
                    'startColumnIndex': 15,
                    'endColumnIndex': 16
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'stringValue': 'ML Record:'
                        },
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True
                            },
                            'horizontalAlignment': 'LEFT'
                        }
                    }]
                }],
                'fields': 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment)',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 4,
                    'endRowIndex': 5,
                    'startColumnIndex': 14,
                    'endColumnIndex': 15
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'formulaValue': '=COUNTIF(K:K, TRUE) & "-" & (MAX(A:A) - COUNTIF(K:K, TRUE))'
                        }
                    }]
                }],
                'fields': 'userEnteredValue',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 4,
                    'endRowIndex': 5,
                    'startColumnIndex': 15,
                    'endColumnIndex': 16
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'stringValue': 'Spread Record:'
                        },
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True
                            },
                            'horizontalAlignment': 'LEFT'
                        }
                    }]
                }],
                'fields': 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment)',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 5,
                    'endRowIndex': 6,
                    'startColumnIndex': 14,
                    'endColumnIndex': 15
                }
            }
        },
        {
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'formulaValue': '=COUNTIF(M:M, TRUE) & "-" & (MAX(A:A) - COUNTIF(M:M, TRUE))'
                        }
                    }]
                }],
                'fields': 'userEnteredValue',
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 5,
                    'endRowIndex': 6,
                    'startColumnIndex': 15,
                    'endColumnIndex': 16
                }
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 14,
                    'endIndex': 15
                },
                'properties': {
                    'pixelSize': 175
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 2,
                    'endIndex': 3
                },
                'properties': {
                    'pixelSize': 50
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 3,
                    'endIndex': 4
                },
                'properties': {
                    'pixelSize': 50
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 6,
                    'endIndex': 7
                },
                'properties': {
                    'pixelSize': 50
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 7,
                    'endIndex': 8
                },
                'properties': {
                    'pixelSize': 50
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 9,
                    'endIndex': 10
                },
                'properties': {
                    'pixelSize': 60
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 11,
                    'endIndex': 12
                },
                'properties': {
                    'pixelSize': 50
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 5,
                    'endIndex': 6
                },
                'properties': {
                    'pixelSize': 50
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': 1
                },
                'properties': {
                    'pixelSize': 75
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 8,
                    'endIndex': 9
                },
                'properties': {
                    'pixelSize': 75
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 10,
                    'endIndex': 11
                },
                'properties': {
                    'pixelSize': 50
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 12,
                    'endIndex': 13
                },
                'properties': {
                    'pixelSize': 75
                },
                'fields': 'pixelSize'
            }
        }
    ]


def get_entry_format(sheet_id, num_games):
    return [
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1 + num_games,
                    'startColumnIndex': 0,
                    'endColumnIndex': len(HEADERS)
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'RIGHT',
                    }
                },
                'fields': 'userEnteredFormat(horizontalAlignment)'
            }
        },
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{
                        'sheetId': sheet_id,
                        'startRowIndex': 1,
                        'endRowIndex': 1 + num_games,
                        'startColumnIndex': 0,
                        'endColumnIndex': len(HEADERS)
                    }],
                    'booleanRule': {
                        'condition': {
                            'type': 'CUSTOM_FORMULA',
                            'values': [{'userEnteredValue': '=ISEVEN(ROW())'}]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.9,
                                'green': 0.9,
                                'blue': 0.9
                            }
                        }
                    }
                },
                'index': 0
            }
        }
    ]


def format_main_sheet(row_index):
    row_index += 1
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    sheet_metadata = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID).execute()
    first_sheet_id = sheet_metadata['sheets'][0]['properties']['sheetId']

    batch_update_body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {
                        'sheetId': first_sheet_id,
                        'startRowIndex': 0,
                        'endRowIndex': 1,
                        'endColumnIndex': 6
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {'bold': True, 'fontSize': 12},
                            'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
                            'horizontalAlignment': 'CENTER'
                        }
                    },
                    'fields': 'userEnteredFormat(textFormat, backgroundColor, horizontalAlignment)'
                }
            },
            {
                'updateBorders': {
                    'range': {
                        'sheetId': first_sheet_id,
                        'startRowIndex': 1,
                        'endRowIndex': row_index,
                        'startColumnIndex': 0,
                        'endColumnIndex': 6
                    },
                    'top': {'style': 'SOLID', 'width': 1},
                    'bottom': {'style': 'SOLID', 'width': 1},
                    'left': {'style': 'SOLID', 'width': 1},
                    'right': {'style': 'SOLID', 'width': 1},
                    'innerHorizontal': {'style': 'SOLID', 'width': 1},
                    'innerVertical': {'style': 'SOLID', 'width': 1}
                }
            },
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': first_sheet_id,
                        'dimension': 'COLUMNS',
                        'startIndex': 0,
                        'endIndex': 6
                    },
                    'properties': {'pixelSize': 120},
                    'fields': 'pixelSize'
                }
            },
            {
                'repeatCell': {
                    'range': {
                        'sheetId': first_sheet_id,
                        'startRowIndex': row_index - 1,
                        'endRowIndex': row_index,
                        'startColumnIndex': 0,
                        'endColumnIndex': 6
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {'red': 0.8, 'green': 0.8, 'blue': 1.0}
                        }
                    },
                    'fields': 'userEnteredFormat.backgroundColor'
                }
            },
            {
                'repeatCell': {
                    'range': {
                        'sheetId': first_sheet_id,
                        'startRowIndex': 1,
                        'endRowIndex': 3,
                        'startColumnIndex': 7,
                        'endColumnIndex': 9
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER',
                            'textFormat': {'bold': True},
                            'borders': {
                                'top': {'style': 'SOLID', 'width': 1},
                                'bottom': {'style': 'SOLID', 'width': 1},
                                'left': {'style': 'SOLID', 'width': 1},
                                'right': {'style': 'SOLID', 'width': 1}
                            }
                        }
                    },
                    'fields': 'userEnteredFormat(horizontalAlignment, textFormat, borders)'
                }
            },
        ]
    }
    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body=batch_update_body
        ).execute()
    except Exception as e:
        print(f"An error occurred: {e}")


def clear_formatting_first_page():
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=credentials)

    sheet_metadata = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID).execute()
    first_sheet_id = sheet_metadata['sheets'][0]['properties']['sheetId']

    requests = [{
        "updateCells": {
            "range": {
                "sheetId": first_sheet_id
            },
            "fields": "userEnteredFormat"
        }
    }]

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests}
        ).execute()
    except Exception as e:
        print(f"An error occurred: {e}")


def summarize_data():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet_metadata = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID).execute()
    sheets = sheet_metadata.get('sheets', '')
    first_sheet_id = sheets[0]['properties']['sheetId']
    headers = ['Date', '# Games', 'ML Wins',
               'Spread Wins', 'Avg Diff ML', 'Avg Diff Spread']
    batch_update_body = {
        'requests': [
            {
                'updateCells': {
                    'range': {
                        'sheetId': first_sheet_id,
                        'startRowIndex': 0,
                        'endRowIndex': 1,
                        'startColumnIndex': 0,
                        'endColumnIndex': len(headers)
                    },
                    'rows': [{
                        'values': [{'userEnteredValue': {'stringValue': h}, 'userEnteredFormat': {'textFormat': {'bold': True}}} for h in headers]
                    }],
                    'fields': 'userEnteredValue,userEnteredFormat.textFormat.bold'
                }
            },
            {
                'repeatCell': {
                    'range': {
                        'sheetId': first_sheet_id,
                        'startRowIndex': 0,
                        'startColumnIndex': 0,
                        'endColumnIndex': 1
                    },
                    'cell': {
                        'userEnteredFormat': {'horizontalAlignment': 'LEFT'}
                    },
                    'fields': 'userEnteredFormat.horizontalAlignment'
                }
            },
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': first_sheet_id,
                        'dimension': 'COLUMNS',
                        'startIndex': 0,
                        'endIndex': len(headers)
                    },
                    'properties': {
                        'pixelSize': 120
                    },
                    'fields': 'pixelSize'
                }
            }
        ]
    }

    row_index = 1
    for sheet in sheets[1:]:
        sheet_name = sheet['properties']['title']
        formulas = [
            sheet_name,
            f"=MAX('{sheet_name}'!A:A)",
            f"='{sheet_name}'!P3",
            f"='{sheet_name}'!P4",
            f"='{sheet_name}'!P1",
            f"='{sheet_name}'!P2"
        ]
        batch_update_body['requests'].append({
            'updateCells': {
                'range': {
                    'sheetId': first_sheet_id,
                    'startRowIndex': row_index,
                    'endRowIndex': row_index + 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': len(headers)
                },
                'rows': [{
                    'values': [
                        {
                            'userEnteredValue': {'stringValue' if i == 0 else 'formulaValue': formula},
                            'userEnteredFormat': {'horizontalAlignment': 'LEFT'}
                        } for i, formula in enumerate(formulas)
                    ]
                }],
                'fields': 'userEnteredValue, userEnteredFormat.horizontalAlignment'
            }
        })

        row_index += 1

    total_row_index = row_index
    total_formulas = [
        "Totals:",
        "=SUM(B2:B" + str(row_index) + ")",
        "=SUM(C2:C" + str(row_index) + ")",
        "=SUM(D2:D" + str(row_index) + ")",
        "=AVERAGE(E2:E" + str(row_index) + ")",
        "=AVERAGE(F2:F" + str(row_index) + ")"
    ]
    batch_update_body['requests'].append({
        'updateCells': {
            'range': {
                'sheetId': first_sheet_id,
                'startRowIndex': total_row_index,
                'endRowIndex': total_row_index + 1,
                'startColumnIndex': 0,
                'endColumnIndex': len(headers)
            },
            'rows': [{
                'values': [{'userEnteredValue': {'stringValue' if i == 0 else 'formulaValue': formula}} for i, formula in enumerate(total_formulas)]
            }],
            'fields': 'userEnteredValue'
        }
    })

    batch_update_body['requests'].append({
        'updateCells': {
            'range': {
                'sheetId': first_sheet_id,
                'startRowIndex': 1,
                'endRowIndex': 2,
                'startColumnIndex': 7,
                'endColumnIndex': 8
            },
            'rows': [{
                'values': [{'userEnteredValue': {'stringValue': 'ML Record'}}]
            }],
            'fields': 'userEnteredValue'
        }
    })

    ml_record_formula = f"=C{row_index +
                             1}&\"-\"& (B{row_index + 1} - C{row_index + 1})"
    batch_update_body['requests'].append({
        'updateCells': {
            'range': {
                'sheetId': first_sheet_id,
                'startRowIndex': 1,
                'endRowIndex': 2,
                'startColumnIndex': 8,
                'endColumnIndex': 9
            },
            'rows': [{
                'values': [{'userEnteredValue': {'formulaValue': ml_record_formula}}]
            }],
            'fields': 'userEnteredValue'
        }
    })

    batch_update_body['requests'].append({
        'updateCells': {
            'range': {
                'sheetId': first_sheet_id,
                'startRowIndex': 2,
                'endRowIndex': 3,
                'startColumnIndex': 7,
                'endColumnIndex': 8
            },
            'rows': [{
                'values': [{'userEnteredValue': {'stringValue': 'Spread Record'}}]
            }],
            'fields': 'userEnteredValue'
        }
    })
    spread_record_formula = f"=D{row_index +
                                 1}&\"-\"& (B{row_index + 1} - D{row_index + 1})"
    batch_update_body['requests'].append({
        'updateCells': {
            'range': {
                'sheetId': first_sheet_id,
                'startRowIndex': 2,
                'endRowIndex': 3,
                'startColumnIndex': 8,
                'endColumnIndex': 9
            },
            'rows': [{
                'values': [{'userEnteredValue': {'formulaValue': spread_record_formula}}]
            }],
            'fields': 'userEnteredValue'
        }
    })
    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body=batch_update_body
        ).execute()
        format_main_sheet(total_row_index)
    except Exception as e:
        print(f"An error occurred: {e}")
