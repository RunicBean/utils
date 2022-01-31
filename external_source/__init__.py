import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

cred_path = './credentials'


def get_google_sheet_to_arrays_with_oath2token(spreadsheet_id, sheet_name, column_range, header_line=0):

    # If modifying these scopes, delete the file token.json.
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    # Get credentials file and save required token to token.json
    creds = None
    gcp_oauth_token_json = os.path.join(cred_path, 'gcp_sa_token.json')
    if os.path.exists(gcp_oauth_token_json):
        creds = Credentials.from_authorized_user_file(gcp_oauth_token_json, scopes)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(os.path.join(cred_path, 'gcp_oath2_token.json'),
                                                             scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(os.path.join(cred_path, 'gcp_sa_token.json'), 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    sample_range_name = f'{sheet_name}!{column_range}'
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                range=sample_range_name).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')

    arrays_with_headers = []
    headers = [header for header in values[header_line]]
    # for row in values[1:]:
    #     arrays_with_headers.append({headers[i]: row[i] for i in range(len(headers))})

    return headers, values


def get_google_sheet_to_arrays(spreadsheet_name, sheet_name):
    credential = ServiceAccountCredentials.from_json_keyfile_name(os.path.join(cred_path, 'ms-development-key.json'),
                                                                  ["https://spreadsheets.google.com/feeds",
                                                                   "https://www.googleapis.com/auth/spreadsheets",
                                                                   "https://www.googleapis.com/auth/drive.file",
                                                                   "https://www.googleapis.com/auth/drive"])
    client = gspread.authorize(credential)
    gsheet = client.open(spreadsheet_name).worksheet(sheet_name)
    records = gsheet.get_all_records()
    return records
