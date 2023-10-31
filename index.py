import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

load_dotenv()

GSS_TEMP_KEY = os.environ['GSS_TEMP_KEY']


def get_gss_workbook():
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credential = ServiceAccountCredentials.from_json_keyfile_name(
        './gss_credential.json', scope)
    wb = gspread.authorize(credential).open_by_key(GSS_TEMP_KEY)
    return wb


def get_gss_worksheet(sheet_name):
    ws = get_gss_workbook().worksheet(sheet_name)
    return ws


def main():
    ws = get_gss_worksheet('名簿')
    value = ws.acell('B10').value
    print(value)


main()
