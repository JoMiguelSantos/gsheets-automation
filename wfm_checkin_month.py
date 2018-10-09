from __future__ import print_function
import httplib2
import os
from pprint import pprint
from apiclient import discovery
from oauth2client import client, tools
from oauth2client.file import Storage
import pytz
import datetime
from time import sleep



def get_credentials():
    """Gets valid user credentials from storage.
    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.
    Returns:
        Credentials, the obtained credential.
    """
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    CLIENT_SECRET_FILE = 'client_secret.json'

    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
    return credentials


def buildGS():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    return service


spreadsheetId = 'spreadsheetID'
getLateness = buildGS().spreadsheets().get(spreadsheetId=spreadsheetId, ranges=['Lateness']).execute()
rowCount = getLateness['sheets'][0]['properties']['gridProperties']['rowCount']
colCount = getLateness['sheets'][0]['properties']['gridProperties']['columnCount']
rangeNameAll = f'Lateness!A1:I{rowCount}'
rangeName = f'Lateness!A3:I{rowCount}'
copyLateness = buildGS().spreadsheets().values().batchGet(spreadsheetId=spreadsheetId, ranges=rangeName).execute()
copyLatenessAll = buildGS().spreadsheets().values().batchGet(spreadsheetId=spreadsheetId, ranges=rangeNameAll).execute()
local = pytz.timezone("Europe/Berlin")
current_month = datetime.datetime.strftime(datetime.datetime.today(),'%Y-%m')
current_day = str(datetime.date.today())

def createOrUpdateCheckInSheet():
    for date in [current_month]:
        try:
            requests = []
            requests.append({'addSheet':
                                 {'properties':
                                      {'title': f'{date}'
                                       }
                                  }
                             }
                            )

            body = {'requests': requests}
            buildGS().spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=body).execute()
            value_range_body = {'values': copyLatenessAll['valueRanges'][0]['values']}
            buildGS().spreadsheets().values().update(spreadsheetId=spreadsheetId, range=f'{date}!A1:I{rowCount}', body=value_range_body,valueInputOption='RAW').execute()
        except:
            value_range_body = {'values': copyLateness['valueRanges'][0]['values']}
            value_input_option = 'RAW'
            buildGS().spreadsheets().values().append(spreadsheetId=spreadsheetId, range=f'{date}!A1:I{rowCount}',
                                                   valueInputOption=value_input_option,
                                                   insertDataOption='INSERT_ROWS', body=value_range_body).execute()

if __name__ == "__main__":
    try:
        createOrUpdateCheckInSheet()
    except:
        sleep(60)
        try:
            createOrUpdateCheckInSheet()
        except:
            sleep(600)
            try:
                createOrUpdateCheckInSheet()
            except:
                sleep(6000)
                createOrUpdateCheckInSheet()





