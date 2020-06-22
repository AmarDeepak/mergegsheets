from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import pandas as pd



SPREADSHEET_ID = 'XXXXXXXXXXXXXXXXXXXXXXXXXX' #UPDATE YOUR SPREADSHEET ID
RANGE_NAME1 = 'Sheet1' #Update sheet name
RANGE_NAME2 = 'Sheet2' #Update sheet name
RANGE_NAME3 = 'Sheet3' #Update sheet name

#add more Range_NameX varable if have more sheets


def get_google_sheet(spreadsheet_id, range_name):
    scopes = 'https://www.googleapis.com/auth/spreadsheets.readonly'
    # Setup the Sheets API
    store = file.Storage('credentials.json') # Update credentials file name, if you have different name.
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_secret.json', scopes)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))

    
    gsheet = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    return gsheet


def gsheet2df(gsheet):
    header = gsheet.get('values', [])[0]
    values = gsheet.get('values', [])[1:]
    if not values:
        print('No data found.')
    else:
        all_data = []
        for col_id, col_name in enumerate(header):
            column_data = []
            for row in values:
                column_data.append(row[col_id])
            ds = pd.Series(data=column_data, name=col_name)
            all_data.append(ds)
        df = pd.concat(all_data, axis=1)
        return df


gsheet1 = get_google_sheet(SPREADSHEET_ID, RANGE_NAME1)
gsheet2 = get_google_sheet(SPREADSHEET_ID, RANGE_NAME2)
gsheet3 = get_google_sheet(SPREADSHEET_ID, RANGE_NAME3)

#add gsheet variable if have more sheets

df1 = gsheet2df(gsheet1)
df2 = gsheet2df(gsheet2)
df3 = gsheet2df(gsheet3)

#add more dataframe variables(df) as per gsheet

df = pd.concat([df1,df2,df3], sort=False)

fData = pd.ExcelWriter(r'C:\Users\amar.deepak1\Desktop\many_to_one\Master_sheet.xlsx') #update with your desired path
df.to_excel(fData)
fData.save()
