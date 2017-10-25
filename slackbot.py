from __future__ import print_function
from slackclient import SlackClient
import time
import datetime
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
CLIENT_SECRET_FILE = 'client_id.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_sheets_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
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
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_drive_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'drive-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, 'https://www.googleapis.com/auth/drive')
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def insertEntry(ts, user, order_for, spreadsheetId, service):
    rangeName = 'Sheet1!A2:C'
    result = service.spreadsheets().values().get(
        spreadsheetId='188OrUd-0quwvus9anW_wE4uCIlkY5QrtvbRo0oh3Fxc', range=rangeName).execute()
    values = result.get('values', [])
    lastRow = len(values) + 2
    rangeName = 'Sheet1!A' + str(lastRow) + ':D' + str(lastRow)
    values = [
        [ts, user, order_for, spreadsheetId]
    ]

    body = {
        'values': values
    }

    service.spreadsheets().values().update(
        spreadsheetId='188OrUd-0quwvus9anW_wE4uCIlkY5QrtvbRo0oh3Fxc', range=rangeName,
        valueInputOption='RAW', body=body).execute()


def process_entry(message, sheets_service, drive_service, sc):
    placing = False

    if message[0]["type"] == "desktop_notification" and 'bot' not in message[0]['subtitle']:
        sc.api_call(
            "chat.postMessage",
            channel=message[0]["channel"],
            text="Ready to place an order? (yes/no)"
        )

        responded = False
        while not responded:
            item = sc.rtm_read()
            print(item)
            if len(item) > 0:
                if item[0]["type"] == "desktop_notification" and 'bot' not in item[0]['subtitle']:
                    if "y" in item[0]["content"].lower():
                        print("changing placing", item)
                        print("bot in message?", 'bot' not in item[0]['subtitle'])
                        placing = True
                    responded = True
            time.sleep(1)

        if placing:
            ts = datetime.datetime.fromtimestamp(
                                int(float(item[0]["event_ts"]))
                                ).strftime('%Y-%m-%d %H:%M:%S')
            sc.api_call(
                "chat.postMessage",
                channel=message[0]["channel"],
                text="Great! At a high level, what is this for? (aero, aux, non-car, etc.)"
            )

            responded = False
            order_for = None
            while not responded:
                item = sc.rtm_read()
                print(item)
                if len(item) > 0:
                    if item[0]["type"] == "desktop_notification" and 'bot' not in item[0]['subtitle']:
                        order_for = item[0]["content"]
                        name = ts.split(" ")[0] + "_" + order_for
                        spreadsheet = setup_spreadsheet(name, sheets_service, drive_service)
                        responded = True
                time.sleep(1)

            sc.api_call(
                "chat.postMessage",
                channel=message[0]["channel"],
                text="Please fill out your order on this spreadsheet: " + spreadsheet + "\nSend \"Done\" here when you're finished with the spreadsheet."
            )

            done = False
            ts = user = None
            while not done:
                item = sc.rtm_read()
                print(item)
                if len(item) > 0:
                    if item[0]["type"] == "desktop_notification" and 'bot' not in item[0]['subtitle']:
                        if "done" in item[0]["content"].lower():
                            print("changing placing", item)
                            print("bot in message?", 'bot' not in item[0]['subtitle'])
                            done = True
                            ts = datetime.datetime.fromtimestamp(
                                int(float(item[0]["event_ts"]))
                                ).strftime('%Y-%m-%d %H:%M:%S')
                            user = item[0]["subtitle"]
                            insertEntry(ts, user, order_for, spreadsheet, sheets_service)
                time.sleep(1)

            connor = "U0AHWLYSC"
            sina = "U0AHVK45R"

            sc.api_call(
                "chat.postMessage",
                channel=connor,
                text="There's a new order in the queue."
            )
            sc.api_call(
                "chat.postMessage",
                channel=sina,
                text="There's a new order in the queue."
            )



def setup_spreadsheet(name, sheets_service, drive_service):
    request = sheets_service.spreadsheets().create(body={'properties': {'title': name}})
    response = request.execute()
    file_id = response["spreadsheetId"]
    folder_id = '0B2WDbiN0AUY7SG1fbUxBc2ZpNk0'

    file = drive_service.files().get(fileId=file_id,
                                     fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    file = drive_service.files().update(fileId=file_id,
                                        addParents=folder_id,
                                        removeParents=previous_parents,
                                        fields='id, parents').execute()

    range_ = "Sheet1!A1:C1"

    body = {
        'values': [["Supplier", "Part Number", "Part Name/Description"]]
    }

    request = sheets_service.spreadsheets().values().update(spreadsheetId=file_id, range=range_, body=body,
                                                            valueInputOption='RAW')
    request.execute()

    return response["spreadsheetUrl"]


def main():
    sheets_credentials = get_sheets_credentials()
    drive_credentials = get_drive_credentials()
    http = sheets_credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    sheets_service = discovery.build('sheets', 'v4', http=http,
                                     discoveryServiceUrl=discoveryUrl)
    http = drive_credentials.authorize(httplib2.Http())
    drive_service = discovery.build('drive', 'v3', http=http)

    slack_token = ""
    sc = SlackClient(slack_token)

    # stuff = sc.api_call(
    #     "conversations.list"
    # )
    # print(stuff)
    # for member in stuff["channels"]:
    #     print(member)
    #     # print(member["name"], member["id"])

    if sc.rtm_connect():
        while True:
            item = sc.rtm_read()
            print(item)
            if len(item) > 0:
                process_entry(item, sheets_service, drive_service, sc)
            time.sleep(1)
    else:
        print("Connection Failed")


if __name__ == '__main__':
    main()
