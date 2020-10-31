from __future__ import print_function
import pickle
import os.path
import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import email
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    # Call the Gmail API
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            print(label['name'])


with open('token.pickle', 'rb') as token:
    creds = pickle.load(token)

service = build('gmail', 'v1', credentials=creds)
results = service.users().messages().list(userId='me').execute()

edf = pd.DataFrame([], columns=['Subject', 'Content', 'html'])
try:
    while 'nextPageToken' in results:
        pt = results['nextPageToken']
        results1 = service.users().messages().list(userId='me', pageToken=pt).execute()
        for i in results1['messages']:
            s = i['id']
            message = service.users().messages().get(userId='me', id=s, format='raw').execute()
            msg_raw = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
            msg_str = email.message_from_bytes(msg_raw)
            Email_id = [msg_str['From']]
            sub = [msg_str['Subject']]
            cont = [msg_str['Content-Type']]
            df_temp = pd.DataFrame([(Email_id, sub, cont)], columns=['Email_id','Subject', 'Content'])
            edf = edf.append(df_temp)
except Exception as e:
    print(str(e))

edf.to_csv('EmailsDetails' + str('.eml'))
print(edf)