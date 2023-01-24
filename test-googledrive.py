from __future__ import print_function

import os.path
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from init_bot import initialize

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive']


def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    if os.path.isfile('subslist.xlsx'):
        print('subslist.xlsx file found. Continuing...')
    else:
        print("No subslist.xlsx file found. Please create one and try again.")

    try:
        service = build('drive', 'v3', credentials=creds)
        subs_excel_file = pd.read_excel("subslist.xlsx")

        id_numbers = subs_excel_file['ID NUMBER'].values
        
        # looping over ID numbers
        for id_number in id_numbers:
            result = searchFile(service=service, filename=id_number)
            print(result)
            changePermissions(folder_obj=result, service=service)
        
        print('')
        print('==================================================')
        print('Done!!!')
        print('Say thank you, Master to James Jilhaney UwU')

    except HttpError as error:
        # TODO(developer) - Handle errors from drive API.
        print(f'An error occurred: {error}')


def searchFile(service, filename):
    print(f'id is: {filename}')
    try:
        page_token = None
        results = service.files().list(q="mimeType = 'application/vnd.google-apps.folder' and name='{filename}'".format(filename=filename),
                                            spaces='drive',
                                            fields='nextPageToken, '
                                                   'files(id, name)',
                                            pageToken=page_token).execute()


        if results == None:
            print(f"Results not found. Skipping searching for folder named {filename}")
        
        for file in results.get('files', []):
                # Process change
                # print(F'Found file: {file.get("name")}, {file.get("id")}')
                name = file.get("name")
                id = file.get("id")
                return {"folder_name": name, "id": id }

    except HttpError as error:
        print(error)

def changePermissions(folder_obj, service):

    try:
        email = folder_obj.get('folder_name')
        id = folder_obj.get('id')
        
        request_body = {
            'role' : 'reader',
            'type' : 'user',
            'emailAddress': email + '@my.xu.edu.ph'
        }
        
        permission_response =  service.permissions().create(
        fileId = folder_obj.get('id'),
        body = request_body
            ).execute()

        # print(permission_response)
    
        getSharableLink(service=service,id=id)
    
    except HttpError as error:
        print(error)

def getSharableLink(service, id):
    share_link = service.files().get(
        fileId=id,
        fields='webViewLink'
    ).execute()

    print(share_link)
    

if __name__ == '__main__':
    main()