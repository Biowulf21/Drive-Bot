from __future__ import print_function

import os.path
import pandas as pd

from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError

from init_bot import initialize



def main():
   # initializes the bot

    service = initialize() 

    try:
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