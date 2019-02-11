from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

AERVLAN_122_110_ID = '1a19b0Y07-Wtb5b46QyxXO1iw7Mv17ZccMkIeDeI9ntE'
AERVLAN_122_110_RANGE = '2017 Inventory!A8:T139'

def main():
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
			flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
			creds = flow.run_local_server()
		# Save the credentials for the next run
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)
	
	service = build('sheets', 'v4', credentials=creds)

	# Call the Sheets API
	sheet = service.spreadsheets()
	result = sheet.values().get(spreadsheetId=AERVLAN_122_110_ID, range=AERVLAN_122_110_RANGE).execute()
	values = result.get('values', [])

	if not values:
        	print('No data found.')
	else:
        	print('Room, Username, WVU login, User Email, Supervisor, Computer name, Serial No., IP Address, MAC, Wall Port, OS, RT No.')
        for row in values:
		# Print columns A and E, which correspond to indices 0 and 4.
		print('%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s/n' % (row[0], row[1], row[2], row[3], row[4], row[5], row[8], row[9], row[10], row[11], row[12], row[18]))

if __name__ == '__main__':
    main()
