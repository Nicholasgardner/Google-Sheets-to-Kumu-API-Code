#Google Sheets Importer API 
#Scripps AOO 
#7/14/17
from __future__ import print_function
import httplib2
import os
import json
import datetime

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
	import argparse
	flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
	flags = None 
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
#SCOPES ERROR OCCURS HERE ^
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python - Sheets to Kumu'


def get_credentials():
	#Allows entry to accounts and sheets: Do not edit unless you know what you are doing :)
	home_dir = os.path.expanduser('~')
	credential_dir = os.path.join(home_dir, '.credentials')
	if not os.path.exists(credential_dir):
		os.makedirs(credential_dir)
	credential_path = os.path.join(credential_dir,'sheets.googleapis.com-python-quickstart.json')

	store = Storage(credential_path)
	credentials = store.get()
	if not credentials or credentials.invalid:
		flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
		flow.user_agent = APPLICATION_NAME
		if flags:
			credentials = tools.run_flow(flow, store, flags)
		else:
			credentials = tools.run(flow, store)
		print('Storing credentials to ' + credential_path)
	return credentials

def main():
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?''version=v4')
	service = discovery.build('sheets', 'v4', http=http,discoveryServiceUrl=discoveryUrl)

	#Read-Sheet (Options for 1. Input / 2. Sample) 
	#TODO: Make UI for inputting Sheet ID
	spreadsheet_id = raw_input("Read (Google Sheet ID)	") 
	#spreadsheet_id = '1AsT4pHBRqlfEHTp6H1-3tTKAEGoJatdaNiEtYu8xO3E'

	#Write-Sheet (Linked to Kumu)
	spreadsheet_id2= '1bewupAJBv40PT-aJLiJHCh5X7jTUhG8mKi0vB0a4Omc'
	
	#Call functions with Read sheet as Arg
	tags(spreadsheet_id, service, discoveryUrl, credentials)
	elements(spreadsheet_id, service, discoveryUrl, credentials)
	connections(spreadsheet_id, service, discoveryUrl, credentials)

def tags(spreadsheet1, service, discoveryUrl, credentials):
	#Automates tagging on Read File
	
	#Elements
	service = service
	discoveryUrl = discoveryUrl
	credentials = credentials
	spreadsheet_id = spreadsheet1
	now = datetime.datetime.now()
	thisYear = now.year
	range_name  = 'Groups!A2:I'
	result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
	values = result.get('values', [])

	#Breaking Dates down into Tags with '|' separators
	j = 0
	for row in values:
			i=0
			list1 = []
			if row[5] == "0":
				row[5] = thisYear
			i = ((int(row[5]) - (int(row[4]))))
			while i >= 0:
				col_i = (int(row[5])) - i
				if i>0:
					list1.append("%s|" % (col_i))
				else:
					list1.append("%s" % (col_i))
				i-=1
			values[j].append(((("".join(str(x) for x in list1)))))
			j+=1
	body={'values': values}
	result2 = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range='Groups!A2:I', valueInputOption='RAW', body=body).execute()

	#Connections
	range_name  = 'Connections!A2:G'
	result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
	values = result.get('values', [])
	
	#Breaking Dates down into Tags with '|' separators
	j = 0
	for row in values:
			i=0
			list1 = []
			if row[4] == "0":
				row[4] == thisYear
			i = ((int(row[4]) - (int(row[3]))))
			while i >= 0:
				col_i = (int(row[4])) - i
				if i>0:
					list1.append("%s|" % (col_i))
				else:
					list1.append("%s" % (col_i))
				i-=1
			values[j].append(((("".join(str(x) for x in list1)))))
			j+=1
	body={'values': values}
	result2 = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range='Connections!A2:G', valueInputOption='RAW', body=body).execute()

def elements(spreadsheet1, service, discoveryUrl, credentials):
	#Reads groups on Read Sheet and writes elements onto Write Sheet
	service = service
	discoveryUrl = discoveryUrl
	credentials = credentials
	#Read
	spreadsheet_id = spreadsheet1
	range_name  = 'Groups!A2:I'
	result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
	values = result.get('values', [])

	#Write
	spreadsheet_id2 = '1bewupAJBv40PT-aJLiJHCh5X7jTUhG8mKi0vB0a4Omc'
	body = {'values': values}
	result2 = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id2, range='Elements!A2:I',valueInputOption='RAW', body=body).execute()

def connections(spreadsheet1, service, discoveryUrl, credentials):
	#Reads connections on Read Sheet and writes connections onto Write Sheet
	service = service
	discoveryUrl = discoveryUrl
	credentials = credentials
	#Read
	spreadsheet_id = spreadsheet1
	range_name  = 'Connections!A2:G'
	result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
	values = result.get('values', [])

	#Write
	body = {'values': values}
	spreadsheet_id2 = '1bewupAJBv40PT-aJLiJHCh5X7jTUhG8mKi0vB0a4Omc'
	result2 = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id2, range='Connections!A2:G',valueInputOption='RAW', body=body).execute()

if __name__ == '__main__':
	main()
