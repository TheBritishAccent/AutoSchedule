from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime

try:
	import argparse
	flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
	flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'AutoSchedule'
calendar_name = 'AutoSchedule'


def get_credentials():
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
								   'AutoSchedule.json')

	store = Storage(credential_path)
	credentials = store.get()
	if not credentials or credentials.invalid:
		flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
		flow.user_agent = APPLICATION_NAME
		if flags:
			credentials = tools.run_flow(flow, store, flags)
		else: # Needed only for compatibility with Python 2.6
			credentials = tools.run(flow, store)
		print('Storing credentials to ' + credential_path)
	return credentials

def main():
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('calendar', 'v3', http=http)

	# Calendar List

	calendar_id = ''
	run = True

	calendar_list = service.calendarList().list().execute()
	for calendar_list_entry in calendar_list['items']:
		if calendar_list_entry['summary'] == calendar_name:
			calendar_id = calendar_list_entry['id']
	
	# Create Calendar
	if calendar_id == '':
		ecalendar = {
			'summary': calendar_name,
		}
		created_calendar = service.calendars().insert(body=ecalendar).execute()
		calendar_id = created_calendar['id']

		print("Created Calendar! Name: {} ID: {}".format(calendar_name, calendar_id))
	else: # Get Calendar
		calendar = service.calendars().get(calendarId=calendar_id).execute()
		print("Calendar Exists! Name: {} ID: {}".format(calendar_name, calendar_id))




	event_list = service.events().list(calendarId=calendar_id).execute()
	for l_event in event_list['items']:
		if 'Day' in l_event['summary']:
			run = False

	if run:
		d = datetime.datetime.now().date()
		day = int(input("Today's Day: "))

		for i in range(0,31):
			if d.isoweekday() in range(1,6):
				formatted_day = "Day {}".format(day)
				formatted_date = str(d)

				event = {
					'summary': formatted_day,
					'start': {
						'date': formatted_date
					},
					'end': {
						'date': formatted_date
					},
					'reminders': {
						"useDefault": False,
						'overrides': [
							{
								'method': 'popup',
								'minutes': 420
							}
						]
					}
				}
				
				c_event = service.events().insert(calendarId=calendar_id, body=event).execute()
				print('Event created. Date: {} Day: {}'.format(d, day))

				day += 1
				
				# Roll-Over
				if day > 6:
					day = 1
			
			d += datetime.timedelta(days=1)
	else:
		print("Events already exist!")

if __name__ == '__main__':
	main()