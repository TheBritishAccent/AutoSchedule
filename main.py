from __future__ import print_function
import httplib2
import os
import csv

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import pytz, datetime

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

class CSV:

	def __init__(self, file_name, start_pos=0):
		self.file_name = file_name
		self.start_pos = start_pos
		self.rows = {}

		self.load_file()

	def load_file(self):
		# Open csv, extract all rows to rows object
		csvLength = self.start_pos

		with open(self.file_name) as csvfile:
			reader = csv.DictReader(csvfile)
			
			for row in reader:
				self.rows[csvLength] = dict(row)
				csvLength += 1
			csvfile.close()
		return self.rows, csvLength
	
	def find(self, search):
		pass

	def get_rows(self):
		return self.rows


class AutoSchedule:

	def __init__(self, calendar_name='AutoSchedule', schedule_csv_name='schedule.csv', day_zero_csv_name='DayZero.csv'):
		self.credentials = self.get_credentials()
		self.http = self.credentials.authorize(httplib2.Http())
		self.service = discovery.build('calendar', 'v3', http=self.http)

		self.calendar_name = calendar_name

		self.schedule_csv = CSV(file_name=schedule_csv_name, start_pos=1)
		self.zero_csv = CSV(file_name=day_zero_csv_name, start_pos=1)

		self.locale = pytz.timezone("America/Nassau")

		

	# Getter/Setter

	def set_calendar_name(self, calendar_name):
		self.calendar_name = calendar_name

	def get_credentials(self):
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

	# Load/Create Calendars

	def load_all_calendars(self):
		# Get all calendars, find specific one needed for this program
		calendar_list = self.service.calendarList().list().execute()
		for calendar_list_entry in calendar_list['items']:
			if calendar_list_entry['summary'] == self.calendar_name:
				return calendar_list_entry['id']

		return None

	def create_calendar(self):
		_ecalendar = {
			'summary': self.calendar_name,
		}
		_created_calendar = self.service.calendars().insert(body=_ecalendar).execute()
		
		_calendar_id = _created_calendar['id']
		print("Created Calendar! Name: {} ID: {}".format(self.calendar_name, _calendar_id))
		return _created_calendar

	def load_calendar(self, calendar_id):
		_calendar = self.service.calendars().get(calendarId=calendar_id).execute()
		
		if not _calendar:
			print("Could not retrieve calendar")
		else:
			print("Calendar Exists! Name: {} ID: {}".format(self.calendar_name, calendar_id))

		return _calendar

	# Manage events

	def check_for_events(self, calendar_id):
		# Check if program has already been run (events already created)
		_event_list = self.service.events().list(calendarId=calendar_id).execute()
		for l_event in _event_list['items']:
			if 'Day' in l_event['summary']:
				return True
		return False

	def create_events(self, calendar_id):
		d = datetime.datetime.utcnow().date()
		
		day = int(input("Today's Day: "))
		total_days = int(input("How many days: "))
		
		# Create csv objects
		rows, schedule_length = self.schedule_csv.load_file()
		zero_rows, zero_length = self.zero_csv.load_file()

		# Iterate for amount of days user specified
		for i in range(0, total_days):
			# Check if date is within weekday (Mon-Fri)
			if d.isoweekday() in range(1,6):
				formatted_date = str(d)
				
				# Check if today is a day zero
				for z in range(1, zero_length):
					if zero_rows[z]["Zero"] == formatted_date:
						c_event = self.service.events().insert(calendarId=calendar_id, body={'summary': 'Day 0', 
							'start': {'date': formatted_date},
							'end': {'date': formatted_date}}).execute()
						if not c_event:
							print("Today ({}) should have been a day 0, but wasn't.".format(formatted_date))
						else:
							print("Event created. Day: 0")
						break
				else:
					# Iterate through all days
					for x in range(1, schedule_length):
						a = x
						
						# Match current day with class in csv
						if rows[a]["Day"] == str(day):
							start_raw = datetime.datetime.strptime("{} {}".format(formatted_date, 
							rows[a]["Start Time"]), "%Y-%m-%d %H:%M:%S")
							sDateTime = self.locale.localize(start_raw, is_dst=None)
							sDateTime = sDateTime.astimezone(pytz.utc)
							
							end_raw = datetime.datetime.strptime("{} {}".format(formatted_date, 
							rows[a]["End Time"]), "%Y-%m-%d %H:%M:%S")
							eDateTime = self.locale.localize(end_raw, is_dst=None)
							eDateTime = sDateTime.astimezone(pytz.utc)

							sDateTime = "{}T{}z".format(sDateTime.strftime ("%Y-%m-%d"), sDateTime.strftime ("%H:%M:%S"))
							eDateTime = "{}T{}z".format(eDateTime.strftime ("%Y-%m-%d"), eDateTime.strftime ("%H:%M:%S"))
							
							event = {
										'summary': rows[a]["Class"],
										'start': {
											'dateTime': sDateTime
										},
										'end': {
											'dateTime': eDateTime
										}
									}
							c_event = self.service.events().insert(calendarId=calendar_id, body=event).execute()
							
							if not c_event:
								print("Error creating event")
							else:
								print('Event created. Class: {} Day: {}'.format(rows[a]["Class"], day))
					
					day += 1
				
				# Roll-Over
				if day > 6:
					day = 1
			
			d += datetime.timedelta(days=1)

def main():
	print("1: Create new calendar")
	print("2: Use existing calendar")
	print("3: Use default calendar")
	print("4: Reset OAuth2")
	user_input = input("Enter: ")

	AS = AutoSchedule()

	if user_input == "1":
		c_name = input("Enter calendar name: ")

		AS.set_calendar_name(c_name)
		cal = AS.create_calendar()
		AS.create_events(cal['id'])
		
	elif user_input == "2":
		calendar_id = input("Enter calendar ID: ")
		cal = AS.load_calendar(calendar_id)
		AS.create_events(cal['id'])
	
	elif user_input == "3":
		cal_id = AS.load_all_calendars()
		if cal_id is not None:
			cal = AS.load_calendar(cal_id)
		else:
			cal = AS.create_calendar()

		if not AS.check_for_events(cal['id']):
			AS.create_events(cal['id'])
		else:
			print("Events already exist!")

	elif user_input == "4":
		credential_path = os.path.join(os.path.join(os.path.expanduser('~'), '.credentials'), 'AutoSchedule.json')
		
		if not os.path.exists(credential_path):
			os.remove(credential_path)
			print("OAuth Removed")
		else:
			print("Did not find file")

		os.system("pause")


if __name__ == '__main__':
	main()
