from __future__ import print_function
import httplib2
import os
import json

 
# from apiclient import discovery
# Commented above import statement and replaced it below because of 
# reader Vishnukumar's comment
# Src: https://stackoverflow.com/a/30811628
 
import googleapiclient.discovery as discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
 
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta, MO
from dateutil import parser


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None
 
# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret_1073221642494-buttpklqbfb4i9drmr3bqqth91k9daen.apps.googleusercontent.com.json'
APPLICATION_NAME = 'test'
 
 
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
                                   'calendar-python-quickstart.json')
 
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
    
def time_difference_in_hours(start_time, end_time):
    """
    Calculate the time difference in hours between two ISO 8601 formatted datetime strings.

    Parameters:
        start_time (str): The starting time as an ISO 8601 formatted string.
        end_time (str): The ending time as an ISO 8601 formatted string.

    Returns:
        float: The time difference in hours.
    """
    # Convert the ISO-formatted start and end times to datetime objects
    dt_start = datetime.fromisoformat(start_time)
    dt_end = datetime.fromisoformat(end_time)

    # Calculate the time difference
    time_difference = dt_end - dt_start

    # Convert the time difference to hours
    time_difference_in_hours = time_difference.total_seconds() / 3600

    return time_difference_in_hours


def get_events_from_last_monday_to_this_monday(service):
    """
    Fetch Google Calendar events from last Monday to this Monday, excluding current Monday.

    Parameters:
        service: Authenticated Google Calendar service object

    Returns:
        events: List of events in the given time frame
    """

    # Get last Monday
    last_monday = datetime.now() + relativedelta(weekday=MO(-1))
    # Make it aware of the timezone
    last_monday = last_monday.replace(tzinfo=datetime.now().astimezone().tzinfo)
    # Set the time to the start of the day
    last_monday = last_monday.replace(hour=0, minute=0, second=0, microsecond=0)

    # Get this Monday
    this_monday = datetime.now() + relativedelta(weekday=MO)
    # Make it aware of the timezone
    this_monday = this_monday.replace(tzinfo=datetime.now().astimezone().tzinfo)
    # Set the time to the start of the day
    this_monday = this_monday.replace(hour=0, minute=0, second=0, microsecond=0)

    # Convert to ISO format
    last_monday_iso = last_monday.isoformat()
    this_monday_iso = this_monday.isoformat()

    # Query Google Calendar API
    events_result = service.events().list(
        calendarId='primary',
        timeMin=last_monday_iso,
        timeMax=this_monday_iso,
        maxResults=100
    ).execute()

    events = events_result.get('items', [])

    return events


def get_day_of_week(dt: str) -> str:
    """
    Get the name of the day of the week from a datetime string.

    Parameters:
        dt: A datetime string in ISO format

    Returns:
        day_name: The name of the day of the week (e.g., 'Monday')
    """
    # Convert the string to a datetime object
    dt_object = parser.parse(dt)
    
    # Extract the name of the day of the week
    day_name = dt_object.strftime('%A')
    
    return day_name


def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    # Call the Calendar API
    events = get_events_from_last_monday_to_this_monday(service)

    work_data = {}

    for event in events:
        if event["summary"] != "work":
            continue

        day_of_week = get_day_of_week(event["start"]["dateTime"])
        start_time = event["start"]["dateTime"]
        end_time = event["end"]["dateTime"]
        time_zone = event["start"].get("timeZone", "Unknown")  # Replace "Unknown" with a default timezone if needed

        # Check if the day already exists in the work_data dictionary
        if day_of_week in work_data:
            work_data[day_of_week].extend([start_time, end_time, time_zone])
        else:
            work_data[day_of_week] = [start_time, end_time, time_zone]

    with open("work_data.json", "w") as f:
        json.dump(work_data, f)




# Call the main function
if __name__ == '__main__':
    main()
