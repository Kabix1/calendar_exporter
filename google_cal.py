#!/usr/bin/env python

import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def get_creds():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('creds/token.pickle'):
        with open('creds/token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'creds/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('creds/token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


def add_event(start, end, summary, location, cal_id):

    creds = get_creds()
    service = build('calendar', 'v3', credentials=creds)
    event = {
        'summary': summary,
        'location': location,
        'start': {
            'dateTime': start.strftime('%Y-%m-%dT%H:%M:%S'),
            'timeZone': "Europe/Stockholm"
        },
        'end': {
            'dateTime': end.strftime('%Y-%m-%dT%H:%M:%S'),
            'timeZone': "Europe/Stockholm"
        }
    }

    # Call the Calendar API
    event = service.events().insert(calendarId=cal_id, body=event).execute()
    print(f"Event created: {event.get('htmlLink')}")


def main():
    creds = get_creds()
    service = build('calendar', 'v3', credentials=creds)
    calendars = service.calendarList().list().execute()['items']
    for cal in calendars:
        print(cal)


if __name__ == "__main__":
    main()
