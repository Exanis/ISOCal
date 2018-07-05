"""
ISOCal
Tool to sync events from office365 to google calendar
"""
import datetime
import time
from httplib2 import Http
from apiclient.discovery import build
from oauth2client import file, client, tools
from O365 import Schedule


O365_LOGIN = ''
O365_PASSWORD = ''
TIMEZONE = 'GMT'


def gcal_auth():
    """
    Log to google api
    """
    scopes = 'https://www.googleapis.com/auth/calendar'
    store = file.Storage('credentials.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_secret.json', scopes)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))
    return service


def gcal_get_events(service):
    """
    Retrieve events from Google Calendar to compare with o365
    """
    now = datetime.datetime.utcnow().isoformat() + 'Z'
    request = service.events().list(calendarId='primary', timeMin=now).execute()
    return request.get('items', [])


def o365_auth(username, password):
    """
    Log to O365 API
    """
    service = Schedule((username, password))
    service.getCalendars()
    return service


def o365_get_events(service):
    """
    Retrieve all event in all calendars
    """
    events = []
    for cal in service.calendars:
        cal.getEvents()
        events = events + [e for e in cal.events]
    return events


def get_google_date(event_date):
    """
    Retrieve the date from a google event
    (google may store the data in two different place depending of the format)
    """
    if 'dateTime' in event_date:
        return time.strptime(event_date['dateTime'].replace(':', ''), "%Y-%m-%dT%H%M%S%z")
    return time.strptime(event_date['date'], '%Y-%m-%d')


def is_same_event(g_event, o_event):
    """
    Return true if two events are the same
    """
    return g_event['summary'] == o_event.getSubject() \
        and o_event.getStart() == get_google_date(g_event['start']) \
        and o_event.getEnd() == get_google_date(g_event['end']) \


def event_exist(g_events, event):
    """
    Return True if the given event (o365) already exist in google list
    """
    search_list = [
        e for e in g_events if is_same_event(e, event)
    ]
    return bool(search_list)


def filter_duplicate_events(g_events, o_events):
    """
    Remove already existing event from the list
    """
    filtered = [e for e in o_events if not event_exist(g_events, e)]
    return filtered


def add_events(g_service, events):
    """
    Add all events in list to google calendar
    """
    for event in events:
        start = event.getStart()
        end = event.getEnd()
        request = g_service.events().insert(
            calendarId="primary",
            body={
                "summary": event.getSubject(),
                "description": event.getBody(),
                "location": event.getLocation(),
                "start": {
                    "dateTime": time.strftime("%Y-%m-%dT%H:%M:%S", start),
                    "timeZone": TIMEZONE
                },
                "end": {
                    "dateTime": time.strftime("%Y-%m-%dT%H:%M:%S", end),
                    "timeZone": TIMEZONE
                }
            }
        )
        request.execute()


if __name__ == "__main__":
    G_SERVICE = gcal_auth()
    G_EVENTS = gcal_get_events(G_SERVICE)
    O_SERVICE = o365_auth(O365_LOGIN, O365_PASSWORD)
    O_EVENTS = o365_get_events(O_SERVICE)
    NEW_EVENTS = filter_duplicate_events(G_EVENTS, O_EVENTS)
    add_events(G_SERVICE, NEW_EVENTS)
