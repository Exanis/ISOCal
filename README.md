# ISOCal
Just a simple tool to sync office 365 calendar with Google Calendar in order to be able to use google home

## Warning
This script was made for my own usage. It may or may not fit your need.

## How to use
- Pre-requisite : have python3 and pip installed
- Fill the `__main__.py` first lines with your own email / password for O365 (lines 6 and 7)
- Also set your timezone if it's not the same as mine
- Go to [this url](https://developers.google.com/calendar/quickstart/python) and follow the step 1 to get a `credentials.json` file
- Put it in the same directory that `__main__.py` and name it `client_secret.json`
- Install the requirements : `pip install -r requirements.txt`
- Run the script : `python __main__.py`
    - If you are running the script for the first time, a browser window will open to auth your google account with the app. Just hit accept :)

If you need to keep your calendars in sync, just add a cron task to run this script every so often.