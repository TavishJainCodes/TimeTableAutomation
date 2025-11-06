from __future__ import print_function
import os
import pickle
import sys
import time
from datetime import datetime, timedelta
import json
import os
from google.oauth2 import service_account
import httplib2
from pytz import timezone
from bs4 import BeautifulSoup as Bs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# --- Constants ---
SCOPES = ['https://www.googleapis.com/auth/calendar.events']
CHROMEDRIVER_PATH = r"C:\Users\tavis\Documents\Tavish Jain\Programming\Python\timetableAutomation\chromedriver-win64\chromedriver.exe"
CLIENT_SECRETS = r'C:\Users\tavis\Documents\Tavish Jain\Programming\Python\timetableAutomation\client_secret_391241584161-qlt50sfo9qn7577dngj8oqu8k0kirsjk.apps.googleusercontent.com.json'
CALENDAR_ID = 'feef0af207b77b65d5cd47adcae6dc74c6af1e34db73a5d326a1e2ce83fd4514@group.calendar.google.com'
BASE_TZ = 'Asia/Kolkata'

def authenticate_google(service_account_file=None):
    """
    Authenticate using a service account.
    If SERVICE_KEY_JSON (env var) is present, use that JSON directly (for CI).
    Otherwise fall back to reading a file path if provided.
    """
    sa_json_env = os.environ.get("SA_KEY")
    if sa_json_env:
        info = json.loads(sa_json_env)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        return creds

    if service_account_file:
        creds = service_account.Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
        return creds

    raise RuntimeError("No service account credentials found. Set SA_KEY env or provide a path.")

# --- Google API helper ---
def build_service_with_timeout(creds, timeout_seconds=15):
    # googleapiclient.build does not allow passing both 'http' and 'credentials'.
    # Pass only credentials to avoid the ValueError. If you need per-call timeouts
    # we can implement them separately (see note below).
    return build('calendar', 'v3', credentials=creds, cache_discovery=False)

# --- Deletion of previously-created timetable events ---
def delete_timetable_events_in_range(service, calendar_id, base_tz=BASE_TZ):
    tz = timezone(base_tz)
    week_start_local = tz.localize(get_current_week_monday())
    week_end_local = week_start_local + timedelta(days=7)

    timeMin = week_start_local.astimezone(timezone('UTC')).isoformat()
    timeMax = week_end_local.astimezone(timezone('UTC')).isoformat()

    print(f"Querying events between {timeMin} and {timeMax} (UTC)...")

    matched = []
    page_token = None
    while True:
        try:
            resp = service.events().list(
                calendarId=calendar_id,
                timeMin=timeMin,
                timeMax=timeMax,
                maxResults=2500,
                singleEvents=True,
                pageToken=page_token
            ).execute()
        except Exception as e:
            print("ERROR listing events:", repr(e))
            return matched

        items = resp.get('items', [])
        for ev in items:
            if 'timetable-script' in (ev.get('description') or '').lower():
                matched.append(ev)
        page_token = resp.get('nextPageToken')
        if not page_token:
            break

    print(f"Found {len(matched)} timetable-script events to delete.")
    for ev in matched:
        attempts = 0
        while attempts < 4:
            try:
                service.events().delete(calendarId=calendar_id, eventId=ev['id']).execute()
                print("Deleted:", ev.get('summary'), ev.get('id'))
                break
            except Exception as e:
                attempts += 1
                wait = 1.5 ** attempts
                print(f"Delete failed (attempt {attempts}) for {ev.get('id')}: {e}. Retrying in {wait:.1f}s")
                time.sleep(wait)
        else:
            print("Giving up on deleting:", ev.get('id'))
    return matched

# --- Timetable scraping & parsing ---
def classInfo(cell):
    """Extract subject, room, and professor info from a timetable cell."""
    info = cell.split('-')
    if len(info) < 4:
        return ["", "", ""]

    subject_map = {
        'TH.COMPUT.': 'TOC', 'TOC': 'TOC', 'SOFT.ENG': 'Software Engineering',
        'ALGO.&ADVDATA': 'Adv. DSA', 'IWP': 'Web Programming'
    }
    prof_map = {'DSN': 'DJ', 'GEV': 'Geetika', 'AVJ': 'AKJ', 'AVJ<': 'AKJ', 'CGT1': 'placeholder'}

    subj = subject_map.get(info[1].upper(), info[1])
    prof = prof_map.get(info[3].upper(), info[3])
    room = info[2]

    if info[0] == "GE":
        return ["GE- Maths", "SNB06", "Rana Sir"]
    if info[0] == "SEC":
        return ["SEC- PFP", "SNB02", "Anand Kumar Singh"]

    return [subj, room, prof]

def get_color_id(location):
    first = location.split()[0]
    if first in ('CS', 'Comp'):
        return '7'
    return '6'  # Default

def getEvents():
    """Scrape timetable and return structured list of events."""
    service = Service(CHROMEDRIVER_PATH)
    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless=new")   # or "--headless" if older Chrome
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")

    driver = webdriver.Chrome(service=service, options=opts)
    driver.get("https://cvs.collegetimetable.in/")

    Select(driver.find_element(By.ID, "classid")).select_by_value("1")
    Select(driver.find_element(By.ID, "semester")).select_by_value("5")
    driver.find_element(By.NAME, "submit").click()

    time.sleep(5)  # wait for page to load
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//table[@border="1" and @align="center"]'))
    )

    soup = Bs(driver.page_source, 'html.parser')
    driver.quit()
    table = soup.find('table', {"align": "center", "border": "1"})
    rows = table.find_all('tr')

    # --- Clean and normalize timings ---
    raw_timings = [t.get_text(strip=True) for t in rows[0].find_all('td')[1:]]
    timings = []
    for t in raw_timings:
        t = t.replace('.', ':').replace('Noon', 'PM').strip()
        t = ''.join(ch for ch in t if ch.isdigit() or ch in [':', 'A', 'P', 'M', ' '])
        t = t.replace("AM", " AM").replace("PM", " PM").strip()
        t = ' '.join(t.split())
        if not any(x in t for x in ["AM", "PM"]):
            t += " AM"
        timings.append(t)

    events = []
    for row in rows[1:7]:  # Monday–Saturday
        cols = [c.get_text(strip=True) for c in row.find_all('td')]
        if not cols:
            continue
        day = cols[0]
        for i, cell in enumerate(cols[1:]):
            if not cell.strip():
                continue
            subject, room, prof = classInfo(cell)
            start_str = timings[i]
            try:
                start_dt = datetime.strptime(start_str, "%I:%M %p")
            except ValueError:
                if "AM" in start_str:
                    start_str = start_str.replace("AM", " AM")
                elif "PM" in start_str:
                    start_str = start_str.replace("PM", " PM")
                start_dt = datetime.strptime(start_str.strip(), "%I:%M %p")

            end_dt = start_dt + timedelta(hours=1)
            events.append({
                "summary": subject,
                "day": day,
                "start_time": start_dt.strftime("%I:%M %p"),
                "end_time": end_dt.strftime("%I:%M %p"),
                "location": f"{room} by {prof}"
            })

    return events

# --- Date/time helpers ---
def get_current_week_monday():
    today = datetime.today()
    return today - timedelta(days=today.weekday())

def convert_to_gcal_event(ev, base_date=None, tz=BASE_TZ):
    """Convert parsed event to Google Calendar format."""
    if base_date is None:
        base_date = get_current_week_monday()
    weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    day_offset = weekdays.index(ev['day'])
    date = base_date + timedelta(days=day_offset)
    fmt = "%Y-%m-%d %I:%M %p"

    start = datetime.strptime(f"{date.strftime('%Y-%m-%d')} {ev['start_time']}", fmt)
    end = datetime.strptime(f"{date.strftime('%Y-%m-%d')} {ev['end_time']}", fmt)

    return {
        'summary': ev['summary'],
        'location': ev['location'],
        'start': {'dateTime': start.isoformat(), 'timeZone': tz},
        'end': {'dateTime': end.isoformat(), 'timeZone': tz}
    }

# --- Main flow ---
def main():
    creds = authenticate_google()
    service = build_service_with_timeout(creds, timeout_seconds=15)

    # Delete only the events this script created (by matching description)
    deleted = delete_timetable_events_in_range(service, CALENDAR_ID, base_tz=BASE_TZ)
    print(f"Deleted {len(deleted)} old timetable events.\n")

    # Add new events
    events = getEvents()
    for ev in events:
        gcal_event = convert_to_gcal_event(ev)
        gcal_event['description'] = 'source: timetable-script'
        gcal_event['colorId'] = get_color_id(ev['location'])
        try:
            service.events().insert(calendarId=CALENDAR_ID, body=gcal_event).execute()
            print(f"Added: {ev['summary']} ({ev['day']} {ev['start_time']})")
        except Exception as e:
            print("Failed to add event:", ev, e)

    print(f"\nTimetable updated successfully with {len(events)} new events.")
    sys.exit(0)

if __name__ == '__main__':
    main()
