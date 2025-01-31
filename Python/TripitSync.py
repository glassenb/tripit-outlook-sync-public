import requests
import icalendar
import datetime
import win32com.client
import pytz
import json
import os
import sys
import time
import hashlib
from zoneinfo import ZoneInfo

def log(msg):
    """Log message with timestamp"""
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

class TripitSync:
    def __init__(self, ics_url):
        self.ics_url = ics_url
        self.state_file = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), 'tripit_sync_state.json')
        self.initialize_outlook()

    def initialize_outlook(self):
        """Initialize Outlook connection and get/create TripIt calendar"""
        try:
            self.outlook = win32com.client.Dispatch('Outlook.Application')
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            raise Exception(f"Failed to initialize Outlook: {e}")

    def get_calendar(self, retry=3):
        """Get or create TripIt calendar with retries"""
        for attempt in range(retry):
            try:
                root_folder = self.namespace.GetDefaultFolder(9)  # 9 = Calendar folder
                
                # Find existing calendar
                for folder in root_folder.Folders:
                    if folder.Name == "TripIt":
                        log("Found existing TripIt calendar")
                        return folder

                # Create new calendar
                calendar = root_folder.Folders.Add("TripIt")
                log("Created new TripIt calendar")
                return calendar

            except Exception as e:
                if attempt == retry - 1:  # Last attempt
                    raise
                log(f"Calendar access failed (attempt {attempt + 1}/{retry}). Retrying...")
                time.sleep(2)
                self.initialize_outlook()  # Reset Outlook connection

    def clear_calendar(self, calendar, retry=3):
        """Clear calendar items with retries"""
        for attempt in range(retry):
            try:
                items = calendar.Items
                if items.Count > 0:
                    log(f"Clearing {items.Count} items...")
                    
                    # Delete items one by one
                    for i in range(items.Count, 0, -1):
                        try:
                            item = items.Item(i)
                            item.Delete()
                        except Exception as e:
                            log(f"Failed to delete item {i}: {e}")
                            continue
                
                log("Calendar cleared")
                return True

            except Exception as e:
                if attempt == retry - 1:  # Last attempt
                    raise
                log(f"Calendar clearing failed (attempt {attempt + 1}/{retry}). Retrying...")
                time.sleep(2)
                calendar = self.get_calendar()  # Get fresh calendar reference

        return False

    def get_ics_content(self):
        """Fetch ICS content and check if it has changed"""
        try:
            response = requests.get(self.ics_url)
            response.raise_for_status()
            current_content = response.text
            current_hash = hashlib.sha256(current_content.encode()).hexdigest()

            # Check if content has changed
            try:
                if os.path.exists(self.state_file):
                    with open(self.state_file, 'r') as f:
                        state = json.load(f)
                        if state.get('hash') == current_hash:
                            log(f"No changes since last sync ({state.get('last_sync', 'unknown')})")
                            return None
            except Exception as e:
                log(f"Error checking state file: {e}")

            # Save new state
            with open(self.state_file, 'w') as f:
                json.dump({
                    'hash': current_hash,
                    'last_sync': datetime.datetime.now().isoformat()
                }, f)

            return current_content
        except Exception as e:
            raise Exception(f"Failed to fetch ICS feed: {e}")

    def parse_datetime(self, dt_value, tzinfo=None):
        """Parse datetime from ICS, handling both date and datetime values"""
        try:
            if isinstance(dt_value, datetime.datetime):
                # If time is naive (no timezone), assume it's UTC
                if dt_value.tzinfo is None:
                    dt_value = dt_value.replace(tzinfo=datetime.timezone.utc)
                
                # Convert to America/Chicago time
                chicago_tz = ZoneInfo('America/Chicago')
                return dt_value.astimezone(chicago_tz).replace(tzinfo=None)
                
            elif isinstance(dt_value, datetime.date):
                dt = datetime.datetime.combine(dt_value, datetime.time.min)
                chicago_tz = ZoneInfo('America/Chicago')
                return dt.replace(tzinfo=chicago_tz).replace(tzinfo=None)
            else:
                raise ValueError(f"Unexpected datetime type: {type(dt_value)}")
        except Exception as e:
            log(f"Error in parse_datetime: {e}")
            raise

    def create_appointment(self, calendar, event_data, retry=3):
        """Create calendar appointment with retries"""
        for attempt in range(retry):
            try:
                # Create the appointment
                appointment = calendar.Items.Add(1)  # 1 = olAppointmentItem
                
                # Set properties
                appointment.Subject = event_data['subject']
                appointment.Location = event_data['location']
                appointment.Body = event_data['description']
                appointment.ReminderSet = False
                appointment.Start = event_data['start'].strftime("%Y-%m-%d %H:%M")
                appointment.End = event_data['end'].strftime("%Y-%m-%d %H:%M")
                appointment.Save()
                
                # Log in concise format
                log(f"Event: {event_data['subject']} ({event_data['start'].strftime('%I:%M %p')} - {event_data['end'].strftime('%I:%M %p')})")
                return True

            except Exception as e:
                if attempt == retry - 1:  # Last attempt
                    raise Exception(f"Failed to create appointment: {str(e)}")
                log(f"Retrying appointment creation... ({attempt + 1}/{retry})")
                time.sleep(2)
                calendar = self.get_calendar()

        return False

    def sync(self):
        """Main sync function"""
        try:
            # Check for changes
            content = self.get_ics_content()
            if not content:
                return

            # Get calendar and clear it
            calendar = self.get_calendar()
            self.clear_calendar(calendar)
            log("Starting sync...")

            # Parse ICS content
            cal = icalendar.Calendar.from_ical(content)
            synced = 0
            errors = 0

            for component in cal.walk():
                if component.name != "VEVENT":
                    continue

                try:
                    # Get event details
                    summary = str(component.get('summary', 'Untitled Event'))
                    description = str(component.get('description', ''))
                    location = str(component.get('location', ''))

                    # Get timezone if specified
                    tz_id = None
                    if 'dtstart' in component:
                        tz_id = component['dtstart'].params.get('TZID')
                    event_tz = ZoneInfo(tz_id) if tz_id else None

                    # Get start/end times
                    start = component.get('dtstart').dt
                    end = component.get('dtend').dt if 'dtend' in component else start

                    # Convert to local time
                    start = self.parse_datetime(start, event_tz)
                    end = self.parse_datetime(end, event_tz)

                    # Create appointment
                    event_data = {
                        'subject': summary,
                        'start': start,
                        'end': end,
                        'location': location,
                        'description': description
                    }
                    
                    if self.create_appointment(calendar, event_data):
                        synced += 1
                    else:
                        errors += 1

                except Exception as e:
                    log(f"Error with event {component.get('summary', 'Unknown')}: {e}")
                    errors += 1

            log(f"Completed: {synced} events synced" + (f", {errors} errors" if errors > 0 else ""))

        except Exception as e:
            log(f"Sync failed: {e}")
            raise

def load_config():
    """Load configuration from config.json"""
    try:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
        with open(config_path, 'r') as f:
            config = json.load(f)
        return config['tripit_url']
    except Exception as e:
        log(f"Failed to load config.json: {e}")
        log("Please copy config.example.json to config.json and add your TripIt URL")
        sys.exit(1)

def main():
    tripit_url = load_config()
    try:
        syncer = TripitSync(tripit_url)
        syncer.sync()
    except Exception as e:
        log(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()