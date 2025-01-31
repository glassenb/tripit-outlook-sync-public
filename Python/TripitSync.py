import requests
import icalendar
import datetime
import win32com.client
from dateutil import tz
import time
import hashlib
import json
import os
import sys

class TripitOutlookSync:
    def __init__(self, ics_url):
        self.ics_url = ics_url
        self.reset_outlook_connection()
        self.state_file = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), 'tripit_sync_state.json')
    
    def reset_outlook_connection(self):
        """Reset the Outlook COM connection"""
        try:
            self.outlook = win32com.client.Dispatch('Outlook.Application')
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            raise Exception(f"Failed to connect to Outlook: {str(e)}")
    
    def _get_or_create_tripit_calendar(self):
        """Get the TripIt calendar or create it if it doesn't exist"""
        root_folder = self.namespace.GetDefaultFolder(9)
        
        for folder in root_folder.Folders:
            if folder.Name == "TripIt":
                print("Found existing TripIt calendar")
                return folder
        
        tripit_calendar = root_folder.Folders.Add("TripIt")
        print("Created new TripIt calendar")
        return tripit_calendar
    
    def _get_ics_content(self):
        """Fetch the ICS content"""
        try:
            response = requests.get(self.ics_url)
            response.raise_for_status()
            return response.text
        except requests.exceptions.RequestException as e:
            raise Exception(f"Failed to fetch ICS feed: {str(e)}")
    
    def _get_content_hash(self, content):
        """Generate a hash of the ICS content"""
        return hashlib.sha256(content.encode('utf-8')).hexdigest()
    
    def _load_last_state(self):
        """Load the last known state from file"""
        try:
            if os.path.exists(self.state_file):
                with open(self.state_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Warning: Could not load last state: {str(e)}")
        return {'last_hash': None, 'last_sync': None}
    
    def _save_state(self, content_hash):
        """Save the current state to file"""
        try:
            state = {
                'last_hash': content_hash,
                'last_sync': datetime.datetime.now().isoformat()
            }
            with open(self.state_file, 'w') as f:
                json.dump(state, f)
        except Exception as e:
            print(f"Warning: Could not save state: {str(e)}")
    
    def needs_sync(self):
        """Check if sync is needed by comparing current ICS content with last known state"""
        try:
            # Get current ICS content
            current_content = self._get_ics_content()
            current_hash = self._get_content_hash(current_content)
            
            # Load last known state
            last_state = self._load_last_state()
            
            if last_state['last_hash'] == current_hash:
                last_sync = datetime.datetime.fromisoformat(last_state['last_sync']) if last_state['last_sync'] else None
                print(f"No changes detected in ICS feed. Last sync: {last_sync}")
                return False, current_content
                
            return True, current_content
            
        except Exception as e:
            print(f"Warning: Could not check sync status: {str(e)}")
            return True, None  # If we can't check, assume sync is needed
    
    def clear_calendar(self, calendar):
        """Remove all events from the calendar"""
        try:
            items = calendar.Items
            if items.Count > 0:
                for i in range(items.Count, 0, -1):
                    try:
                        items[i-1].Delete()
                    except:
                        continue
            print("Cleared existing calendar items")
        except Exception as e:
            print(f"Warning: Could not clear all calendar items. Continuing anyway. ({str(e)})")
        return calendar
    
    def sync_events(self):
        """Sync TripIt events to dedicated Outlook calendar if needed"""
        print("Checking if sync is needed...")
        
        should_sync, ics_content = self.needs_sync()
        if not should_sync:
            return
        
        print("Changes detected, starting sync...")
        
        try:
            # Get fresh calendar reference and clear it
            calendar = self._get_or_create_tripit_calendar()
            calendar = self.clear_calendar(calendar)
            
            # Parse ICS content
            cal = icalendar.Calendar.from_ical(ics_content)
            events = cal.walk('vevent')
            local_tz = tz.tzlocal()
            synced_count = 0
            
            for event in events:
                try:
                    # Extract event details
                    subject = str(event.get('summary'))
                    description = str(event.get('description', ''))
                    location = str(event.get('location', ''))
                    
                    # Convert times to local timezone
                    start = event.get('dtstart').dt
                    end = event.get('dtend').dt
                    
                    # Handle datetime vs date objects
                    if isinstance(start, datetime.date) and not isinstance(start, datetime.datetime):
                        start = datetime.datetime.combine(start, datetime.time.min)
                        end = datetime.datetime.combine(end, datetime.time.max)
                    
                    if start.tzinfo is None:
                        start = start.replace(tzinfo=local_tz)
                    if end.tzinfo is None:
                        end = end.replace(tzinfo=local_tz)
                    
                    # Create new appointment
                    appointment = calendar.Items.Add(1)  # 1 = olAppointmentItem
                    print(f"Creating event: {subject}")
                    
                    # Set appointment properties
                    appointment.Subject = subject
                    appointment.Start = start
                    appointment.End = end
                    appointment.Location = location
                    appointment.Body = description
                    appointment.ReminderSet = False
                    
                    # Save the appointment
                    appointment.Save()
                    synced_count += 1
                    
                except Exception as e:
                    print(f"Error processing event '{subject}': {str(e)}")
                    continue
            
            # Save state after successful sync
            self._save_state(self._get_content_hash(ics_content))
            
            print(f"Calendar sync completed! Synced {synced_count} events.")
            
        except Exception as e:
            print(f"Error during sync: {str(e)}")
            print("Attempting to reset connection and retry...")
            time.sleep(2)
            try:
                self.reset_outlook_connection()
                self.sync_events()  # Retry once
            except Exception as retry_error:
                print(f"Retry failed: {str(retry_error)}")

def main():
    # TripIt ICS URL
    tripit_url = "https://www.tripit.com/feed/ical/private/87AC5237-5F0628B3291C83F354859DD2E88969FC/tripit.ics"
    
    try:
        syncer = TripitOutlookSync(tripit_url)
        syncer.sync_events()
    except Exception as e:
        print(f"Fatal error during sync: {str(e)}")

if __name__ == "__main__":
    main()