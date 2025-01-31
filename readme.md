# TripIt to Outlook Calendar Sync

Automatically syncs your TripIt itineraries to a dedicated calendar in Outlook.

## Features

- Creates a dedicated "TripIt" calendar in Outlook
- Syncs all TripIt events including flights, hotels, and other travel plans
- Handles timezone conversions correctly
- Efficient syncing - only updates when changes are detected
- No reminders by default
- Logs all operations with timestamps

## Requirements

- Python 3.x
- Outlook (tested with Outlook 365)
- Windows (uses win32com for Outlook integration)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/tripit-outlook-sync.git
```

2. Install required packages:
```bash
pip install requests icalendar pywin32 python-dateutil pytz
```

3. Update the TripIt ICS URL in `TripitSync.py`

## Usage

Run manually:
```bash
python TripitSync.py
```

Or set up as a scheduled task using the included batch file:
```bash
run_tripit_sync.bat
```

## Setting up Automated Sync

1. Save both `TripitSync.py` and `run_tripit_sync.bat` in a permanent location
2. Create a scheduled task in Windows Task Scheduler to run `run_tripit_sync.bat` hourly
3. Make sure to set the task to "Run whether user is logged on or not"

## License

MIT License

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.