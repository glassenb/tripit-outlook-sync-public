# TripIt to Outlook Calendar Sync

Syncs your TripIt travel itinerary to a dedicated calendar in Outlook. Handles timezone conversions correctly and only updates when changes are detected.

## Features

- Creates a dedicated "TripIt" calendar in Outlook
- Proper timezone handling for flight times
- Efficient syncing (only updates when TripIt feed changes)
- Detailed logging
- Automatic retry on Outlook COM errors

## Requirements

- Python 3.9+
- Microsoft Outlook
- Windows

# Compatibility and Disclaimers

## Tested Environments
This script has only been tested with:
- Microsoft Outlook 365 (Classic Desktop version)
- Windows 11
- Python 3.9+

It may work with other versions of Outlook or Windows, but this has not been verified.

## ⚠️ Important Disclaimers

This script modifies your Outlook calendar by:
- Creating a dedicated "TripIt" calendar
- Adding, modifying, and removing events in this calendar
- Accessing your TripIt private calendar URL

While the script has been tested and includes error handling:
- Use at your own risk
- Back up your Outlook data before first use
- Monitor the calendar after initial setup to ensure events appear correctly
- The authors are not responsible for any data loss or calendar issues that may occur
- This is not an official TripIt or Microsoft tool

If you encounter issues:
1. Stop the script
2. Delete the TripIt calendar from Outlook if necessary
3. File an issue on GitHub

## Installation

1. Clone this repository
```bash
git clone https://github.com/yourusername/tripit-outlook-sync.git
cd tripit-outlook-sync
```

2. Install required packages
```bash
pip install requests icalendar pywin32 python-dateutil
```

3. Create a config file
```bash
cp config.example.json config.json
```

4. Edit `config.json` and add your TripIt private calendar URL:
- Go to https://www.tripit.com/account/calendar
- Copy your private calendar feed URL
- Paste it into config.json

## Usage

Run the script manually:
```bash
python tripit_sync.py
```

### Setting up Automatic Sync

To run the sync automatically:

1. Create a scheduled task in Windows Task Scheduler
2. Set it to run every hour (or your preferred interval)
3. Point it to run_tripit_sync.bat

## Security Notes

- Keep your config.json file private - it contains your TripIt private calendar URL
- The script stores a state file (tripit_sync_state.json) to track changes
- Both files are listed in .gitignore to prevent accidental commits

## Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)
