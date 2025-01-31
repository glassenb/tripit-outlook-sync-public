@echo off
echo Starting TripIt Sync at %date% %time%
cd /d %~dp0
"C:\Python312\python.exe" "TripitSync.py" >> tripit_sync.log 2>&1
echo Finished at %date% %time%
echo ---------------------------------------- >> tripit_sync.log