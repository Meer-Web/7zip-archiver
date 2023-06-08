# Archive using 7zip

This script is used for archiving major data.

## Requirements
The script needs to know where to find the CSV file which needs to be loaded.
This CSV path can be set in the global vars part of the script.
- 7zip needs to be installed
- robocopy needs to be installed

## Configuration
$MAX_FOLDER_SIZE = Max folder size to zip, otherwise switch to robocopy.
$TIMESTAMP = Timeformat to use for the logfile
$TARGET_FILENAME = Target archive name
$CSVFILE = Source CSV file
$LOGFILE = Log file
$TEMPFOLDER_RUNNING = Source folder name for when archive is running
$TEMPFOLDER_DONE = Source folder name for when archive is done
$TEMPFOLDER_FAILED = Source folder name for when archive is failed

### CSV template
Create a CSV file containing the following setup:
```
FROM,TO
Source path, Target path
```
You can add multiple lines where the script loops through.

Example:
```
FROM,TO
C:\temp\source_folder,D:\archive\target_folder
\\nas.local\documents,\\nas.local\archive
```

## Run script
> .\7zip_to_netapp.ps1

Just sit back and relax...