# Archive using 7zip

This script is used for archiving major data.

## Requirements
The script only needs to know where to find the CSV file which needs to be loaded.
This CSV path can be set in the global vars part of the script.

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
> .\7zip_to_netapp_v2.ps1

Just sit back and relax...
