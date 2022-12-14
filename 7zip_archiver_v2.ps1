# 7zip_archiver_v2.ps1
# Author: Meer-Web (info@meer-web.nl)
# Version 2.1.0

# Set global vars
$TIMESTAMP = Get-Date -Format "yyyyMMddHHmm"
$TARGET_FILENAME = "${TIMESTAMP}-archive.7z"
$CSV = "C:\Scripts\compress and move\dfs04zip.csv"
$LOGFILE = "C:\Scripts\compress and move\$TIMESTAMP.log"
write-host "===========================STARTING ARCHIVE==========================="

# Writelog function
if (!(test-path $LOGFILE)) {
    New-Item -ItemType File $LOGFILE
}
function WRITELOG {
    Param ([string]$LOGSTRING)
    $LOG_TIMESTAMP = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LOGMESSAGE = "${LOG_TIMESTAMP}: ${LOGSTRING}"
    Add-content $LOGFILE -value $LOGMESSAGE
}
WRITELOG "Archiving started"

# Check if CSV file exists
if (test-path $CSV) {
    WRITELOG "Loading CSV file $CSV"
    $CSVFILE = import-csv -Path 'C:\Scripts\compress and move\dfs04zip.csv'
} else {
    WRITELOG "CSV file not found! ($CSV)"
    Write-Output "$CSV does not exists!"
    exit
}

# Create alias
WRITELOG "Creating 7zip alias"
set-alias 7z "$env:ProgramFiles\7-Zip\7z.exe"

# Create empty folder for robocopy to mirror over
WRITELOG "Creating empty temp folder"
$EMPTY_FOLDER = "$env:TEMP\emptyfolder"
if (test-path $EMPTY_FOLDER) {
    Remove-Item $EMPTY_FOLDER -Force -Confirm:$false -Recurse
    New-Item -ItemType Directory $EMPTY_FOLDER
} else {
    New-Item -ItemType Directory $EMPTY_FOLDER
}

# START LOOPING TROUGH CSV FILE
$COUNTER = 0
WRITELOG "Looping through CSV file"
foreach ($SOURCE_ENTRY in $CSVFILE) {
    $COUNTER++
    $SOURCE = ${SOURCE_ENTRY}.FROM
    $TARGET = ${SOURCE_ENTRY}.TO
    WRITELOG "Archiving: $COUNTER - $SOURCE"
    Write-Output "Archiving $SOURCE"

    # Create stagefolders which give the state of the source folder
    Write-Output "Creating stage folder"
    $TEMPFOLDER_RUNNING = ${SOURCE} + "_being_archived_by_solvinity"
    $TEMPFOLDER_DONE = ${SOURCE} + "_is_archived_by_solvinity"
    $TEMPFOLDER_FAILED = ${SOURCE} + "_archived_failed_by_solvinity"
    WRITELOG "Creating stage folder $TEMPFOLDER_RUNNING"
    if (!(test-path $TEMPFOLDER_RUNNING)){
        New-Item -ItemType Directory $TEMPFOLDER_RUNNING
    }

    # Adjust ACL so that users cannot access the archived folder anymore
    WRITELOG "Locking ACL on $SOURCE"
    Write-Output "Set ACL on $SOURCE"
    $ACL = get-acl -Path $SOURCE
    $ACL.SetAccessRuleProtection($True, $False)
    $LOCALADMIN_FULLCONTROL = New-Object system.security.accesscontrol.filesystemaccessrule("builtin\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "none", "Allow") 
    $ACL.SetAccessRule($LOCALADMIN_FULLCONTROL)
    Set-Acl -Path $SOURCE -AclObject $ACL

    # Close down open files
    WRITELOG "Closing OpenFiles on $SOURCE"
    Write-Output "Closing open files"
    $OPENFILES = Get-SmbOpenFile | Where-Object -Property path -like *$SOURCE*
    $OPENFILES_COUNT = $OPENFILES.count
    WRITELOG "$OPENFILES_COUNT files closed"
    Write-Output "OpenFiles will be closed: $OPENFILES_COUNT"
    $OPENFILES | Close-SmbOpenFile -Force

    # Start archiving
    WRITELOG "7ZIP $SOURCE"
    Write-Output "7ZIP $SOURCE"
    7z a -mx3 -t7z -r "$TARGET\$TARGET_FILENAME" "$SOURCE\*"

    # Validate number of files in archive and DFS
    WRITELOG "Compare the number of files in DFS and archive"
    write-output "Compare the number of files in DFS and archive"
    ## 7zip
    WRITELOG "Counting files in 7zip archive"
    $7ZIP_FILECOUNT = '0'
    7z l $TARGET\$TARGET_FILENAME | Select-Object -Last 1 |
    Select-String '([0-9]+) files(?:, ([0-9]+) folders)?' |
    
    ForEach-Object {
        $7ZIP_FILECOUNT = [Int] $_.Matches[0].Groups[1].Value
    }
    
    ## DFS
    WRITELOG "Counting files in source path"
    $dirandfilelist = Get-ChildItem -Recurse -force $SOURCE
    $DFS_FILECOUNT = ($dirandfilelist | Where-Object { ! $_.PSIsContainer }).count
       
    Write-Output "${SOURCE}: ${7ZIP_FILECOUNT} / ${DFS_FILECOUNT}"
    WRITELOG "Comparing source and target files"
    if (($7ZIP_FILECOUNT -eq $DFS_FILECOUNT)) {
        WRITELOG "Number of files matching! Cleaning source folder"
        Write-Output "Cleaning up $SOURCE"
        robocopy /MIR $EMPTY_FOLDER $SOURCE
        remove-item -Force -Recurse $SOURCE
        WRITELOG "Source has been archived!"
        if (!(test-path $TEMPFOLDER_DONE)){
            WRITELOG "Renaming stage folder to $TEMPFOLDER_DONE"
            Rename-Item -Path $TEMPFOLDER_RUNNING -NewName $TEMPFOLDER_DONE
        }
    } else {
        WRITELOG "CRITICAL - Compare failed, the number of files are not matching. Counted: $7ZIP_FILECOUNT / $DFS_FILECOUNT"
        Write-Output "Skipping $SOURCE as file counts are not matching!"
        if (!(test-path $TEMPFOLDER_FAILED)){
            WRITELOG "CRITICAL - Renaming stage folder to $TEMPFOLDER_FAILED"
            Rename-Item -Path $TEMPFOLDER_RUNNING -NewName $TEMPFOLDER_FAILED
        }
    }
}
write-host "===========================ARCHIVE ENDED==========================="