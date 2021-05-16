'Declare statements
Dim strVolumeName, strDriveLetter ,strSourcePath, strModifiedYear, strModifiedMonth, strModifiedDay, strHomePath, strPhotoFolder
Dim objFolder, objShell


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
set fsoDrives = objFSO.Drives


'Set variable values
strVolumeName = "NIKON D7000"			' Name of the volume where photos are stored
strSourcePath = ""		' This is the folder on the SD Card where photos are stored
strDestinationPath = "" ' This is the folder we will be backing up to.
strHomePath = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
strPhotoFolder = strHomePath & "\OneDrive\Pictures\!inProgress\"
strSourceGUID = "{3774BBA4-1549-42D7-A915-BDCEAB130C3E}"
strDestinationGUID = "{E8F28F48-718C-4AD6-8999-18CD8BE4B5D5}"
boolFoundSourceDrive = false
boolFoundDestinationDrive = false

'processDrives(strVolumeName)

findSourceDrive

'psudo code
' 1) Read each drive letter
'   a) look for the guid {3774BBA4-1549-42D7-A915-BDCEAB130C3E}
'   b) tag the drive as the correct drive
sub findSourceDrive
	on error resume next
	for each objDrive in fsoDrives
		if objFSO.FileExists(objDrive.DriveLetter & ":\" & strSourceGUID) then
			wscript.echo "Source: Found on " & objDrive.DriveLetter & ":\."
			strSourcePath = objDrive.DriveLetter & ":\"
			boolFoundSourceDrive = true
		end if
	next
	if boolFoundSourceDrive = true then
		wscript.echo "Destination: Found at " & strSourcePath
		findDestinationDrive
	else
		wscript.echo "ERROR: Source SD Card not found, is it inserted?"
	end if
end sub

' 2) Next find the backup drive
'	a) look for the guid {E8F28F48-718C-4AD6-8999-18CD8BE4B5D5}
sub findDestinationDrive
	on error resume next
	for each objDrive in fsoDrives
		if objFSO.FileExists(objDrive.DriveLetter & ":\" & strDestinationGUID) then
			wscript.echo "Found DESTINATION on " & objDrive.DriveLetter & ":\."
			strDestinationPath = objDrive.DriveLetter & ":\scratch"
			boolFoundDestinationDrive = true
		end if
	next
	if boolFoundDestinationDrive = true then
		wscript.echo "We will be backing up to " & strDestinationPath
		checkScratch
	else
		wscript.echo "ERROR: Destination HDD no found, is it plugged in and powered?"
	end if
end sub

'	b) Create the scratch folder at the destination
sub checkScratch
'	In this sub we are going to check to see if the scratch folder exists.
'	if it does not, we will create it
	if objFSO.FolderExists(strDestinationPath) then
		wscript.echo "Good news, the destination path (" & strDestinationPath & ") exists."
		copyFolders
	else
		wscript.echo "the folder (" & strDestinationPath & " ) does not exist, creating it"
		Set createFolder = filesys.CreateFolder(strDestinationFolder)
	end if
end sub



sub copyFolders
'	This sub is responsible for copying the folders found in the root of the SD card
'	to the destination folder (scratch)
	on error resume next
	for each objFolder in objFSO.GetFolder(strSourcePath).SubFolders
		wscript.echo "Folder Name: " & objFolder
		objFSO.CopyFolder objFolder, strDestinationPath & "\" & objFolder.name
	next	
end sub



' 3) Give it a date and compress it

wscript.echo "Done"
