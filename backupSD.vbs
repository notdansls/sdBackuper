'Declare statements
Dim strVolumeName, strDriveLetter ,strSourcePath, strModifiedYear, strModifiedMonth, strModifiedDay, strHomePath, strPhotoFolder
Dim objFolder, objShell

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
set fsoDrives = objFSO.Drives

'Set variable values
strSourcePath = ""		' This is the folder on the SD Card where photos are stored
strDestinationPath = "" ' This is the folder we will be backing up to.
strSourceGUID = "{3774BBA4-1549-42D7-A915-BDCEAB130C3E}"
strDestinationGUID = "{E8F28F48-718C-4AD6-8999-18CD8BE4B5D5}"
boolFoundSourceDrive = false
boolFoundDestinationDrive = false



sub findSourceDrive
	' This sub will parse the drives until it finds the GUID referenced by strSourceGUID. 
	on error resume next
	for each objDrive in fsoDrives
		if objFSO.FileExists(objDrive.DriveLetter & ":\" & strSourceGUID) then
			wscript.echo "Source: Found on " & objDrive.DriveLetter & ":\."
			strSourcePath = objDrive.DriveLetter & ":\"
			boolFoundSourceDrive = true
		end if
	next
	if boolFoundSourceDrive = true then
		findDestinationDrive
	else
		wscript.echo "ERROR: Source SD Card not found, is it inserted?"
	end if
end sub



sub findDestinationDrive
	' This sub will parse the drives until it finds the GUID referenced by strDestinationGUID. 
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



sub checkScratch
	' In this sub we are going to check to see if the scratch folder exists.
	' if it does not, we will create it
	if objFSO.FolderExists(strDestinationPath) then
		wscript.echo "Good news, the destination path (" & strDestinationPath & ") exists."
		copyFolders
	else
		wscript.echo "the folder (" & strDestinationPath & " ) does not exist, creating it"
		Set createFolder = filesys.CreateFolder(strDestinationFolder)
	end if
end sub



sub copyFolders
	' This sub is responsible for copying the folders found in the root of the SD card
	' to the destination folder (scratch)
	on error resume next
	wscript.echo "Copying the following folder:"
	for each objFolder in objFSO.GetFolder(strSourcePath).SubFolders
		wscript.echo "  - " & objFolder
		'objFSO.CopyFolder objFolder, strDestinationPath & "\" & objFolder.name
	next	
	compressFiles
end sub



sub compressFiles
	' This sub handles the folder and compresses the content
	strRun = """c:\Program Files (x86)\7-zip\7z.exe""" & " a c:\open\2015-05-16-backup.7z c:\scratch"
	wscript.echo "Compressing archive..."
	'result = objShell.run( strRun, 0, True )
	wscript.echo "                      ...[Done]"
end sub


findSourceDrive
wscript.echo ""
wscript.echo "Backup is complete"