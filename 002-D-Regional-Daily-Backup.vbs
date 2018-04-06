option Explicit

'=============================================================================================================
Const CScriptVersion	= "4.5.5"
Const CScriptDate		= "26/01/2011"
Const CScriptCommit		= ""
Const CScriptName		= "Regional Backup Script"
Const CScriptOwner		= "S. HASTIE"
'=============================================================================================================


'=============================================================================================================

On error resume next

'=============================================================================================================
'								Declarations
' BackupMode value:
'		0 = No copy
'		1 = Copy the backup only on the first remote server
'		2 = Copy the backup on both remote servers
'
' generalError value:
'		0 = No Error
'		1 = Warning
'		2 = Error
'=============================================================================================================


' ***  CONSTANTS  *** '
CONST ForReading = 1
CONST ForWriting = 2
CONST ForAppending = 8
CONST NBBACKUPTOKEEP = 12
Const HARD_DISK = 3


' ***  VARIABLES  *** '
dim oDrive, oFso,WshNetwork,strServerName, intWeekNumber, strFileName, intParity,strLocalServer,srvType
dim strWeekParity, strDayOfTheWeek, objShell, intDay, strLogFileFolder, strBackupUserName
dim regKey, fLog, strFileContent, oTs, fFolder,fFile, files
dim strScriptFileLog, strLog, strType,fNewLogFile,Return, strRemoteServer1, remotePath2, remotePath1
dim oArgs, strCMD, strRemoteServer2,monthlyBackup,weekBackup', boolDelete
dim month, backupMode, generalError, mailMessage,backuppath,backupDriveLetter, objWMIService
dim colDisks, objDisk, localUserDomain, deleteBackupCount
Dim monthRetention, monthRetention1, monthRetention2, weekRetention, weekRetention1, weekRetention2
dim dayRetention,dayRetention1,dayRetention2,backuppath1,backuppath2,backupDriveLetter2,backupDriveLetter1
Dim strProfileStatus, noBackup, MINIMUM_DISK_SPACE, arrTargets, i, strTargets, arrBackupPath
Dim TempLogFile, strTempContent

' ***  OBJECTS  *** '
Set oFso = CreateObject("Scripting.FileSystemObject")
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set objShell = Wscript.CreateObject("Wscript.Shell")
Set oArgs = wscript.arguments

strServerName = WshNetwork.ComputerName
strBackupUserName = WshNetwork.UserName
localUserDomain = WshNetwork.UserDomain

backupMode = 0
generalError = -1
mailMessage = ""
noBackup = 0


'=============================================================================================================
'						Get information from the INI file
'								and check validity
'=============================================================================================================
Err.clear
backuppath = GetiniV("LOCALHOST","backuppath")

If backuppath = "" Then
	backuppath = "e:\data\backup\"
End If

Dim scriptLogFilePath: scriptLogFilePath = backuppath & "logs\"
backupDriveLetter = UCase(Left(backuppath,InStr(backuppath,"\")-1))

' check to see if the logs need rotating
LogRotate scriptLogFilePath

TempLogFile = scriptLogFilePath & "temp.log"

WriteLog scriptLogFilePath & "general.log", "**** Backup process starting on " & strServerName _
	& " ****", "INFORMATION"
WriteLog scriptLogFilePath & "general.log", strProfileStatus, "INFORMATION"
WriteLog scriptLogFilePath & "general.log", CScriptName & " " & CScriptVersion & " (" & CScriptDate _
	& ")", "INFORMATION"

strLocalServer = GetiniV("LOCALHOST","servername")

If Err Then
	Err.Clear
End If

If strLocalServer = "" Then
	WriteLog scriptLogFilePath & "general.log", "strLocalServer not defined for the section LOCALHOST" _
		& " in the ini file. It will be defined as localhost.", "WARNING"
	strLocalServer = "localhost"
End if

Set oDrive = oFso.GetDrive(oFso.GetDriveName(backuppath))

monthRetention = GetiniV("LOCALHOST","monthretention")
If monthRetention = "" Then
	WriteLog scriptLogFilePath & "general.log", "MonthRetention not defined for the section LOCALHOST" _
		& " in the ini file. It will be defined as 6 months.", "WARNING"
	monthRetention = 6
End If

weekRetention = GetiniV("LOCALHOST","weekretention")
If weekRetention = "" Then
	WriteLog scriptLogFilePath & "general.log", "WeekRetention not defined for the section LOCALHOST" _
		& " in the ini file. It will be defined as 5 weeks.", "WARNING"
	weekRetention = 5
End If

dayRetention = GetiniV("LOCALHOST","dayretention")
If dayRetention = "" Then
	WriteLog scriptLogFilePath & "general.log", "DayRetention not defined for the section LOCALHOST" _
		& " in the ini file. It will be defined as 7 days.", "WARNING"
	dayRetention = 7
End If

srvType = GetiniV("LOCALHOST","type")
If srvType = "" Then
	WriteLog scriptLogFilePath & "general.log", "srvType not defined for the section LOCALHOST in the" _
		& " ini file. It will be defined as DC.", "WARNING"
	srvType = "DC"
End If

If instr(srvType, "2008") then

Else
	' *** Check the profile path of the account used to launch the backup (usually ssvc-scheduler)
	If oFso.FolderExists("C:\Documents and Settings\" & strBackupUserName & "." & localUserDomain) Then
		strLogFileFolder = "C:\Documents and Settings\" & strBackupUserName & "." & localUserDomain _
			& "\Local Settings\Application Data\Microsoft\Windows NT\NTBackup\data"
		strProfileStatus = "Profile exists, using: " & strLogFileFolder
	Else
		strLogFileFolder = "C:\Documents and Settings\" & strBackupUserName _
			& "\Local Settings\Application Data\Microsoft\Windows NT\NTBackup\data"
		strProfileStatus = "Profile doesn't exist, using: " & strLogFileFolder
	End If
End If

strRemoteServer1 = GetiniV("SERVER1","servername")
strRemoteServer2 = GetiniV("SERVER2","servername")
backupMode = 0

' *** If a Remote server is configured ***
If Len(strRemoteServer1) > 2 Then
	backuppath1 = GetiniV("SERVER1","backuppath")
	If backuppath1 = "" Then
		WriteLog scriptLogFilePath & "general.log", "Backuppath not defined for the SERVER1" _
			& " section in the ini file. It will be defined as e:\data\backup\.", "WARNING"
		backuppath1 = "e:\data\backup\"
	End If
	If InStr(UCase(backuppath1),"BACKUP") Then
		remotePath1 = Right(backuppath1,Len(backuppath1) - InStr(UCase(backuppath1),"BACKUP") + 2)
	Else
		remotePath1 = "\backup\"
	End If

	' *** Check if the remote path exists *** '
	If oFso.FolderExists("\\" & strRemoteServer1 & Left(remotePath1,Len(remotePath1) - 1)) Then
		WriteLog scriptLogFilePath & "general.log", "Remote Path exists on " & strRemoteServer1, _
			"INFORMATION"
	Else
		WriteLog scriptLogFilePath & "general.log", "Remote Path does not exist on " _
			& strRemoteServer1 & vbCrLf & "Creation of the remote path.", "INFORMATION"
		If oFso.CreateFolder("\\" & strRemoteServer1 & Left(remotePath1,Len(remotePath1) - 1)) Then
			WriteLog scriptLogFilePath & "general.log", "Creation of the remote path " _
				& "successful : \\" & strRemoteServer1 & Left(remotePath1,Len(remotePath1) - 1), "INFORMATION"
		Else
			WriteLog scriptLogFilePath & "general.log", "Creation of the remote path error : \\" _
				& strRemoteServer1 & Left(remotePath1,Len(remotePath1) - 1), "INFORMATION"
			WriteLog scriptLogFilePath & "general.log", "Remote Path will be defined as : \\" _
				& strRemoteServer1 & "\backup","WARNING"
			generalError = 1
		End If
	End If

	backupDriveLetter1 = UCase(Left(backuppath1,InStr(backuppath1,"\")-1))
	monthRetention1 = GetiniV("SERVER1","monthretention")

	If monthRetention1 = "" Then
		WriteLog scriptLogFilePath & "general.log", "MonthRetention not defined for the SERVER1 section " _
			& "in the ini file. It will be defined as 6 months.", "WARNING"
		monthRetention1 = 6
	End If

	weekRetention1 = GetiniV("SERVER1","weekretention")
	If weekRetention1 = "" Then
		WriteLog scriptLogFilePath & "general.log", "WeekRetention not defined for the SERVER1 section " _
			& "in the ini file. It will be defined as 5 weeks.", "WARNING"
		weekRetention1 = 5
	End If

	dayRetention1 = GetiniV("SERVER1","dayretention")
	If dayRetention1 = "" Then
		WriteLog scriptLogFilePath & "general.log", "DayRetention not defined for the SERVER1 section in " _
			& "the ini file. It will be defined as 7 days.", "WARNING"
		dayRetention1 = 7
	End If

	backupMode = 1

	' *** If a second Remote server is configured *** '
	If Len(strRemoteServer2) > 2 Then
		backuppath2 = GetiniV("SERVER2","backuppath")
		If backuppath2 = "" Then
			WriteLog scriptLogFilePath & "general.log", "Backuppath not defined for the SERVER2 section " _
				& "in the ini file. It will be defined as e:\data\backup\.", "WARNING"
			backuppath2 = "e:\data\backup\"
		End If

		If InStr(UCase(backuppath2),"BACKUP") Then
			remotePath2 = Right(backuppath2,Len(backuppath2) - InStr(UCase(backuppath2),"BACKUP") + 2)
		Else
			remotePath2 = "\backup\"
		End If
		
		' *** Check if the remote path exists *** '
		If oFso.FolderExists("\\" & strRemoteServer2 & Left(remotePath2,Len(remotePath2) - 1)) Then
			WriteLog scriptLogFilePath & "general.log", "Remote Path exists on " _
				& strRemoteServer2, "INFORMATION"
		Else
			WriteLog scriptLogFilePath & "general.log", "Remote Path does not exist on " _
				& strRemoteServer2 & vbCrLf & "Creation of the remote path.", "INFORMATION"
			If oFso.CreateFolder("\\" & strRemoteServer2 & Left(remotePath2,Len(remotePath2) - 1)) Then
				WriteLog scriptLogFilePath & "general.log", "Creation of the remote path successful : \\" _
					& strRemoteServer2 & Left(remotePath2,Len(remotePath2) - 1), "INFORMATION"
			Else
				WriteLog scriptLogFilePath & "general.log", "Creation of the remote path error : \\" _
					& strRemoteServer2 & Left(remotePath2,Len(remotePath2) - 1), "INFORMATION"
				WriteLog scriptLogFilePath & "general.log", "Remote Path will be defined to : \\" _
					& strRemoteServer2 & "\backup","WARNING"
				generalError = 1
			End If
		End If

		backupDriveLetter2 = UCase(Left(backuppath2,InStr(backuppath2,"\")-1))
		monthRetention2 = GetiniV("SERVER2","monthretention")
		If monthRetention2 = "" Then
			WriteLog scriptLogFilePath & "general.log", "MonthRetention not defined for the SERVER2 " _
				& "section in the ini file. It will be defined as 6 months.", "WARNING"
			monthRetention2 = 6
		End If

		weekRetention2 = GetiniV("SERVER2","weekretention")
		If weekRetention2 = "" Then
			WriteLog scriptLogFilePath & "general.log", "WeekRetention not defined for the SERVER2 " _
				& "section in the ini file. It will be defined as 5 weeks.", "WARNING"
			weekRetention2 = 5
		End If

		dayRetention2 = GetiniV("SERVER2","dayretention")
		If dayRetention2 = "" Then
			WriteLog scriptLogFilePath & "general.log", "DayRetention not defined for the SERVER2 " _
				& "section in the ini file. It will be defined as 7 days.", "WARNING"
			dayRetention2 = 7
		End If

		backupMode = 2

	Else
		dayRetention2 = 0
		weekRetention2 = 0
		monthRetention2 = 0
	End if
Else
	dayRetention1 = 0
	weekRetention1 = 0
	monthRetention1 = 0
End If


'=============================================================================================================
'						   Definitions, cleanup, etc.
'=============================================================================================================

' ***  Define Script path *** '
dim scriptFullpath,tempstr
scriptFullpath = ""
tempstr = wscript.ScriptFullName

while instr(tempstr ,"\")
	scriptFullpath = scriptFullpath & left(tempstr ,instr(tempstr ,"\"))
	tempstr = Right(tempstr, Len(tempstr) - instr(tempstr ,"\"))
Wend

If scriptFullpath = "" Then
	WriteLog scriptLogFilePath & "general.log", "scriptFullpath was not correctly defined. " _
		& "It will be defined as D:\Batch\.", "WARNING"
	scriptFullpath = "D:\Batch\"
End If


' ***  Create log folder if it doesn't exist  *** '
If not (oFso.FolderExists (scriptLogFilePath)) then
	Err.clear
	oFso.CreateFolder (scriptLogFilePath)
	If err then
		CreateEvent "ERROR", "Unable to create the LOG Folder! (" & scriptLogFilePath & ")", 96
		mailMessage = mailMessage & "Unable to create the LOG Folder! (" & scriptLogFilePath & ")" & vbcrlf
		generalError = 1
	Else
		WriteLog scriptLogFilePath & "general.log", "Log Folder Created successfully.", "INFORMATION"
	End if
End if


' ***  Get Week Number  *** '
err.clear
intWeekNumber = Cint (DatePart("ww", date))

If err then
	WriteLog scriptLogFilePath & "general.log", "Unable to get the Week Parity.", "ERROR"
	WScript.quit
End if

if instr(srvType, "2008") then

Else
	' ***  limit the number of NTBACKUP log files to only one  *** '
	regKey = "HKEY_CURRENT_USER\Software\Microsoft\Ntbackup\Log Files"
	Err.clear
	objShell.RegWrite regKey, 1, "REG_DWORD"

	If err.number <> 0 Then
		WriteLog scriptLogFilePath & "general.log", "Unable to Change the registry", "WARNING"
	End If
End if

' ***  Set the Week Parity  *** '
intParity =  intWeekNumber Mod 2
Select Case intParity
	Case 0
		strWeekParity = "Even"
	Case 1
		strWeekParity = "Odd"
	Case else
		WriteLog scriptLogFilePath & "general.log", "Week Parity Error.", "WARNING"
		WScript.Quit
End select


' ***  Get Day of the Week  *** '
intDay = DatePart("w", date)


' ***  Select the day of the week  *** '
Select Case intDay
	Case 1
		strDayOfTheWeek = "Sunday"
	Case 2
		strDayOfTheWeek = "Weekly"
		weekBackup = true
	Case 3
		strDayOfTheWeek = "Tuesday"
	Case 4
		strDayOfTheWeek = "Wednesday"
	Case 5
		strDayOfTheWeek = "Thursday"
	Case 6
		strDayOfTheWeek = "Friday"
	Case 7
		strDayOfTheWeek = "Saturday"
End select

If instr(srvType, "2008") Then

Else
	'  ***  Delete old ntbackup logfiles  *** '
	err.clear
	oFso.DeleteFolder(strLogFileFolder)
	If err then
		WriteLog scriptLogFilePath & "general.log", "Unable to purge Log file. (code 1)", "WARNING"
	End if

	err.clear
	oFso.CreateFolder(strLogFileFolder)
	If err then
		WriteLog scriptLogFilePath & "general.log", "Unable to purge Log file. (code 2)", "WARNING"
	Else
		WriteLog scriptLogFilePath & "general.log", "NTBackup Log files purged.", "INFORMATION"
	End If
End If

' *** Define backup name *** '   
Month= DatePart("m", date)

If month < 10 then
	Month = "0" & DatePart("m", date)
End If


'If it's a month backup
If DatePart("d",Date) = 1 Then
	strFileName = strServerName & "-Monthly_Backup" & "-" & month & "-" & DatePart("yyyy", date)
	monthlyBackup = true
Else
	strFileName = strServerName & "-W" & intWeekNumber & "-D-" & srvType & "-" & strDayOfTheWeek
End If





'=============================================================================================================
'							 Check available disk space
'								locally and remotely
'=============================================================================================================

deleteBackupCount = 0

'set MINIMUM_DISK_SPACE to the largest previous backup or 2Gigs
MINIMUM_DISK_SPACE = GetLargestBackupSize("localhost", "\backup\", scriptLogFilePath)

' ***  Check the free disk space locally  *** '
While oDrive.FreeSpace/1073741824 < MINIMUM_DISK_SPACE And deleteBackupCount < 2
	Err.Clear
	DeleteOlderBackupFile "localhost","\backup\", scriptLogFilePath
	deleteBackupCount = deleteBackupCount + 1
	Wscript.sleep 3000
Wend

If deleteBackupCount = 2 Then
	generalError = 1
	mailMessage = mailMessage & vbcrlf & "Unable to delete the old backup file on localhost." & vbcrlf
End If

If instr(srvType, "2008") Then

Else
	' ***  Check the free disk space on the first remote server *** '
	If backupMode = 1 Then
		If strRemoteServer1 <> "" Then
			Err.Clear
			Set objWMIService = GetObject("winmgmts:" & _
				"{impersonationLevel=impersonate}!\\" & strRemoteServer1 & "\root\cimv2")
			' If there is a WMI connection problem
			If Err Then
				CreateEvent "WARNING", "Can not connect to " & strRemoteServer1 & " with WMI.", 95
				Writelog scriptLogFilePath & "general.log", "Warning: Can not connect to " & strRemoteServer1 _
					& " with WMI.", "WARNING"
				mailMessage = mailMessage & "Warning: Can not connect to " & strRemoteServer1 & " with WMI." _
					& vbcrlf
				generalError = 1
				Err.Clear
			Else
				Set colDisks = objWMIService.ExecQuery _
					("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "") 
				If Err Then 
					CreateEvent "WARNING", "Can not retrieve disk information on " & strRemoteServer1 _
						& " with WMI.", 95
					Writelog scriptLogFilePath & "general.log", "Warning: Can not retrieve disk information on" _
						& strRemoteServer1 & " with WMI: Can not check disk space to copy the backup.", "WARNING"
					Err.Clear
				Else
					For Each objDisk in colDisks
						If objDisk.DeviceID = backupDriveLetter1 Then
							If (objDisk.FreeSpace /1073741824) < MINIMUM_DISK_SPACE Then
								DeleteOlderBackupFile strRemoteServer1,remotePath1
							End If
						End If
					Next
				End If
			End If
		End If
	End If


	' ***  Check the free disk space on the second remote server *** '
	If backupMode = 2 Then
		If strRemoteServer2 <> "" then
			Err.clear
			Set objWMIService = GetObject("winmgmts:\\" & _
				"{impersonationLevel=impersonate}!\\" & strRemoteServer2 & "\root\cimv2")
			' If there is a WMI connection problem
			If Err Then
				CreateEvent "WARNING", "Can not connect to " & strRemoteServer2 & " with WMI.", 95
				Writelog scriptLogFilePath & "general.log", "Warning: Can not connect to " & strRemoteServer2 _
					& " with WMI.","WARNING"
				mailMessage = mailMessage & "Warning: Can not connect to " & strRemoteServer2 & " with WMI." _
					& vbcrlf
				generalError = 1
				Err.Clear
			Else 
				Set colDisks = objWMIService.ExecQuery _
					("Select * from Win32_LogicalDisk Where DriveType = " &  HARD_DISK & "")
				If Err Then 
					CreateEvent "WARNING", "Can not retrieve disk information on " & strRemoteServer2 _
						& " with WMI.", 95
					Writelog scriptLogFilePath & "general.log", "Warning: Can not retrieve disk information on" _
						& strRemoteServer2 & " with WMI: Can not check disk space to copy the backup.", "WARNING"
					Err.Clear
				Else
					For Each objDisk in colDisks
						If objDisk.DeviceID = backupDriveLetter2 Then
							If (objDisk.FreeSpace /1073741824) < MINIMUM_DISK_SPACE Then
								DeleteOlderBackupFile strRemoteServer2,remotePath2
							End If
						End If
					Next
				End If
			End If
		End If
	End If

End If


'=============================================================================================================
'							  Launch the backup 
'=============================================================================================================

'If daily retention is defined to 0 and if it's a daily backup, do not need to backup ! 
If dayRetention = 0 And dayRetention1 = 0 And dayRetention2 = 0 And InStr(LCase(strFileName),"day") Then
	generalError = 0
	noBackup = 1
	mailMessage = mailMessage & "Do not need to backup today. Daily retention is defined to 0 and today" _
		& " is a daily backup!"
ElseIf (weekRetention = 0 And weekRetention1 = 0 And weekRetention2 = 0 And InStr(LCase(strFileName),"weekly")) Then
	generalError = 0
	noBackup = 1
	mailMessage = mailMessage & "Do not need to backup today. Weekly retention is defined to 0 and today" _
		& " is a weekly backup!"
ElseIf (monthRetention = 0 And monthRetention1 = 0 And monthRetention2 = 0 And InStr(LCase(strFileName),"month")) Then
	generalError = 0
	noBackup = 1
	mailMessage = mailMessage & "Do not need to backup today. Monthly retention is defined to 0 and today" _
		& " is a monthly backup!"  
ElseIf InStr(srvType, "2008") Then
	WriteLog scriptLogFilePath & "general.log", "Server identified as running Windows 2008", "INFORMATION"
	arrTargets = FileToArray(scriptFullpath & "Tools\001-D-" & srvType & ".bks", False)
	i = 0
	On Error Resume next
	While i <= UBound(arrTargets)
		If arrTargets(i) = "SystemState" Then
		Else
			If i = 0 Then
				strTargets = arrTargets(i)
			Else
				If arrTargets(i) <> "" Then
					strTargets = strTargets & "," & arrTargets(i)
				End If
			End If
		End If
		i = i + 1
	Wend
	On Error Goto 0
	arrBackupPath = Split(backuppath,"\")
	strCMD = "wbadmin start backup -backupTarget:" & arrBackupPath(0) & " -include:" & strTargets & " -quiet"
	WriteLog scriptLogFilePath & "general.log", "Backup started : " &  backuppath _
		& strFileName & "Backup", "INFORMATION"
	WriteLog scriptLogFilePath & "general.log", "Command line : " & strCMD, "INFORMATION"
	Return = objShell.run (strCMD, 7, true)
	oFso.CreateFolder  backuppath & strFileName
	oFso.MoveFolder arrBackupPath(0) & "\WindowsImageBackup", backuppath & strFileName & "\Backup"

	if Return = 0 then
		WriteLog scriptLogFilePath & "general.log", "Backup Finished. See the log file: " &  scriptLogFilePath _
			& strFileName & ".txt", "INFORMATION"
	else
		WriteLog scriptLogFilePath & "general.log", "BACKUP FAILED (error code:" & Return & ")." , "ERROR"
		mailMessage = mailMessage & "BACKUP FAILED (error code:" & Return & ")." & vbcrlf
		CreateEvent "ERROR", "BACKUP FAILED (error code:" & Return & ").", 96
		generalError = 2
	end if

	strCMD = "wbadmin start systemstatebackup -backupTarget:" & arrBackupPath(0) & " -quiet"
	WriteLog scriptLogFilePath & "general.log", "System State Backup started : " &  backuppath _
		& strFileName & "Backup", "INFORMATION"
	WriteLog scriptLogFilePath & "general.log", "Command line : " & strCMD, "INFORMATION"
	Return = objShell.run (strCMD, 7, true)
	oFso.MoveFolder arrBackupPath(0) & "\WindowsImageBackup", backuppath & strFileName & "\SystemState"

	if Return = 0 then
		WriteLog scriptLogFilePath & "general.log", "System State Backup Finished. See the log file: " &  scriptLogFilePath _
			& strFileName & ".txt", "INFORMATION"
	else
		WriteLog scriptLogFilePath & "general.log", "BACKUP FAILED (error code:" & Return & ")." , "ERROR"
		mailMessage = mailMessage & "BACKUP FAILED (error code:" & Return & ")." & vbcrlf
		CreateEvent "ERROR", "BACKUP FAILED (error code:" & Return & ").", 96
		generalError = 2
	end if

	if nobackup = 1 then
		WriteLog scriptLogFilePath & "general.log", "No backup done today, nothing to compress", "INFORMATION"
	else
		strCMD = scriptFullpath & "Tools\7z.exe a -t7z " & backuppath & strFileName & ".7z " & backuppath _
			& strFileName
		WriteLog scriptLogFilePath & "general.log", "Backup Compression Started : " &  backuppath _
			& strFileName, "INFORMATION"
		WriteLog scriptLogFilePath & "general.log", "Command line : " &  strCMD, "INFORMATION"
		Return = objShell.run (strCMD, 7, true)
		WriteLog scriptLogFilePath & "general.log", "Backup Compression Finished : " &  Return, "INFORMATION"
		oFso.DeleteFolder backuppath & strFileName
	end if

	If generalError = -1 Then
		generalError = 0
	End if

Else
	strCMD = "ntbackup.exe backup @" & scriptFullpath & "Tools\001-D-" & srvType & ".bks /n 001-D-" _
		& srvType & "-SystemBackup /d 001-D-" & srvType & "-SystemBackup /v:yes /r:no /rs:no /hc:off /m" _
			& " normal /j 001-D-" & srvType & "-SystemBackup /l:s /f " & backuppath & strFileName & ".bkf"
	WriteLog scriptLogFilePath & "general.log", "System State Backup started : " & backuppath _
		& strFileName & ".bkf", "INFORMATION"
	WriteLog scriptLogFilePath & "general.log", "Command line : " & strCMD, "INFORMATION"
	Return = objShell.run (strCMD, 7, true)

	if Return = 0 then
		Set oTs = oFso.GetFolder(strLogFileFolder)
		Set fFolder = oTs.Files

		For each files in fFolder
			Set fLog = oFso.GetFile(strLogFileFolder & "\" & files.name)
			Set fFile = fLog.OpenAsTextStream(ForReading,-1)
			strFileContent =  fFile.readAll
			generalError = AnalyseLog (strFileContent)
			fFile.close
		Next
		
		If generalError = -1 Then
			generalError = 0
		End if

		WriteLog scriptLogFilePath & "general.log", "Backup Finished. See the log file: " &  scriptLogFilePath _
			& strFileName & ".txt", "INFORMATION"
		if nobackup = 1 then
			WriteLog scriptLogFilePath & "general.log", "No backup done today, nothing to compress", "INFORMATION"
		else
			strCMD = scriptFullpath & "Tools\7z.exe a -t7z " & backuppath & strFileName & ".7z " & backuppath _
				& strFileName & ".bkf"
			WriteLog scriptLogFilePath & "general.log", "Backup Compression Started : " &  backuppath _
				& strFileName & ".bkf", "INFORMATION"
			WriteLog scriptLogFilePath & "general.log", "Command line : " &  strCMD, "INFORMATION"
			Return = objShell.run (strCMD, 7, true)
			WriteLog scriptLogFilePath & "general.log", "Backup Compression Finished : " & Return, "INFORMATION"
			oFso.DeleteFile backuppath & strFileName & ".bkf"
		end if
	else
		WriteLog scriptLogFilePath & "general.log", "BACKUP FAILED (error code:" & Return & ")." , "ERROR"
		mailMessage = mailMessage & "BACKUP FAILED (error code:" & Return & ")." & vbcrlf
		CreateEvent "ERROR", "BACKUP FAILED (error code:" & Return & ").", 96
		generalError = 2
	end if

end if
' *** If the backup was successful


' *** Copy the new backup on remote servers *** '
If noBackup = 1 Then
	WriteLog scriptLogFilePath & "general.log", "No backup done today, nothing to copy.", "INFORMATION"
	CreateEvent "INFORMATION", "No backup done today, nothing to copy.", 97
	mailMessage = mailMessage & vbCrLf & "No backup done today, nothing to copy."
Else
	If backupMode > 0 Then
		oFso.CopyFile backuppath & strFileName & ".7z","\\" & strRemoteServer1 & remotePath1 & strFileName _
			& ".7z",true 
		If Err Then
			WriteLog scriptLogFilePath & "general.log", "Unable to copy the backup file from " _
				& backuppath & strFileName & ".7z to \\" & strRemoteServer1 & remotePath1 & strFileName _
					& ".7z" & vbCrLf & "Error number " & Err.Number & " -- " & Err.description , "ERROR"
			CreateEvent "WARNING", "Unable to copy the backup file from " & backuppath & strFileName _
				& ".7z to \\" & strRemoteServer1 & remotePath1 & strFileName & ".7z" & vbCrLf _
					& "Error number " & Err.Number & " -- " & Err.description, 95
			mailMessage = mailMessage & "Unable to copy the backup file from " & backuppath & strFileName _
				& ".7z to \\" & strRemoteServer1 & remotePath1 & strFileName & ".7z" & vbCrLf _
					& "Error number " & Err.Number & " -- " & Err.description
		Else
			WriteLog scriptLogFilePath & "general.log", "Successfully copied the backup file from " _
				& backuppath & strFileName & ".7z to \\" & strRemoteServer1 & remotePath1 & strFileName _
					& ".7z", "INFORMATION"
			CreateEvent "INFORMATION", "Successfully copied the backup file from " & backuppath _
				& strFileName & ".7z to \\" & strRemoteServer1 & remotePath1 & strFileName & ".7z", 97
			mailMessage = mailMessage & "Successfully copied the backup file from " & backuppath _
				& strFileName & ".7z to \\" & strRemoteServer1 & remotePath1 & strFileName & ".7z"
		End If
		If backupMode = 2 Then
			oFso.CopyFile backuppath & strFileName & ".7z","\\" & strRemoteServer2 & remotePath2 _
				& strFileName & ".7z",true
			If Err Then
				WriteLog scriptLogFilePath & "general.log", "Unable to copy the backup file from " _
					& backuppath & strFileName & ".7z to \\" & strRemoteServer2 & remotePath2 & strFileName _
						& ".7z" & vbCrLf & "Error number " & Err.Number & " -- " & Err.description, "ERROR"
				CreateEvent "WARNING", "Unable to copy the backup file from " & backuppath & strFileName _
					& ".7z to \\" & strRemoteServer2 & remotePath2 & strFileName & ".7z" & vbCrLf _
						& "Error number " & Err.Number & " -- " & Err.description, 95
				mailMessage = mailMessage & "Unable to copy the backup file from " & backuppath _
					& strFileName & ".7z to \\" & strRemoteServer2 & remotePath2 & strFileName & ".7z " _
						& vbCrLf & "Error number " & Err.Number & " -- " & Err.description
			Else
				WriteLog scriptLogFilePath & "general.log", "Successfully copied the backup file from " _
					& backuppath & strFileName & ".7z to \\" & strRemoteServer2 & remotePath2 & strFileName _
						& ".7z", "INFORMATION"
				CreateEvent "INFORMATION", "Successfully copied the backup file from " & backuppath _
					& strFileName & ".7z to \\" & strRemoteServer2 & remotePath2 & strFileName & ".7z", 97
				mailMessage = mailMessage & "Successfully copied the backup file from " & backuppath _
					& strFileName & ".7z to \\" & strRemoteServer2 & remotePath2 & strFileName & ".7z"
			End If
		End If
	End If
End If


'=============================================================================================================
'							  Delete stale backups
'							  locally and remotely  
'=============================================================================================================


Err.Clear
' *** Delete stale backups locally *** '
DeleteStaleBackups 0
' *** Delete stale backups on the first server *** '
If backupMode > 0 Then
	DeleteStaleBackups 1
	' *** Delete stale backups on the second server *** '
	If backupMode = 2 Then
		DeleteStaleBackups 2
	End If
End If


'=============================================================================================================
'								Error handling
'=============================================================================================================

If generalError = 0 Then
	' ***  Copy the NTbackup log to the Log Folder
	set fNewLogFile = oFso.CreateTextFile(scriptLogFilePath & strFileName & ".txt", true)
	WriteLog scriptLogFilePath & "general.log", "**** Backup process ended successfully on " _
		& strServerName & " ****", "INFORMATION"
	fNewLogFile.writeline "SUCCESS"
	fNewLogFile.write strFileContent
	fNewLogFile.close
	if instr(srvType, "2008") Then
		SendMail "SUCCESS: " & strServerName & " Backup Log", "SUCCESS " & vbCrLf & mailMessage
	Else
		SendMail "SUCCESS: " & strServerName & " Backup Log", "SUCCESS " & vbCrLf & mailMessage & vbCrLf _
			& vbCrLf & strFileContent
	End If
ElseIf generalError = 1 Then
	' ***  Copy the NTbackup log to the Log Folder
	set fNewLogFile = oFso.CreateTextFile(scriptLogFilePath & strFileName & ".txt", true)
	WriteLog scriptLogFilePath & "general.log", "**** Backup process ended with a warning on " _
		& strServerName & " ****", "INFORMATION"
	fNewLogFile.writeline "WARNING"
	fNewLogFile.write strFileContent
	fNewLogFile.Close
	if instr(srvType, "2008") Then
		SendMail "WARNING: " & strServerName & " Backup Log", "WARNING" & vbCrLf & mailMessage
	Else
		SendMail "WARNING: " & strServerName & " Backup Log", "WARNING" & vbCrLf & mailMessage & vbCrLf _
			& vbCrLf & "NT Backup Logs :" & vbCrLf & strFileContent
	End If
Else
	' ***  Copy the NTbackup log to the Log Folder
	set fNewLogFile = oFso.CreateTextFile(scriptLogFilePath & strFileName & ".txt", true)
	WriteLog scriptLogFilePath & "general.log", "**** Backup Process ended with an error on " _
		& strServerName & " ****", "INFORMATION"
	fNewLogFile.writeline "ERROR"
	fNewLogFile.write strFileContent
	fNewLogFile.close
	if not instr(srvType, "2008") Then
		SendMail "ERROR: " & strServerName & " Backup Log", "ERROR" & vbCrLf & mailMessage
	Else
		SendMail "ERROR: " & strServerName & " Backup Log", "ERROR" & vbCrLf & mailMessage & vbCrLf _
			& vbCrLf & "NT Backup Logs :" & vbCrLf & strFileContent
	End If
End If

oFso.DeleteFile(TempLogFile)

'=============================================================================================================
' Purpose : Create an event on the specified computer
' IN: Type of event, Message, Destination Computer
' OUT : -
'=============================================================================================================
Sub CreateEvent( strType, strMsg, iID)
	dim oShell
	Set oShell = Wscript.CreateObject("Wscript.Shell")
	wscript.echo ("eventcreate /T " & strTYPE & " /ID " & iID & " /L APPLICATION /SO " & chr(34) _
		& "Regional Backup Script" & chr(34) & " /D " & strMsg)
	oShell.run ("eventcreate /T " & strTYPE & " /ID " & iID & " /L APPLICATION /SO " & chr(34) _
		& "Regional Backup Script" & chr(34) & " /D " & chr(34) & strMsg & chr(34))
End Sub


'=============================================================================================================
' Purpose : Write all logs in a file
'=============================================================================================================
Sub WriteLog (strScriptFileLog, strLog, strType)
	Dim fLogFich,dDate,strText
	err.Clear
	Set fLogFich = oFso.OpenTextFile(strScriptFileLog, 8, True)
	dDate=now()
	strText =dDate & ";" & strType & ";=> " & strLog
	fLogFich.WriteLine (StrText)
	fLogFich.Close
	Set fLogFich = oFso.OpenTextFile(TempLogFile, 8, True)
	fLogFich.WriteLine (StrText)
	fLogFich.Close
End sub


'=============================================================================================================
' Purpose : Delete the older backup file
'=============================================================================================================
Sub DeleteOlderBackupFile(strRemoteServer, remotePath, scriptLogFilePath)
	Dim aMyArray(), Index, fso, f, fc, f1, numEls, firstItem, lastSwap, value, indexLimit, deletedFile
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder("\\" & strRemoteServer & remotePath)
	Set fc = f.Files
	Index = 0
	' fill the array with the Last modified date
	For Each f1 in fc
		ReDim preserve aMyArray (Index)
		If right (f1.name,2) = "7z" or right(f1,name,3) = "bkf" Then
			aMyArray (Index) = f1.DateLastModified
			Index = Index + 1
		End if
	Next
	numEls = UBound(aMyArray)
	firstItem = LBound(aMyArray)
	lastSwap = numEls
	' Sort the older file
	Do
		indexLimit = lastSwap - 1
		lastSwap = 0
		For Index = firstItem To indexLimit
			value = aMyArray(Index)
			If value > aMyArray(Index + 1) Then
				' if the items are not in order, swap them
				aMyArray(Index) = aMyArray(Index + 1)
				aMyArray(Index + 1) = value
				lastSwap = Index
			End If
		Next
	Loop While lastSwap
	' Delete the older file
	For Each f1 in fc
		If f1.DateLastModified = aMyArray(0) and (right(f1.name,2) = "7z" or right(f1,name,3) = "bkf") then
			Err.Clear
			WriteLog scriptLogFilePath & "general.log", "Insufficient free space, deleting old backup: " _
				& f1.Name & " on " & strRemoteServer, "INFORMATION"   
			deletedFile = f1.Name
			fso.DeleteFile "\\" & strRemoteServer & remotePath & f1.Name
			If Err Then
				WriteLog scriptLogFilePath & "general.log", "Unable to delete the old backup file: \\" _
					& strRemoteServer & remotePath & f1.Name, "WARNING"
				CreateEvent "WARNING", "Unable to delete the old backup file on " & strRemoteServer, 95
				generalError = 1
				Err.Clear
			Else
				WriteLog scriptLogFilePath & "general.log", "Old backup file deleted successfully.", _
					"INFORMATION"
				CreateEvent "INFORMATION", "Old backup file successfully deleted on " & strRemoteServer, 97 
			End If
		End if
	Next
End Sub


'=============================================================================================================
' Purpose : Analyse backup log file and write events in the event log.
'=============================================================================================================
Function AnalyseLog(strLog)
	If (InStr(1,strLog, "Verify completed") <> 0) then
		CreateEvent "INFORMATION", "Regional Domain Controller Backup completed successfully.", 97
	ElseIf (InStr(1,strLog, "The operation did not successfully complete.") <> 0) then
		CreateEvent "ERROR", "Regional Domain Controller backup did not successfully complete.", 96
		AnalyseLog = 2
	Else
		CreateEvent "ERROR", "Unable to determine the Regional domain controller backup status.", 96
		AnalyseLog = 2
	End if
End Function


'=============================================================================================================
' Purpose : Get variable in an ini section and return the value
'			Return nothing if the variable or the section does not exist
'=============================================================================================================
Function GetiniV(sec, var)
	Dim oFsoINI,objFile , findSection, section, variable, findVariable, iniFile
	iniFile = Left(WScript.ScriptFullName,Len(WScript.ScriptFullName) - 4) & ".ini"
	findSection = False
	findVariable = False
	GetiniV = ""
	Set oFsoINI = CreateObject("Scripting.FileSystemObject")
	Set objFile = oFsoINI.OpenTextFile(iniFile,1)

	While Not objFile.AtEndOfStream AND Not findSection
		section = objFile.ReadLine
		If instr(section,"#") Then
		else
			If InStr(UCase(section),UCase(sec)) Then
				findSection = True
				While Not objFile.AtEndOfStream AND Not findVariable
					variable = objFile.ReadLine
					If InStr(UCase(variable),UCase(var)) Then
						findVariable = True
						GetiniV = Right(variable,Len(variable) - InStr(variable, "=") - 1)
					End If
				Wend
			End If
		End if
	Wend
	objFile.Close
End Function


'=============================================================================================================
' Purpose : send the backup log by email
'=============================================================================================================
Sub SendMail (MailSubject, MailText)
	Dim iMsg, iConf, Flds, Latestlog, Value
	Const cdoSendUsingPort = 2
	Const strSmartHost = "smtpgw.contoso.com"
	On Error Resume Next
		'Create the message object.
		Set iMsg = CreateObject("CDO.Message")
		'Create the configuration object.
		Set iConf = iMsg.Configuration
		'Set the fields of the configuration object to send by using SMTP through port 25.
		With iConf.Fields
			.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
			.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmartHost
			.Update
		End With

		With iMsg
			.To = "gdc-log@contoso.com; simon.hastie@contoso.com"
			.From = "RegionalBackupScript@contoso.com"
			.Subject = MailSubject
			.TextBody = "----------------------------------------" & vbNewLine & CScriptName & " " & _
				CScriptVersion & vbNewLine & CScriptCommit & vbNewLine & "-------------------------------" _
					& "---------" & vbNewLine & vbNewLine & MailText & vbNewLine & vbNewLine
			.AddAttachment TempLogFile
			.Send
		End With
		set iMsg = Nothing
	On Error GoTo 0
End Sub


'=============================================================================================================
' Purpose : Sets MINIMUM_DISK_SPACE to the size of the largest 
'			previous backup or 2Gigs
'=============================================================================================================
Function GetLargestBackupSize(strRemoteServer, remotePath, scriptLogFilePath)
	Dim fso, f, fc, f1, value
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder("\\" & strRemoteServer & remotePath)
	Set fc = f.Files

	' Set initial value to 2Gigs
	value = 2147483648
	For Each f1 In fc
		If Right (f1.Name,2) = "7z" or right(f1,name,3) = "bkf" Then
			' Find the biggest previous backup and assign it's size to
			' the value variable if it is bigger than 2Gigs
			If value < f1.Size Then
				value = f1.Size
			End If
		End If
	Next
	GetLargestBackupSize = (value/1073741824)
	WriteLog scriptLogFilePath & "general.log", "Setting required disk space to: " & Left _
		(GetLargestBackupSize,4) & "Gigs", "INFORMATION"
End Function



'=============================================================================================================
' Purpose : Rotate logs the first of every month
'=============================================================================================================
Sub LogRotate(scriptLogFilePath)
	Dim strComputer, objWMIService, colItems, objItem, objFSO
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_LocalTime")
	For Each objItem In colItems
		If objItem.Day = 1 Then
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			objFSO.MoveFile scriptLogFilePath & "general.log", scriptLogFilePath & objItem.Year & "-" _
				& objItem.Month -1 & " -- general.log"
			WriteLog scriptLogFilePath & "general.log", "Logs rotated, old log file renamed to: " _
				& objItem.Year & "-" & objItem.Month -1 & " -- general.log", "INFORMATION"
		End If
	Next
End Sub


'=============================================================================================================
' Purpose : Delete stale backups locally and remotely 
'=============================================================================================================
Sub DeleteStaleBackups(host)
	Dim   f, fc, f1,Index, bdel, myMonthRetention, myWeekRetention, myDayRetention, myBackupPath, fileName
	bdel = True
	Index = 0  
	Select Case host
		Case 0 'localhost
			myMonthRetention = monthRetention
			myWeekRetention = weekRetention
			myDayRetention = dayRetention
			myBackupPath = backuppath
		Case 1 'first remote server
			myMonthRetention = monthRetention1
			myWeekRetention = weekRetention1
			myDayRetention = dayRetention1
			myBackupPath = "\\" & strRemoteServer1 & remotePath1
		Case 2 'second remote server
			myMonthRetention = monthRetention2
			myWeekRetention = weekRetention2
			myDayRetention = dayRetention2
			myBackupPath = "\\" & strRemoteServer2 & remotePath2
	End Select  

	Set f = oFso.GetFolder(myBackupPath)
	Set fc = f.Files

	' If it is a monthly backup, we delete backup files older than 
	' (Month - MonthRetention variable in the ini file) date
	If monthlyBackup Then
		WriteLog scriptLogFilePath & "general.log", "Looking for stale monthly backups to delete in " _
			& myBackupPath, "INFORMATION" 
		For Each f1 in fc
			If InStr(UCase(f1.Name),"MONTHLY_BACKUP") then
			WriteLog scriptLogFilePath & "general.log", "Checking file " & f1.Name & ", " _
				& "last modified " & f1.DateLastModified & ", " & "retention limit " & DateAdd _
					("m", - myMonthRetention, Date), "INFORMATION"  
				If (DateDiff("d",f1.DateLastModified, DateAdd("m", - myMonthRetention, Date))) >= 0 Then
					Err.Clear
					fileName = f1.Name
					bdel = f1.Delete()
					If Err Or bdel Then
						Err.Clear
						fso.DeleteFile myBackupPath & f1.Name
						If Err Then
							WriteLog scriptLogFilePath & "general.log", "Unable to delete the stale " _
								& "backup file: " & myBackupPath & fileName , "ERROR"
							CreateEvent "WARNING", "Unable to delete the stale backup file: " _
								& myBackupPath & fileName , 95
							mailMessage = mailMessage & vbcrlf & "Unable to delete the stale backup file:" _
								& myBackupPath & fileName
							Err.Clear
						Else
							WriteLog scriptLogFilePath & "general.log", "The stale backup file " _
								& myBackupPath & fileName &" was removed successfully.", "INFORMATION"
						End If
					Else
						WriteLog scriptLogFilePath & "general.log", "The stale backup file " _
							& myBackupPath & fileName &" was removed successfully.", "INFORMATION"
					End If
				End If
			End if
			Index = Index + 1
			bdel = True
		next
	ElseIf weekBackup Then
		WriteLog scriptLogFilePath & "general.log", "Looking for stale weekly backups to delete in " _
			& myBackupPath, "INFORMATION" 
		For Each f1 in fc
			If InStr(UCase(f1.Name),"WEEKLY") then
				WriteLog scriptLogFilePath & "general.log", "Checking file " & f1.Name & ", " _
					& "last modified " & f1.DateLastModified & ", " & "retention limit " _
						& DateAdd("ww", - myWeekRetention, Date), "INFORMATION"   
				If (DateDiff("d",f1.DateLastModified, DateAdd("ww", - myWeekRetention, Date))) >= 0 Then
					Err.Clear
					fileName = f1.Name
					bdel = f1.Delete()
					If Err Or bdel Then
						Err.Clear
						fso.DeleteFile myBackupPath & f1.Name
						If Err Then
							WriteLog scriptLogFilePath & "general.log", "Unable to delete the stale " _
								& "backup file: " & myBackupPath & fileName , "ERROR"
							CreateEvent "WARNING", "Unable to delete the stale backup file: " _
								& myBackupPath & fileName , 95
							mailMessage = mailMessage & vbcrlf & "Unable to delete the stale backup file: " _
								& myBackupPath & fileName
							Err.Clear
						Else
							WriteLog scriptLogFilePath & "general.log", "The stale backup file " _
								& myBackupPath & fileName & " was removed successfully.", "INFORMATION"
						End If
					Else
						WriteLog scriptLogFilePath & "general.log", "The stale backup file " _
							& myBackupPath & fileName & " was removed successfully.", "INFORMATION"
					End If
				End If
			End If
			Index = Index + 1
			bdel = True
		Next
	Else 
		WriteLog scriptLogFilePath & "general.log", "Looking for stale daily backups to delete in " _
			& myBackupPath, "INFORMATION" 
		For Each f1 in fc
			If InStr(UCase(f1.Name),"MONTHLY") Then
			ElseIf InStr(UCase(f1.Name),"WEEKLY") Then
			ElseIf InStr(UCase(f1.Name),"-D-") then
				WriteLog scriptLogFilePath & "general.log", "Checking file " & f1.Name & ", " _
					& "last modified " & f1.DateLastModified & ", " & "retention limit " & DateAdd _
						("d", - myDayRetention, Date), "INFORMATION"
				If (DateDiff("d",f1.DateLastModified, DateAdd("d", - myDayRetention, Date))) >= 0 Then
					Err.Clear
					fileName = f1.Name
					bdel = f1.Delete()
					If Err Or bdel Then
						Err.Clear
						fso.DeleteFile myBackupPath & f1.Name
						If Err Then
							WriteLog scriptLogFilePath & "general.log", "Unable to delete the stale " _
								& "backup file: " & myBackupPath & fileName , "ERROR"
							CreateEvent "WARNING", "Unable to delete the stale backup file: " _
								& myBackupPath & fileName, 95
							mailMessage = mailMessage & vbcrlf & "Unable to delete the stale backup file: " _
								& myBackupPath & fileName
							Err.Clear
						Else
							WriteLog scriptLogFilePath & "general.log", "The stale backup file " _
								& myBackupPath & fileName & " was removed successfully.", "INFORMATION"
						End If
					Else
							WriteLog scriptLogFilePath & "general.log", "The stale backup file " _
								& myBackupPath & fileName & " was removed successfully.", "INFORMATION"
					End If
				End If
			End if
			Index = Index + 1
			bdel = True
		next
	End If
End Sub


Function FileToArray(ByVal strFile, ByVal blnUNICODE)
	Const FOR_READING = 1
	Dim objFSO, objTS, strContents
	FileToArray = Split("")
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(strFile) Then
		On Error Resume Next
		Set objTS = objFSO.OpenTextFile(strFile, FOR_READING, False, blnUNICODE)
		If Err = 0 Then
			strContents = objTS.ReadAll
			objTS.Close
			FileToArray = Split(strContents, vbNewLine)
		End If
	End If
End Function
