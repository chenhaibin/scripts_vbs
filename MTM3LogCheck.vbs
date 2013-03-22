Option Explicit
On Error Resume Next
Dim arrThreadData, strCmd, tmpTarget, strHeader, strTMP

'User Configurable Variables
Const strComputer ="localhost"
Const MaxThreadCount = 15
Const ThreadMonitorSleep = 1000 'Miliseconds
Const LogFileToCheck = "c:\Program Files\Millennium 3\install.log"
Const EXEFileToCheck = "c:\Program Files\Millennium 3\Millennium.exe"
strCmd = "cmd.exe /c ""cscript //nologo " & wscript.scriptfullname & """ /t:"
strHeader = "Computer" & VBTAB & "WMI Query Result" & VBTAB & "OSQuery" & VBTAB & "Install.Log File Exists" & VBTAB & "Install.Log Last Modified" & VBTAB & "Millennium.exe Exists" & VBTAB & "Millennium.exe Version" & VBTAB & "Last Logged On User" 'Log file Header Line
Const strLogFile = "C:\Matt\1-05-10 M3 Log Check Post Upgrade.txt" 'Log File Location
Const rndOffsetUpper = 3000 'Maximum size of random offset for file write
Const rndOffsetLower = 300 'Minimum Size of random offset for file write
Const intMaxRetry = 100 'How many times to retry the File write
'End user Configurable Variables

strTMP = ""


tmpTarget = LCase(Wscript.Arguments.Named("T"))

If tmpTarget <> "" Then
	strTMP = TargetInfo(tmpTarget, LogFileToCheck, EXEFileToCheck)
	MTFileWriter strTMP, strLogFile, rndOffsetUpper, rndOffsetLower, intMaxRetry
Else
	arrThreadData = arrNetCompList
	MTFileWriter strHeader, strLogFile, rndOffsetUpper, rndOffsetLower, intMaxRetry 'Writes Header Line
	ThreadControlByArray strCmd, strComputer, MaxThreadCount, ThreadMonitorSleep, arrThreadData
End If

'****************************************ProcessMonitor
Private Function ProcessMonitor(StrComputer, PID)
Dim objWMIService
Dim colItems
Dim objItem
Dim PIDPresent
PIDPresent = False
Set objWMIService = GetObject("winmgmts:\\" & StrComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT ProcessId FROM Win32_Process",,48) 
For Each objItem in colItems 
	If PID = objItem.ProcessId Then
	PIDPresent = True
	End If
Next
ProcessMonitor = PIDPresent
End Function
'***************************************arrNetCompList
Private Function arrNetCompList()
Dim objShell
Dim objExec
Dim arrNetComputers()
Dim intCount1

Const NetViewCommand = "%comspec% /c net view | find ""\\"""
Const MaxNetBiosNameLength = 16

Set objShell = WScript.CreateObject("WScript.Shell")
Set objExec = objShell.Exec(NetViewCommand)'Runs a Net View Command filtered by Find
intCount1 = 0
Do Until objExec.StdOut.AtEndOfStream 'Cleans the net view output and loads it into the array
	ReDim Preserve arrNetComputers(intCount1)
	arrNetComputers(intCount1) = Rtrim(Left(Mid(objExec.StdOut.Readline(), 3), MaxNetBiosNameLength))
	intCount1 = intCount1 + 1
Loop
arrNetCompList = arrNetComputers
End Function
'*****************************CreateProcess
Function CreateProcess(strComputer, strCommand)
On Error Resume Next
Dim errReturn, objWMIService, intProcessID
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
	errReturn = objWMIService.Create (strCommand, Null, Null, intProcessID)
If errReturn = 0 Then
CreateProcess = intProcessID
Else
CreateProcess = -1
wscript.echo "Error Creating Process: [" & strCommand & "] on target system: [" & strComputer & "]"
End If
End Function
'******************************************OSQuery
Private Function OSQuery(strComputer)
On Error Resume Next
Dim objWMIService, objItem, colItems
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT Primary, Caption FROM Win32_OperatingSystem where Primary=true",,48) 
If Err = 0 Then
For Each objItem in colItems 
    OSQuery = objItem.Caption
Next
Else
OSQuery = "Not Found"
Err.Clear
End If
End Function

'**************************GetFileInfo
Function GetFileInfo(strComputer, FileToCheck)
Dim objWMIService, objItem, colFiles, ReturnValue(2)
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


Set colFiles = _
    objWMIService.ExecQuery("Select LastModified, Version From CIM_DataFile Where Name = '" & Replace(FileToCheck, "\", "\\") & "'")

If colFiles.Count < 1 Then
    ReturnValue(0) = "N"
    ReturnValue(1) = "NA"
    ReturnValue(2) = "NA"
Else
    ReturnValue(0) = "Y"
    for each objitem in colFiles
	ReturnValue(1) = WMIDateStringToDate(objItem.LastModified)
	ReturnValue(2) = objItem.Version
	Next
End If
GetFileInfo = ReturnValue
End Function
'*********************************TargetInfo
Private Function TargetInfo(strComputer, LogFileToCheck, EXEFileToCheck)
On Error Resume Next
Dim OSDetected, arrResult, strLine
	strLine = strComputer
	OSDetected = OSQuery(strComputer)
	If OSDetected = "Not Found" Then
	strLine = strLine & VBTAB & "Unsuccessful" & vbtab & "Not Found" & vbtab & "NA" & vbtab & "NA" & vbtab & "NA" & vbtab & "NA" & vbtab & "NA"
	Else
	strLine = strLine & VBTAB & "Successful" & vbtab & OSDetected & vbtab
	arrResult = GetFileInfo(strComputer, LogFileToCheck)
	strLine = strLine & arrResult(0) & vbtab & arrResult(1)
	arrResult = GetFileInfo(strComputer, EXEFileToCheck)
	strLine = strLine & vbtab & arrResult(0) & vbtab & arrResult(2)
	strLine = strLine & vbtab & GetLastLogon(strComputer, True)
	End If
	TargetInfo = strLine
End Function
'*****************************************ThreadControlByArray
Private Sub ThreadControlByArray(strCmd, strComputer, MaxThreadCount, ThreadMonitorSleep, arrThreadData)
On Error Resume Next
Dim intThreadDataCount, arrThreads(), ThreadCntr
ReDim arrThreads(MaxThreadCount - 1)
intThreadDataCount = 0

Do While intThreadDataCount <= Ubound(arrThreadData)
	
	For ThreadCntr = 0 to Ubound(arrThreads)
		If arrThreads(ThreadCntr) = "" Then
			'Debug:
			'wscript.echo "Max Process Count Not Reached, starting new"
			arrThreads(ThreadCntr)=CreateProcess(strComputer, strCmd & arrThreadData(intThreadDataCount))
			aWScript.stderr.writeline(intThreadDataCount + 1 & " of " & Ubound(arrThreadData) + 1 & " Executions completed | " & arrThreadData(intThreadDataCount))
			intThreadDataCount = intThreadDataCount + 1
		Else
			If Not(ProcessMonitor(strComputer, arrThreads(ThreadCntr))) Then
				arrThreads(ThreadCntr)=CreateProcess(strComputer, strCmd & arrThreadData(intThreadDataCount))
				WScript.stderr.writeline(intThreadDataCount + 1 & " of " & Ubound(arrThreadData) + 1 & " Executions completed | " & arrThreadData(intThreadDataCount))
				intThreadDataCount = intThreadDataCount + 1
				'Debug: 
				'wscript.echo "Monitored Process no longer running, starting new"
			End If
		End If
	Next
	'Debug: 
	'wscript.echo "Thread Manager Sleeping " & ThreadMonitorSleep/1000 & " Seconds"
	Wscript.sleep(ThreadMonitorSleep)
Loop
End Sub
'***************************MTFileWriter
Private Sub MTFileWriter(strLine, strLogFile, rndOffsetUpper, rndOffsetLower, intMaxRetry)
On Error Resume Next
Dim objFSO, objFile, rndOffset, booWriteFlag, cntIteration, Err

Const ForAppending = 8
Err.Clear
booWriteFlag = False

Do While ((booWriteFlag = False) And (intMaxRetry > cntIteration))
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strLogfile) Then
	Set objFile = objFSO.OpenTextFile(strLogFile, ForAppending)
Else
	Set objFile = objFSO.CreateTextFile(strLogfile)
	objFile.Close
	Set objFile = objFSO.OpenTextFile(strLogfile, ForAppending)
End If

If Err = 0 then
	objFile.WriteLine strLine
	objFile.Close
	booWriteFlag = True
Else
	wscript.echo Err
	Err.Clear
	Randomize
	rndOffset = Int((rndOffsetUpper - rndOffsetlower + 1) * Rnd + rndOffsetlower)
	wscript.sleep(rndOffset)
End If
	cntIteration = cntIteration + 1
Loop
objFile.Close
End Sub
'***************************WMIDateStringToDate
Function WMIDateStringToDate(dtmInstallDate)
 WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
 Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
 & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
 Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
 13, 2))
End Function
'***************************GetLastLogon
Private Function GetLastLogon(strComputer, OSPreVista)
'On Error Resume Next
Dim strKeyPath, objReg, subkey, arrsubkeys, RegCheck, ValueName

Const HKCR = &H80000000 'HKEY_CLASSES_ROOT
Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const HKU = &H80000003 'HKEY_USERS
Const HKCC = &H80000005 'HKEY_CURRENT_CONFIG
If OSPreVista Then
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
ValueName = "DefaultUserName"
Else
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
ValueName = "LastLoggedOnUser"
End If
Set objReg=GetObject("winmgmts:\\" & _
	strComputer & "\root\default:StdRegProv")
objReg.GetStringValue HKLM, strkeyPath, valuename, regcheck
	GetLastLogon = regcheck
End Function