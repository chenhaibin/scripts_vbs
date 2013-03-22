Option Explicit
On Error Resume Next
Dim arrThreadData, strCmd, tmpTarget, strHeader,arrHeaderValues , strTMP, i

'User Configurable Variables
Const strComputer ="localhost"
Const MaxThreadCount = 10
Const ThreadMonitorSleep = 1000 'Miliseconds
strCmd = "cmd.exe /c ""cscript //nologo " & wscript.scriptfullname & """ /t:"
strHeader = "Computer,WMI Query Result,OSQuery,Dell Service Tag,virtualhost,Drive Letter,File System,Logical Partition Size, Freespace(MB), Freespace(%),Physical Layout from OS"
Const strLogFile = "C:\Users\steven.bambling\Desktop\Matt Scripts\LogicalDiskInventoryV2 12-8-09.tsv" 'Log File Location
Const rndOffsetUpper = 3000 'Maximum size of random offset for file write
Const rndOffsetLower = 300 'Minimum Size of random offset for file write
Const intMaxRetry = 100 'How many times to retry the File write
'End user Configurable Variables
strTMP = ""
arrHeaderValues = Split(strHeader,",")
strHeader = ""

For i = 0 to Ubound(arrHeaderValues)
	If i <> Ubound(arrHeaderValues) Then
		strHeader = strHeader & arrHeaderValues(i) & VBTAB
	else
		strHeader = strHeader & arrHeaderValues(i)
	End If
Next

strTMP = ""

tmpTarget = LCase(Wscript.Arguments.Named("T"))

If tmpTarget <> "" Then
	strTMP = TargetInfo(tmpTarget)
	MTFileWriter strTMP, strLogFile, rndOffsetUpper, rndOffsetLower, intMaxRetry
Else
	MTFileWriter strHeader, strLogFile, rndOffsetUpper, rndOffsetLower, intMaxRetry 'Writes Header Line
	arrThreadData = arrNetCompList
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
'******************************************GetST
Private Function GetST(strComputer)
On Error Resume Next
Dim objWMIService,objSTInfo, colSTInfo, strReturn
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colSTInfo = objWMIService.ExecQuery( _
    "SELECT SerialNumber FROM Win32_Bios",,48) 
For Each objSTInfo in colSTInfo 
strReturn =  objSTInfo.SerialNumber
Next
GetST = strReturn
End Function
'*********************************TargetInfo
Private Function TargetInfo(strComputer)
On Error Resume Next
Dim OSDetected, strResult, strLine, arrResult, i

	strLine = strComputer
	OSDetected = OSQuery(strComputer)
	If OSDetected = "Not Found" Then
		strLine = strLine & VBTAB & "Unsuccessful" & vbtab & "Not Found" & vbtab & "NA"
	Else
		strLine = strLine & VBTAB & "Successful" & vbtab & OSDetected & vbtab
		strResult = GetST(strComputer)
		If Not(RegExMatch(strResult, "^\w{5,7}$", "n")) Then
			strResult = "NA"
		End If
		strLine = strLine & strResult & vbtab
		strResult = GetVH(strComputer) & vbtab
		if len(strResult) = 0 Then
			strLine = strLine & "NA"
		Else
			strLine = strLine & strResult
		End If
		arrResult = BuildDiskDisplayString(DiskData(strComputer))
		strResult=""
		for i=0 to Ubound(arrResult)
			If i = Ubound(arrResult) Then
				strResult = strResult & strLine & arrResult(i)
			Else
				strResult = strResult & strLine & arrResult(i) & VBCRLF
			End If
		next
		strLine = strResult
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
'*****************************************RegExMatch
Private Function RegExMatch(strMessage, RegEx, CS)
Dim objRE,colMatches, flag
Set objRE = New RegExp
objRE.Global = True
If LCase(CS) = "n" Then
	objRE.ignorecase = true
Elseif LCase(CS) = "y" then
	objRE.ignorecase = false
End If
objRE.Pattern = RegEx
flag = objRE.Test(strMessage)
RegExMatch = flag
End Function
'******************************************GetVH
Private Function GetVH(strComputer)
'On Error Resume Next
Dim strKeyPath, objReg, subkey, arrsubkeys, RegCheck, ValueName

Const HKCR = &H80000000 'HKEY_CLASSES_ROOT
Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const HKU = &H80000003 'HKEY_USERS
Const HKCC = &H80000005 'HKEY_CURRENT_CONFIG

strKeyPath = "SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters"
ValueName = "Hostname"
Set objReg=GetObject("winmgmts:\\" & _
	strComputer & "\root\default:StdRegProv")
objReg.GetStringValue HKLM, strkeyPath, valuename, regcheck
	GetVH = regcheck
End Function
'*******************************************************
Private Function DiskData(strComputer)
On Error Resume Next
Dim colPartitions, colDisks, colVolumes, colPerfMons
Dim objWMI, objPartition, objDisk, objVolume, objPerfmon
Dim arrResult(), arrPhysicalDisk(), intResultCount, intPhysicalDiskCount

intPhysicalDiskCount = -1
intResultCount = -1
set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colVolumes = objWMI.ExecQuery("Select * From Win32_LogicalDisk Where drivetype = 3")

For Each objVolume in colVolumes
	Set colPartitions = objWMI.ExecQuery("Associators of {Win32_LogicalDisk.DeviceID='" & objVolume.DeviceID & _
	"'} WHERE RESULTCLASS = Win32_DiskPartition")

	intResultCount = intResultCount + 6
	Redim Preserve arrResult(intResultCount)

	arrResult(intResultCount - 5) = objVolume.DeviceID 'DriveLetter
	arrResult(intResultCount - 4) = objVolume.FileSystem 'FileSystem
	arrResult(intResultCount - 3) = FormatNumber((objVolume.Size / 1024) / 1024, 2) 'PartitionSize
	arrResult(intResultCount - 2) = FormatNumber((objVolume.Freespace / 1024) / 1024, 2) 'PartitionFreespace in MB

	Err.Clear
	Set colPerfMons = objWMI.ExecQuery( _
    	"SELECT PercentFreeSpace FROM Win32_PerfFormattedData_PerfDisk_LogicalDisk where Name = '" & objVolume.DeviceID & "'",,48) 
	For Each objPerfmon in colPerfMons
		IF Err <> 0 Then
			arrResult(intResultCount - 1) = "NA"
			Err.Clear
		Else
			arrResult(intResultCount - 1) = objPerfmon.PercentFreeSpace
		End If
	Next
	For Each objPartition in colPartitions
		intPhysicalDiskCount = intPhysicalDiskCount + 1

		Redim Preserve arrPhysicalDisk(1,intPHysicalDiskCount)

		arrPhysicalDisk(0,intPhysicalDiskCount) = objPartition.DeviceID

		Set colDisks = objWMI.ExecQuery("Associators of {Win32_DiskPartition.DeviceID='" & objPartition.DeviceID & _
		"'} WHERE RESULTCLASS = Win32_DiskDrive")
		For Each objDisk in colDisks

			arrPhysicalDisk(1,intPhysicalDiskCount) = objDisk.Caption
		Next
	Next
	arrResult(intResultCount) = arrPhysicalDisk
	intPhysicalDiskCount = -1
	Redim arrPHysicalDisk(1,0)
Next
DiskData = arrResult
End Function
'*******************************************************
Private Function BuildDiskDisplayString(arrResults)
Dim  Result, i, intFormatCounter, strTMP, strPhysicalDiskTMP, j, arrOutput(), k
intFormatCounter = 0
j = 0
k = 0
For each result in arrResults
	j = j + 1
	If IsArray(Result) Then
		For i=0 to Ubound(Result,2)

			strPhysicalDiskTMP = Result(0,i) & "[" & Result(1,i) & "]" & strPhysicalDiskTMP
			'wscript.echo "Physical Disk: " & Result(0,i)
			'wscript.echo "Caption: " & Result(1,i)
			If i < Ubound(Result,2) Then
				strPhysicalDiskTMP = "/" & strPhysicalDiskTMP
			End If

		Next
		strTMP = strTMP & vbtab & strPhysicalDiskTMP
		strPhysicalDiskTMP = ""
	Else
	Select Case intFormatCounter 
	Case 0 
	   'wscript.echo "Drive Letter: " & Result
		strTMP = Result
	Case 1 
	   'wscript.echo "File System: " & Result
		strTMP = strTMP & vbtab & Result
	Case 2 
	   'wscript.echo "Partion Size in MB: " & Result
		strTMP = strTMP & vbtab & Result
	Case 3
	   'wscript.echo "Freespace in MB: " & Result
		strTMP = strTMP & vbtab & Result
	Case 4
	   'wscript.echo "Percent Freespace: " & Result
		strTMP = strTMP & vbtab & Result
	End Select 
	End If
If intFormatCounter < 5 Then
	intFormatCounter = intFormatCounter +1
Else
	intFormatCounter = 0
	ReDim Preserve arrOutput(k)
	arrOutput(k) = strTMP
	strTMP = ""
	k = k + 1
End If
Next
BuildDiskDisplayString = arrOutput
End Function
'*******************************************************