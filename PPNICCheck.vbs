Option Explicit
'On Error Resume Next
Dim arrThreadData, strCmd, tmpTarget, strHeader, arrHeaderValues, strTMP, i

'User Configurable Variables
Const strLogFile = "C:\Matt\NICCheck7-14-09.log"
strHeader = "Computer,WMI Query Result,OSQuery,MAC_Address,IP_Addresses,Default_Gateways,DNS_Servers,Primary_WINS,Secondary_WINS,DHCP_Server,DHCP_Enabled,DHCP_Lease_Expires,DHCP_Lease_Obtained,Adapter_Speed,Adapter_MaxSpeed,Auto-negotiation_Enabled,NIC_Manufacturer,Adapter_Name,NetConnection_ID,Network_Connection_Status"

Const strComputer ="localhost"
Const MaxThreadCount = 15
Const ThreadMonitorSleep = 1000 'Miliseconds
strCmd = "cmd.exe /c ""cscript //nologo " & wscript.scriptfullname & """ /t:"
Const rndOffsetUpper = 3000 'Maximum size of random offset for file write
Const rndOffsetLower = 300 'Minimum Size of random offset for file write
Const intMaxRetry = 100 'How many times to retry the File write
'End user Configurable Variables

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
'*********************************TargetInfo
Private Function TargetInfo(strComputer)
'On Error Resume Next
Dim OSDetected, arrResult, strLine, strLineStart, j,k
	strLine = ""
	strLineStart = ""
	OSDetected = OSQuery(strComputer)
	If OSDetected = "Not Found" Then
	strLine = strComputer & VBTAB & "Unsuccessful" & vbtab & "Not Found"
	Else
	strLineStart = strComputer & VBTAB & "Successful" & vbtab & OSDetected
	arrResult = IPConfigData(strComputer)
	For j = 0 to Ubound(arrResult,2)
		strLine = strLine & strLineStart
		For k = 0 to Ubound(arrResult,1)
			strLine = strLine & VBTAB & arrResult(k,j)
		Next
		if j < Ubound(arrResult,2) Then
			strLine = strLine & VBCRLF
		End If
	Next
	
	End if
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
'*****************************
Private Function IPConfigData(strComputer)
On Error Resume Next
Dim objWMI, colConfigs, objConfig, colAdapters, objAdapter, h, i, strIPTMP, intTest, booErrFlag
Dim objNetConnectionStatusLookup
ReDim arrAdapterData(16,0)

Set objNetConnectionStatusLookup = CreateObject("Scripting.Dictionary")

objNetConnectionStatusLookup.Add 0, "Disconnected"
objNetConnectionStatusLookup.Add 1, "Connecting"
objNetConnectionStatusLookup.Add 2, "Connected"
objNetConnectionStatusLookup.Add 3, "Disconnecting"
objNetConnectionStatusLookup.Add 4, "Hardware not present"
objNetConnectionStatusLookup.Add 5, "Hardware disabled"
objNetConnectionStatusLookup.Add 6, "Hardware malfunction"
objNetConnectionStatusLookup.Add 7, "Media disconnected"
objNetConnectionStatusLookup.Add 8, "Authenticating"
objNetConnectionStatusLookup.Add 9, "Authentication succeeded"
objNetConnectionStatusLookup.Add 10, "Authentication failed"
objNetConnectionStatusLookup.Add 11, "Invalid address"
objNetConnectionStatusLookup.Add 12, "Credentials required"
set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colAdapters = objWMI.ExecQuery("Select * From Win32_NetworkAdapter Where NetConnectionStatus IS NOT NULL and ServiceName !='NIC1394'")
h = 0
intTest = 0
intTest = colAdapters.count 'used to test for WMI Nullish return
If Err<>0 Then
arrAdapterData(0,0) = "Error"
arrAdapterData(1,0) = "Win32_NetworkAdapater"
Err.Clear
Else
For Each objAdapter in colAdapters
	ReDim Preserve arrAdapterData(16,h)

	Set colConfigs = objWMI.ExecQuery("Associators of {Win32_NetworkAdapter.DeviceID='" & objAdapter.DeviceID & _
	"'} WHERE RESULTCLASS = Win32_NetworkAdapterConfiguration")
	intTest = colConfigs.count 'used to test for WMI Nullish return
	If Err<>0 Then
		arrAdapterData(0,0) = "Error"
		arrAdapterData(1,0) = "Win32_NetworkAdapaterConfiguration"
	Err.Clear
	Else
	For Each objConfig in colConfigs
		If NOT ISNull(objConfig.IPAddress) Then
		For i=0 to Ubound(objConfig.IPAddress)
			If i=0 Then
				strIPTMP = objConfig.IPAddress(i) & "\" & objConfig.IPSubnet(i)
			Else
				strIPTMP = strIPTMP & "," & objConfig.IPAddress(i) & "/" & objConfig.IPSubnet(i)
			End If
		Next
		Else
		strIPTMP = "-\-"
		End If

		arrAdapterData (0,h) = TestForNull(objAdapter.MACAddress) 			'MAC_Address
		arrAdapterData (1,h) = strIPTMP							'IP_Addresses
		arrAdapterData (2,h) = JoinArray(objConfig.DefaultIPGateway)			'Default_Gateways
		arrAdapterData (3,h) = JoinArray(objConfig.DNSServerSearchOrder)		'DNS_Servers
		arrAdapterData (4,h) = TestForNull(objConfig.WINSPrimaryServer)			'Primary_WINS
		arrAdapterData (5,h) = TestForNull(objConfig.WINSSecondaryServer)		'Secondary_WINS
		arrAdapterData (6,h) = TestForNull(objConfig.DHCPServer)			'DHCP_Server
		arrAdapterData (7,h) = TestForNull(objConfig.DHCPEnabled)			'DHCP_Enabled
		arrAdapterData (8,h) = WMIDateStringToDate(objConfig.DHCPLeaseExpires)	'DHCP_Lease_Expires
		arrAdapterData (9,h) = WMIDateStringToDate(objConfig.DHCPLeaseObtained)	'DHCP_Lease_Obtained
		arrAdapterData (10,h) = TestForNull(objAdapter.speed)				'Adapter_Speed
		arrAdapterData (11,h) = TestForNull(objAdapter.maxspeed)			'Adapter_MaxSpeed
		arrAdapterData (12,h) = TestForNull(objAdapter.autosense)			'Auto-negotiation_Enabled
		arrAdapterData (13,h) = TestForNull(objAdapter.Manufacturer)			'NIC_Manufacturer
		arrAdapterData (14,h) = TestForNull(objAdapter.name)				'Adapter_Name
		arrAdapterData (15,h) = TestForNull(objAdapter.NetConnectionID)			'NetConnection_ID
		If IsNull(objAdapter.NetConnectionStatus) Then					'Network_Connection_Status
			arrAdapterData (16,h) = TestForNull(objAdapter.NetConnectionStatus)
		Else
			arrAdapterData (16,h) = objNetConnectionStatusLookup.item(objAdapter.NetConnectionStatus)
		End If
		h = h + 1

	Next
	End If
Next
End If
IPConfigData = arrAdapterData
End Function

'*****************************
Function TestForNull(objToTest)
If Not IsNull(objToTest) Then
TestForNull = objToTest
Else
TestForNull = "-"
End If
End Function
'*****************************
Function JoinArray(arrToJoin)
If IsArray(arrToJoin) and Not IsNull(arrToJoin) Then
JoinArray = Join(arrToJoin, ",")
Else
JoinArray = "-"
End If
End Function
'*****************************
Function WMIDateStringToDate(strUTCDate)

If Not IsNull(strUTCDate) Then
 WMIDateStringToDate = CDate(Mid(strUTCDate, 5, 2) & "/" & _
 Mid(strUTCDate, 7, 2) & "/" & Left(strUTCDate, 4) _
 & " " & Mid (strUTCDate, 9, 2) & ":" & _
 Mid(strUTCDate, 11, 2) & ":" & Mid(strUTCDate, _
 13, 2))
Else
 WMIDateStringToDate = "-"
End If
End Function