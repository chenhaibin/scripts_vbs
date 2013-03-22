Option Explicit
On Error Resume Next
Dim arrThreadData, strCmd, tmpTarget, strHeader,arrHeaderValues , strTMP, i

'User Configurable Variables
Const strComputer ="localhost"
Const MaxThreadCount = 10
Const ThreadMonitorSleep = 1000 'Miliseconds
strCmd = "cmd.exe /c ""cscript //nologo " & wscript.scriptfullname & """ /t:"
strHeader = "Computer,WMI Query Result,OSQuery,Dell Service Tag,CPU Name,Win32_ComputerSystem.NumberofProcessors,RAM,VirtualHost,DELL:InstalledCPUs,DELL:CoreCountPerCPU,DELL:TotalSystemCores,Dell:ClockSpeed,Dell:SocketType,Dell:ProcessorFamily,DELL:VT_Supported,DELL:DBS_Supported,DELL:ExecuteDisable_Supported,DELL:HyperThreading_Supported,DELL:VT_Enabled,DELL:DBS_Enabled,DELL:ExecuteDisable_Enabled,DELL:HyperThreading_Enabled"
Const strLogFile = "C:\Users\steven.bambling\Desktop\Matt Scripts\BriefInventory 10-11-2009.txt" 'Log File Location
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
Dim OSDetected, strResult, strLine, arrReturn, objCursor, objDictionary

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
		strLine = strLine & GetCPUName(strComputer) & vbtab
		strLine = strLine & GetLogicalCPUCount(strComputer) & vbtab
		strLine = strLine & GetTotalRAM(strComputer) & vbtab
		strLine = strLine & GetVH(strComputer)
		'Get Dell CPU Information
		arrReturn = GetDellCPUData(strComputer)
		'Installed CPUs
		strLine =  strLine & vbtab & Ubound(arrReturn,2) + 1
		'Core Count Per CPU 
		strLine =  strLine & vbtab & arrReturn(0,0)
		'Total System Cores
		strLine =  strLine & vbtab & (Ubound(arrReturn,2) + 1) * arrReturn(0,0)
		'Clock Speed
		strLine =  strLine & vbtab & arrReturn(1,0) & "MHz"
		'Socket Type
		strLine =  strLine & vbtab & arrReturn(2,0)
		'Processor Family
		strLine =  strLine & vbtab & arrReturn(3,0)
		Set objDictionary = arrReturn(4,0)
		for each objCursor in objDictionary
			strLine =  strLine & vbtab & objDictionary(objCursor)
		Next
		Set objDictionary = arrReturn(5,0)
		for each objCursor in objDictionary
			strLine =  strLine & vbtab & objDictionary(objCursor)
		Next

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
'*******************************************************
Private Function GetCPUName(strComputer)
Dim objWMIService, colItems, objItem
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT Name FROM Win32_Processor",,48) 
For Each objItem in colItems 
    GetCPUName = objItem.Name
Next
End Function
'*******************************************************
Private Function GetLogicalCPUCount(strComputer)
Dim objWMIService, colItems, objItem
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT NumberofProcessors FROM Win32_ComputerSystem",,48) 
For Each objItem in colItems 
    GetLogicalCPUCount = objItem.NumberofProcessors
Next
End Function
'*******************************************************
Private Function GetTotalRAM(strComputer)
Dim objWMIService, colItems, objItem
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT TotalPhysicalMemory FROM Win32_ComputerSystem",,48) 
For Each objItem in colItems 
    GetTotalRAM = FormatNumber(objItem.TotalPhysicalMemory/1024/1024,0) & " MB"
Next
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
'*******************************************************
Private Function GetDellCPUData(strComputer)
Dim objWMIService, colItems, objItem, objCursor, objExtendedCharacteristics, objExtendedStates, intArrayCounter
Dim arrReturn() '[Corecount per CPU, Clock Speed, SocketType, Family, ExtendedCharacteristics(dictionary), ExtendedStates(dictionary)]
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Dell") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM CIM_Processor",,48) 

intArrayCounter = 0
Redim Preserve arrReturn(5,intArrayCounter)

For Each objItem in colItems 

	Redim Preserve arrReturn(5,intArrayCounter)
	
	arrReturn(0,intArrayCounter) = objItem.CoreCount
	arrReturn(1,intArrayCounter) = objItem.CurrentClockSpeed
	arrReturn(2,intArrayCounter) = ParseCPUUpgradeMethod(objItem.UpgradeMethod)
	arrReturn(3,intArrayCounter) = ParseCPUFamily(objItem.Family)
	Set arrReturn(4,intArrayCounter) = ParseCPUExtendedCharacteristics(objitem.ExtendedCharacteristics)
	Set arrReturn(5,intArrayCounter) = ParseCPUExtendedStates(objitem.ExtendedStates)
	intArrayCounter = intArrayCounter + 1
Next
GetDellCPUData = arrReturn
End Function
'*******************************************************
Private Function ParseCPUExtendedStates(intValue)
Dim objCPUStates, objCPUStateData, objState
Set objCPUStates = CreateObject("Scripting.Dictionary")
Set objCPUStateData = CreateObject("Scripting.Dictionary")

objCPUStates.Add "VT_Enabled", 1
objCPUStates.Add "DBS_Enabled", 2
objCPUStates.Add "ExecuteDisable_Enabled", 4
objCPUStates.Add "HyperThreading_Enabled", 8

for each objState in objCPUStates
If objCPUStates(objState) And intValue Then
	objCPUStateData.Add objState, True
Else
	objCPUStateData.Add objState, False
End If
Next
Set ParseCPUExtendedStates = objCPUStateData
End Function
'*******************************************************
Private Function ParseCPUUpgradeMethod(intValue)
Dim objCPUUpgradeMethod, objCPUUMData, objUM
Set objCPUUpgradeMethod = CreateObject("Scripting.Dictionary")

objCPUUpgradeMethod.Add "Other",1
objCPUUpgradeMethod.Add "Unknown",2
objCPUUpgradeMethod.Add "Daughter board",3
objCPUUpgradeMethod.Add "ZIF socket",4
objCPUUpgradeMethod.Add "Replacement/piggy back",5
objCPUUpgradeMethod.Add "None",6
objCPUUpgradeMethod.Add "LIF socket",7
objCPUUpgradeMethod.Add "Slot 1",8
objCPUUpgradeMethod.Add "Slot 2",9
objCPUUpgradeMethod.Add "370-pin socket",10
objCPUUpgradeMethod.Add "Socket mPGA604",19
objCPUUpgradeMethod.Add "Socket LGA771",20
objCPUUpgradeMethod.Add "Socket LGA775",21
objCPUUpgradeMethod.Add "Socket S1",22
objCPUUpgradeMethod.Add "Socket AM2",23
objCPUUpgradeMethod.Add "Socket F (1207)",24
objCPUUpgradeMethod.Add "Socket LGA1366",25

for each objUM in objCPUUpgrademethod
If objCPUUpgradeMethod(objUM) = intValue Then
	objCPUUMData = objUM
End If
Next
ParseCPUUpgradeMethod = objCPUUMData
End Function
'*******************************************************
Private Function ParseCPUFamily(intValue)
Dim objCPUFamilies, objCPUFamily, objCPUFamilyData
Set objCPUFamilies = CreateObject("Scripting.Dictionary")

objCPUFamilies.Add "Other",1
objCPUFamilies.Add "Unknown",2
objCPUFamilies.Add "8086",3
objCPUFamilies.Add "80286",4
objCPUFamilies.Add "80386",5
objCPUFamilies.Add "80486",6
objCPUFamilies.Add "8087",7
objCPUFamilies.Add "80287",8
objCPUFamilies.Add "80387",9
objCPUFamilies.Add "80487",10
objCPUFamilies.Add "Pentium® Brand",11
objCPUFamilies.Add "Pentium Pro",12
objCPUFamilies.Add "Pentium II",13
objCPUFamilies.Add "Pentium processor with MMX™ technology",14
objCPUFamilies.Add "Celeron™",15
objCPUFamilies.Add "Pentium II Xeon",16
objCPUFamilies.Add "Pentium III",17
objCPUFamilies.Add "M1 family",18
objCPUFamilies.Add "M2 family",19
objCPUFamilies.Add "AMD® Duron™ Processor",24
objCPUFamilies.Add "K5 family",25
objCPUFamilies.Add "K6 family",26
objCPUFamilies.Add "K6-2",27
objCPUFamilies.Add "K6-3",28
objCPUFamilies.Add "AMD Athlon™ Processor Family",29
objCPUFamilies.Add "AMD29000 Family",30
objCPUFamilies.Add "K6-2+",31
objCPUFamilies.Add "Power PC Family",32
objCPUFamilies.Add "Power PC 601",33
objCPUFamilies.Add "Power PC 603",34
objCPUFamilies.Add "Power PC 603+",35
objCPUFamilies.Add "Power PC 604",36
objCPUFamilies.Add "Power PC 620",37
objCPUFamilies.Add "Power PC X704",38
objCPUFamilies.Add "Power PC 750",39
objCPUFamilies.Add "Alpha Family",48
objCPUFamilies.Add "Alpha 21064",49
objCPUFamilies.Add "Alpha 21066",50
objCPUFamilies.Add "Alpha 21164",51
objCPUFamilies.Add "Alpha 21164PC",52
objCPUFamilies.Add "Alpha 21164a",53
objCPUFamilies.Add "Alpha 21264",54
objCPUFamilies.Add "Alpha 21364",55
objCPUFamilies.Add "MIPS Family",64
objCPUFamilies.Add "MIPS R4000",65
objCPUFamilies.Add "MIPS R4200",66
objCPUFamilies.Add "MIPSR4400",67
objCPUFamilies.Add "MIPS R4600",68
objCPUFamilies.Add "MIPS R10000",69
objCPUFamilies.Add "SPARC Family",80
objCPUFamilies.Add "SuperSPARC",81
objCPUFamilies.Add "microSPARC II",82
objCPUFamilies.Add "microSPARC IIep",83
objCPUFamilies.Add "UltraSPARC",84
objCPUFamilies.Add "UltraSPARC II",85
objCPUFamilies.Add "UltraSPARC IIi",86
objCPUFamilies.Add "UltraSPARC III",87
objCPUFamilies.Add "UltraSPARC IIIi",88
objCPUFamilies.Add "68040",96
objCPUFamilies.Add "68xxx Family",97
objCPUFamilies.Add "68000",98
objCPUFamilies.Add "68010",99
objCPUFamilies.Add "68020",100
objCPUFamilies.Add "68030",101
objCPUFamilies.Add "Hobbit family",112
objCPUFamilies.Add "Crusoe™ 5000 Family",120
objCPUFamilies.Add "Crusoe 3000 Family",121
objCPUFamilies.Add "Efficeon™8000 Family",122
objCPUFamilies.Add "Weitek",128
objCPUFamilies.Add "Itanium™ Processor",130
objCPUFamilies.Add "AMD Athlon 64 Processor Family",131
objCPUFamilies.Add "AMD Opteron™ Processor Family",132
objCPUFamilies.Add "AMD Sempron Processor Family",133
objCPUFamilies.Add "AMD Turion™ 64 Mobile Technology",134
objCPUFamilies.Add "Dual-Core AMD Opteron Processor family",135
objCPUFamilies.Add "AMD Athlon 64 X2 Dual-Core Processor family",136
objCPUFamilies.Add "AMD Turion 64 X2 Mobile Technology",137
objCPUFamilies.Add "Quad-Core AMD Opteron Processor Family",138
objCPUFamilies.Add "Third-Generation AMD Opteron Processor Family",139
objCPUFamilies.Add "PA-RISC family",144
objCPUFamilies.Add "PA-RISC 8500",145
objCPUFamilies.Add "PA-RISC 8000",146
objCPUFamilies.Add "PA-RISC 7300LC",147
objCPUFamilies.Add "PA-RISC 7200",148
objCPUFamilies.Add "PA-RISC 7100LC",149
objCPUFamilies.Add "PA-RISC 7100",150
objCPUFamilies.Add "V30 family",160
objCPUFamilies.Add "Dual-Core Intel® Xeon processor 5200 Series",171
objCPUFamilies.Add "Dual-Core Intel Xeon processor 7200 Series",172
objCPUFamilies.Add "Quad-Core Intel Xeon processor 7300 Series",173
objCPUFamilies.Add "Quad-Core Intel Xeon processor 7400 Series",174
objCPUFamilies.Add "Multi-Core Intel Xeon processor 7400 Series",175
objCPUFamilies.Add "Pentium® III Xeon",176
objCPUFamilies.Add "Pentium III Processor with Intel SpeedStep™",177
objCPUFamilies.Add "Technology",178
objCPUFamilies.Add "Pentium 4",179
objCPUFamilies.Add "Intel Xeon",180
objCPUFamilies.Add "AS400 Family",181
objCPUFamilies.Add "Intel Xeon Processor MP",182
objCPUFamilies.Add "AMD Athlon XP family",183
objCPUFamilies.Add "AMD Athlon MP family",184
objCPUFamilies.Add "Intel Itanium 2",185
objCPUFamilies.Add "Intel Pentium M processor",186
objCPUFamilies.Add "Intel Celeron D Processor",187
objCPUFamilies.Add "Intel Pentium D Processor",188
objCPUFamilies.Add "Intel Pentium Extreme Edition processor",189
objCPUFamilies.Add "Intel Core 2 processor",190
objCPUFamilies.Add "Intel Core i7 Processor",198
objCPUFamilies.Add "Dual-Core Intel Celeron Processor",199
objCPUFamilies.Add "S/390 and zSeries family",200
objCPUFamilies.Add "ESA/390 G4",201
objCPUFamilies.Add "ESA/390 G5",202
objCPUFamilies.Add "ESA/390 G6",203
objCPUFamilies.Add "z/Architecture base",204
objCPUFamilies.Add "Multi-Core Intel Xeon® processor",214
objCPUFamilies.Add "Dual-Core Intel Xeon processor 3xxx Series",215
objCPUFamilies.Add "Quad-Core Intel Xeon processor 3xxx Series",216
objCPUFamilies.Add "Dual-Core Intel Xeon processor 5xxx Series",218
objCPUFamilies.Add "Quad-Core Intel Xeon processor 5xxx Series",219
objCPUFamilies.Add "Dual-Core Intel Xeon processor 7xxx Series",221
objCPUFamilies.Add "Quad-Core Intel Xeon processor 7xxx Series",222
objCPUFamilies.Add "Multi-Core Intel Xeon processor 7xxx Series",223
objCPUFamilies.Add "Embedded AMD Opteron Quad-Core Processor Family",230
objCPUFamilies.Add "AMD Phenom™ Triple-Core Processor Family",231
objCPUFamilies.Add "AMD Turion Ultra Dual-Core Mobile Processor Family",232
objCPUFamilies.Add "AMD Turion Dual-Core Mobile Processor Family",233
objCPUFamilies.Add "AMD Athlon Dual-Core Processor Family",234
objCPUFamilies.Add "AMD Sempron™ SI Processor Family",235
objCPUFamilies.Add "AMD Opteron Six-Core Processor Family",238
objCPUFamilies.Add "i860™",250
objCPUFamilies.Add "i960™",251
objCPUFamilies.Add "SH-3",260
objCPUFamilies.Add "SH-4",261
objCPUFamilies.Add "ARM",280
objCPUFamilies.Add "StrongARM",281
objCPUFamilies.Add "6x86",300
objCPUFamilies.Add "MediaGX",301
objCPUFamilies.Add "MII",302
objCPUFamilies.Add "WinChip",320
objCPUFamilies.Add "DSP",350
objCPUFamilies.Add "Video processor",500


for each objCPUFamily in objCPUFamilies
If objCPUFamilies(objCPUFamily) = intValue Then
	objCPUFamilyData = objCPUFamily
End If
Next
ParseCPUFamily = objCPUFamilyData
End Function
'*******************************************************
Private Function ParseCPUExtendedCharacteristics(intValue)
Dim objCPUCharacteristics, objCPUData, objCPUCharacteristic
Set objCPUCharacteristics = CreateObject("Scripting.Dictionary")
Set objCPUData = CreateObject("Scripting.Dictionary")

objCPUCharacteristics.Add "VT_Supported", 1
objCPUCharacteristics.Add "DBS_Supported", 2
objCPUCharacteristics.Add "ExecuteDisable_Supported", 4
objCPUCharacteristics.Add "HyperThreading_Supported", 8

for each objCPUCharacteristic in objCPUCharacteristics
If objCPUCharacteristics(objCPUCharacteristic) And intValue Then
	objCPUData.Add objCPUCharacteristic, True
Else
	objCPUData.Add objCPUCharacteristic, False
End If
Next
Set ParseCPUExtendedCharacteristics = objCPUData
End Function