Option Explicit
Dim strComputer, arrTMP, j, arrResults, Result, NestedResult, disktmp

strComputer = "term-server4"
disktmp = DiskData(strComputer)
Build disktmp, "test"
'*******************************************************
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
'*******************************************************
Private Function GetLastLogon(strComputer, OSPreVista)
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
Function UTCDateStrToSQLDateTime(strUTCDate)
Dim objRE,colMatches, strReturn

Set objRE = New RegExp
objRE.Global = True
objRE.ignorecase = false
objRE.Pattern = "^(\d{8})(\d{2})(\d{2})(.{6})"
strReturn = ""
Set colMatches = objRE.Execute(strUTCDate)
If colMatches.count >=1 Then
		If  (colMatches(0).submatches.count >= 4)Then
			strReturn = 	colMatches(0).submatches(0) & " " & _
							colMatches(0).submatches(1) & ":" & _
							colMatches(0).submatches(2) & ":" & _
							colMatches(0).submatches(3)
		End If
End If
UTCDateStrToSQLDateTime = strReturn
End Function
'*******************************************************
Private Function GetInstalledApps(strComputer)
Dim objReg, strSubKey, arrSubKeys, errCheck, strValue, arrReturn(), i

Const HKLM = &H80000002
Const strBaseKey = "Software\Microsoft\Windows\CurrentVersion\Uninstall\"

Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")

objReg.EnumKey HKLM, strBaseKey, arrSubKeys
i=0
For Each strSubKey In arrSubKeys

    errCheck = objReg.GetStringValue(HKLM, strBaseKey & strSubKey, "DisplayName", strValue)

    If errCheck <> 0 Then
        errCheck = objReg.GetStringValue(HKLM, strBaseKey & strSubKey, "QuietDisplayName", strValue)
    End If

    If (strValue <> "") and (errCheck = 0) Then
	ReDim Preserve arrReturn(i)
	arrReturn(i) = strValue
	i = i + 1
    End If

Next
GetInstalledApps = arrReturn
End Function
'*******************************************************
Private Function Build(arrResults, strRepeatedData)
Dim  Result, i, intFormatCounter, strTMP, strPhysicalDiskTMP, j, strOutput
intFormatCounter = 0
strTMP = strRepeatedData
j = 0
For each result in arrResults
	j = j + 1
	If IsArray(Result) Then
		For i=0 to Ubound(Result,2)

			strPhysicalDiskTMP = Result(0,i) & "(" & Result(1,i) & ")" & strPhysicalDiskTMP
			wscript.echo "Physical Disk: " & Result(0,i)
			wscript.echo "Caption: " & Result(1,i)
			If i < Ubound(Result,2) Then
				strPhysicalDiskTMP = "/" & strPhysicalDiskTMP
			End If

		Next
		strTMP = strTMP & vbtab & strPhysicalDiskTMP
		strPhysicalDiskTMP = ""
	Else
	Select Case intFormatCounter 
	Case 0 
	   wscript.echo "Drive Letter: " & Result
		strTMP = Result
	Case 1 
	   wscript.echo "File System: " & Result
		strTMP = strTMP & vbtab & Result
	Case 2 
	   wscript.echo "Partion Size in MB: " & Result
		strTMP = strTMP & vbtab & Result
	Case 3
	   wscript.echo "Freespace in MB: " & Result
		strTMP = strTMP & vbtab & Result
	Case 4
	   wscript.echo "Percent Freespace: " & Result
		strTMP = strTMP & vbtab & Result
	End Select 
	End If
If intFormatCounter < 5 Then
	intFormatCounter = intFormatCounter +1
Else
	intFormatCounter = 0
	if j < Ubound(arrResults) Then
		strOutput = strOutput & strRepeatedData & strTMP & vbcrlf
	Else
		strOutput = strOutput & strRepeatedData & strTMP
	End If
End If
Next
wscript.echo strOutput
Build = strOutput
End Function

