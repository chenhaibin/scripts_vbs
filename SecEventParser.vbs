Option Explicit
Dim strSummaryDataFilename, strDetailDataFilename, strComputer
Dim Today, BeginTime, EndTime, DayOffset
Dim arrEventIDDesc, arrLogonTypeDesc, EventData, CountData, LogonFailureTable

'User Configured Variables
strComputer = "ts.kaiglan.com"
strSummaryDataFilename = strComputer & " Summary Data.html"
strDetailDataFilename = strComputer & " Detail Data.html"
DayOffset = 1 'Set DayOffset to 0 for Today, 1 for Yesterday etc...
'End User Configured Variables

'Date Range Preparation
	Today = Now
	'Converts a VBScript date object to UTC time with the specified dayoffset
	BeginTime = (Year(Today)*100 + Month(Today) )*100 + (Day(Today) - DayOffset) & "000000.000000-240"
	EndTime = (Year(Today)*100 + Month(Today) )*100 + (Day(Today) - DayOffset) & "235959.999999-240"

'Data Collection
arrEventIDDesc = EventIDDesc()'Multidimensional - Columns:  EventID/EventIDDescription
arrLogonTypeDesc = LogonType()'Multidimensional - Columns:  LogonType/LogonTypeDescription
EventData = GetEvents(strComputer, BeginTime, EndTime, arrEventIDDesc)'Mulidimensional - Columns:See Function
CountData = CountTableData(EventData, arrEventIDDesc, arrLogonTypeDesc)
LogonFailureTable = GenerateLogonFailureTable(EventData)

'Format/Output Data
GenerateSummaryFile strSummaryDataFilename, BeginTime, arrLogonTypeDesc, arrEventIDDesc, CountData, strComputer, LogonFailureTable
GenerateDetailFile strDetailDataFilename, BeginTime, arrLogonTypeDesc, arrEventIDDesc, EventData, strComputer, CountData
'*********************************Get Events**************************************
'1.)  Queries a target computer security event log for events matching those from the function EventIDDesc.
'2.)  Returns a multidimensional array containing the EventID:TimeGenerated:Message data for each sample
Private Function GetEvents(strComputer, BeginTime, EndTime, arrEventIDDesc)
Dim objWMIService
Dim colItems, objItem
Dim WMIQueryEventID
ReDim arrSample(10,0)
Dim i
Dim b


for i=0 to Ubound(arrEventIDDesc, 2) 'Builds EventID list
	If i = 0 Then
		WMIQueryEventID = "EventIdentifier = '" & arrEventIDDesc(0,i) &"' "
	Else
		WMIQueryEventID = WMIQueryEventID & "Or EventIdentifier = '" & arrEventIDDesc(0,i) & "' "
	End If
Next
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("Select EventIdentifier, TimeGenerated, Message, LogFile  From Win32_NTLogEvent where logfile = 'Security' and timegenerated > '"& BeginTime & "' and timegenerated < '" & EndTime & "' and (" & WMIQueryEventID & ")",, 48) 
b = -1
For Each objItem in colItems 
	b = b + 1
	ReDim Preserve arrSample(10,b)
	arrSample(0,b) = objItem.EventIdentifier
    arrSample(1,b) = objItem.TimeGenerated
    arrSample(2,b) = objItem.Message
    arrSample(3,b)= ParseMessage(objItem.Message, "User Name:")
    arrSample(4,b)= ParseMessage(objItem.Message, "Workstation Name:")
    arrSample(5,b)= ParseMessage(objItem.Message, "Domain:")
    arrSample(6,b)= ParseMessage(objItem.Message, "Logon Type:")
    arrSample(7,b)= ParseMessage(objItem.Message, "Logon Process:")
    arrSample(8,b)= ParseMessage(objItem.Message, "Source Network Address:")
    arrSample(9,b)= ParseMessage(objItem.Message, "Source Port:")
    arrSample(10,b)= WMIDateStringToDate(objItem.TimeGenerated)
Next
GetEvents = arrSample
End Function

'******************************EventIDDesc**********************************************
'Returns a Multidimensional array containing:  EventID:Description
Private Function EventIDDesc()
Dim arrDescription (1,14)
'EventID Descriptions
arrDescription(0,0)  =	528 
arrDescription(1,0)  = 	"A user successfully logged on to a computer. (See logon types)"
arrDescription(0,1)  = 	529
arrDescription(1,1)  =	"Logon failure. A logon attempt was made with an unknown user name or a known user name with a bad password."
arrDescription(0,2)  =	530
arrDescription(1,2)  =	"Logon failure.  A logon attempt was made user account tried to log on outside of the allowed time."
arrDescription(0,3)  =	531
arrDescription(1,3)  =	"Logon failure.  A logon attempt was made using a disabled account."
arrDescription(0,4)  =	532
arrDescription(1,4)  =	"Logon failure.  A logon attempt was made using an expired account."
arrDescription(0,5)  =	533
arrDescription(1,5)  =	"Logon failure.  A logon attempt was made by a user who is not allowed to log on at this computer."
arrDescription(0,6)  =	534
arrDescription(1,6)  =	"Logon failure.  The user attempted to log on with a type that is not allowed."
arrDescription(0,7)  =	535
arrDescription(1,7)  =	"Logon failure.  The password for the specified account has expired."
arrDescription(0,8)  =	536
arrDescription(1,8)  =	"Logon failure. The Net Logon service is not active."
arrDescription(0,9)  =	537
arrDescription(1,9)  =	"Logon failure. The logon attempt failed for other reasons."
arrDescription(0,10) =	538
arrDescription(1,10) =	"The logoff process was completed for a user."
arrDescription(0,11) =	539
arrDescription(1,11) =	"Logon failure. The account was locked out at the time the logon attempt was made."
arrDescription(0,12) =	540
arrDescription(1,12) =	"A user successfully logged on to a network."
arrDescription(0,13) =	682
arrDescription(1,13) =	"A user has reconnected to a disconnected terminal server session."
arrDescription(0,14) =	683
arrDescription(1,14) =	"A user has disconnected from a terminal server session without logging off."
EventIDDesc = arrDescription
End Function
'*********************************LogonType*****************************************
'Returns a Multidimensional array containing:  LogonType:Description
Private Function LogonType()
Dim arrLogonType (1, 8)
'Logon Type Descriptions
arrLogonType(0,0) =	2
arrLogonType(1,0) =	"A user logged on to this computer."
arrLogonType(0,1) =	3
arrLogonType(1,1) =	"A user or computer logged on to this computer from the network."
arrLogonType(0,2) =	4
arrLogonType(1,2) =	"Batch logon type is used by batch servers, where processes may be executing on behalf of a user without their direct intervention."
arrLogonType(0,3) =	5
arrLogonType(1,3) =	"A service was started by the Service Control Manager."
arrLogonType(0,4) =	7
arrLogonType(1,4) =	"This workstation was unlocked."
arrLogonType(0,5) =	8
arrLogonType(1,5) =	"A user logged on to this computer from the network. The user's password was passed to the authentication package in its unhashed form. The built-in authentication packages all hash credentials before sending them across the network. The credentials do not traverse the network in plaintext (also called cleartext)."
arrLogonType(0,6) =	9
arrLogonType(1,6) =	"A caller cloned its current token and specified new credentials for outbound connections. The new logon session has the same local identity, but uses different credentials for other network connections."
arrLogonType(0,7) =	10
arrLogonType(1,7) =	"A user logged on to this computer remotely using Terminal Services or Remote Desktop."
arrLogonType(0,8) =	11
arrLogonType(1,8) =	"A user logged on to this computer with network credentials that were stored locally on the computer. The domain controller was not contacted to verify the credentials."
LogonType = arrLogonType
End Function
'************************************PrintMultidimensionalArr******************
'For Testing Only!
Private Sub PrintMultiDimensionalarr(PassedArray)
Dim i
Dim j

For i = 0 to Ubound(PassedArray,2)
	For j = 0 to Ubound(PassedArray,1)
		wscript.echo "[" & PassedArray(j,i) & "]"
	Next
Next
End Sub
'************************************ParseMessage*****************************
Private Function ParseMessage(Message, EventSearch)
Dim objRE, objMatch, colMatches
Dim objRECleanup
Dim filtered
Set objRECleanup = New RegExp
objRECleanup.pattern = "\n|\t|\f|\r|\v|" & EventSearch
objRECleanup.IgnoreCase = True
objRECleanup.Global = True

Set objRE = New RegExp
objRE.Global = True
objRE.ignorecase = true
objRE.Pattern = "\t" & EventSearch & ".{0,}\n"
filtered = ""
Set colMatches = objRE.Execute(Message)
For Each objMatch in colMatches
	filtered = Trim(objRECleanup.Replace(objMatch.Value, ""))
Next
If LCase(filtered)=" " or len(filtered) = 0 Then
	filtered = "-"
End If
ParseMessage = filtered
End Function
'******************************FunctionWMIDateStringToDate*****************************
Function WMIDateStringToDate(dtmInstallDate)
 WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
 Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
 & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
 Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
 13, 2))
End Function
'**************************CountTableData*********************************
Private Function CountTableData(Events, EventDescriptor, LogonTypeDescriptor)
Dim a
Dim b
Dim c
ReDim tmpArray (Ubound(LogonTypeDescriptor,2) + 2 , Ubound(EventDescriptor,2))

For a=0 to Ubound(tmpArray, 2)
tmpArray(0,a) = EventDescriptor(0,a)
Next

for a=0 to Ubound(Events,2)
	for b=0 to Ubound(tmpArray, 2)
		If Events(0,a) = tmpArray(0,b)Then
			tmpArray(1,b) = tmpArray(1,b) + 1
			for c=0 to Ubound(LogonTypeDescriptor,2)
				If LCase(Events(6,a)) = LCase(LogonTypeDescriptor(0,c)) Then
					tmpArray(c+2, b) = tmpArray(c+2, b) + 1
				End If
			Next
		End if
	Next
Next
'Replace Empty Values with -
For a = 0 to Ubound(tmpArray,2)
	For b = 0 to Ubound(tmpArray,1)
		If Len(tmpArray(b,a)) = 0 Then
		tmpArray(b,a) = "-"
		End If 
	Next
Next
CountTableData = tmpArray
End Function
'**************************LogonFailureTableData*********************************
Private Function GenerateLogonFailureTable(Events)
Const TBLOpen = "<TABLE BORDER=1 BorderColor=Gray CellSpacing=0 CellPadding=1 Rules=All align=center>"
Const TBLClose = "</Table>"
Const TBLHeader = "<tr><th colspan=""2"">Logon Failure Summary</th></tr><tr><th colspan=""1"">Username</th><th colspan =""1"">Event 529 Count</th></tr>"

Dim i
Dim j
Dim KeysArray, strGeneratedTable
Dim objFailedLoginData
Set objFailedLoginData = CreateObject("Scripting.Dictionary")
For i = 0 to Ubound(Events,2)

If Events(0,i) = "529" Then
	If objFailedLoginData.Exists(Lcase(Events(3,i))) then
		objFailedLoginData.item(Lcase(Events(3,i))) = objFailedLoginData.item(Lcase(Events(3,i))) + 1
	Else
		objFailedLoginData.Add Lcase(Events(3,i)), 1
	End If
End If

Next

KeysArray = objFailedLoginData.Keys
strGeneratedTable = TBLOpen & TBLHeader
for i=0 to Ubound(KeysArray)
strGeneratedTable = strGeneratedTable & "<tr><td align=left>" & KeysArray(i) & "</td>" & vbtab & "<td align=center>" & objFailedLoginData.Item(KeysArray(i)) & "</td></tr>"
Next
strGeneratedTable = strGeneratedTable & TBLClose
GenerateLogonFailureTable = strGeneratedTable
End Function
'***********************ChooseDescriptor**********************************
Private Function ChooseDescriptor(strToCompare, arrCompareAgainst)
Dim i
For i = 0 to Ubound(arrCompareAgainst, 2)
If LCase(strToCompare) = LCase(arrCompareAgainst(0,i)) Then
ChooseDescriptor = arrCompareAgainst(1,i)
End If
Next
End Function
'***********************GenerateSummaryFile**********************************
Private Sub GenerateSummaryFile(FileName, BeginTime, LogonTypeDescriptor, EventDescriptor, CountData, TargetComputer, LogonFailureTable)
Dim objFSO, objFile
Dim j
Dim i
Dim HTMLOpen
Dim HTMLClose
Dim TBLOpen
Dim TBLClose
Dim tmpLine
HTMLOpen = "<html>" & "<body><h3 align=center>" & TargetComputer & " Security Event Log Review for " & WMIDateStringToDate(BeginTime) & "</h3>"
HTMLClose = "</body></html>"
TBLOpen = "<TABLE BORDER=1 BorderColor=Gray CellSpacing=0 CellPadding=1 Rules=All align=center>"
TBLClose = "</Table>"

Const ForWriting = 2
Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(FileName)
objFile.Close
Set objFile = objFSO.OpenTextFile(FileName, ForWriting)

objFile.WriteLine HTMLOpen

objFile.WriteLine TBLOpen
objFile.WriteLine "<tr><th colspan=" & (ubound(LogonTypeDescriptor, 2) + 3) & ">Security Event Count Table</th></tr>" & vbCrLf
objFile.WriteLine "<tr><th colspan=""2"">EventID Count</th><th colspan =" & (ubound(LogonTypeDescriptor, 2) + 1) & ">Logon Type Count</th></tr>" & vbCrLf
tmpLine = "<tr><th aligh=center>EventID</th><th align=center>Total Event Occurrences</th>"
for i=0 to Ubound(LogonTypeDescriptor,2)
tmpLine = tmpLine & "<th align=center title=""" & LogonTypeDescriptor(1,i) & """>" & LogonTypeDescriptor(0,i) & "</th>"
Next
tmpLine = tmpLine & "</tr>" & vbCrLf
objFile.Writeline tmpLine
tmpLine = ""

For i=0 to Ubound(CountData, 2)
tmpLine = tmpLine & "<tr><th align=center title=""" & ChooseDescriptor(CountData(0,i), EventDescriptor) & """>" & CountData(0,i) & "</th>"
tmpLine = tmpLine & "<td align=center>" & CountData(1,i) & "</td>"
	For j=2 to Ubound(CountData, 1)
	tmpLine = tmpLine & "<td align=center width=35>" & CountData(j,i) & "</td>" & vbtab 
	Next
tmpline = tmpLine & "</tr>" &  vbCrLf
objFile.Writeline tmpLine
tmpLine = ""
Next
objFile.Writeline TBLClose
objFile.Writeline "<p></p>"
objFile.Writeline LogonFailureTable
objFile.Writeline "<p></p>"
objFile.WriteLine TBLOpen
objFile.WriteLine "<tr><th colspan=2>EventID Description Table</th></tr>" & vbCrLf
objFile.writeline "<tr><th>EventID</th><th>EventID Description</th></tr>" &  vbCrLf
For i=0 to Ubound(EventDescriptor, 2)
tmpLine = tmpLine & "<tr><td align=center title=""" & EventDescriptor(1,i) & """>" & EventDescriptor(0,i) & "</td>"
tmpLine = tmpLine & "<td align=center>" & Eventdescriptor(1,i) & "</td>" & vbtab
tmpline = tmpLine & "</tr>" &  vbCrLf
objFile.Writeline tmpLine
tmpLine = ""
Next

objFile.Writeline TBLClose
objFile.Writeline "<p></p>"

objFile.WriteLine TBLOpen
objFile.WriteLine "<tr><th colspan=2>Logon Type Description Table</th></tr>" & vbCrLf
objFile.writeline "<tr><th>Logon Type</th><th>Logon Type Description</th></tr>" &  vbCrLf
For i=0 to Ubound(LogonTypeDescriptor, 2)
tmpLine = tmpLine & "<tr><td align=center title=""" & LogonTypeDescriptor(1,i) & """>" & LogonTypeDescriptor(0,i) & "</td>"
tmpLine = tmpLine & "<td align=center>" & LogonTypedescriptor(1,i) & "</td>" & vbtab
tmpline = tmpLine & "</tr>" &  vbCrLf
objFile.Writeline tmpLine
tmpLine = ""
Next

objFile.Writeline TBLClose
objFile.WriteLine HTMLClose
objFile.Close	
End Sub
'***********************GenerateDetailFile**********************************
Private Sub GenerateDetailFile(FileName, BeginTime, LogonTypeDescriptor, EventDescriptor, Events, TargetComputer, CountData)
Dim objFSO, objFile
Dim j
Dim i
Dim HTMLOpen
Dim HTMLClose
Dim TBLOpen
Dim TBLClose
Dim tmpLine
HTMLOpen = "<html>" & "<body><h3 align=center>" & TargetComputer & " Security Event Log Review for " & WMIDateStringToDate(BeginTime) & "</h3>"
HTMLClose = "</body></html>"
TBLOpen = "<TABLE BORDER=1 BorderColor=Gray CellSpacing=0 CellPadding=1 Rules=All align=center>"
TBLClose = "</Table>"

Const ForWriting = 2
Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(FileName)
objFile.Close
Set objFile = objFSO.OpenTextFile(FileName, ForWriting)

objFile.WriteLine HTMLOpen

objFile.WriteLine TBLOpen
objFile.WriteLine "<tr><th colspan=" & (ubound(LogonTypeDescriptor, 2) + 3) & ">Security Event Count Table</th></tr>" & vbCrLf
objFile.WriteLine "<tr><th colspan=""2"">EventID Count</th><th colspan =" & (ubound(LogonTypeDescriptor, 2) + 1) & ">Logon Type Count</th></tr>" & vbCrLf
tmpLine = "<tr><th aligh=center>EventID</th><th align=center>Total Event Occurrences</th>"
for i=0 to Ubound(LogonTypeDescriptor,2)
tmpLine = tmpLine & "<th align=center title=""" & LogonTypeDescriptor(1,i) & """>" & LogonTypeDescriptor(0,i) & "</th>"
Next
tmpLine = tmpLine & "</tr>" & vbCrLf
objFile.Writeline tmpLine
tmpLine = ""

For i=0 to Ubound(CountData, 2)
tmpLine = tmpLine & "<tr><th align=center title=""" & ChooseDescriptor(CountData(0,i), EventDescriptor) & """>" & CountData(0,i) & "</th>"
tmpLine = tmpLine & "<td align=center>" & CountData(1,i) & "</td>"
	For j=2 to Ubound(CountData, 1)
	tmpLine = tmpLine & "<td align=center width=35>" & CountData(j,i) & "</td>" & vbtab 
	Next
tmpline = tmpLine & "</tr>" &  vbCrLf
objFile.Writeline tmpLine
tmpLine = ""
Next
objFile.Writeline TBLClose
objFile.Writeline "<p></p>"
objFile.WriteLine TBLOpen
objFile.WriteLine "<tr><th colspan=2>EventID Description Table</th></tr>" & vbCrLf
objFile.writeline "<tr><th>EventID</th><th>EventID Description</th></tr>" &  vbCrLf
For i=0 to Ubound(EventDescriptor, 2)
tmpLine = tmpLine & "<tr><td align=center title=""" & EventDescriptor(1,i) & """>" & EventDescriptor(0,i) & "</td>"
tmpLine = tmpLine & "<td align=center>" & Eventdescriptor(1,i) & "</td>" & vbtab
tmpline = tmpLine & "</tr>" &  vbCrLf
objFile.Writeline tmpLine
tmpLine = ""
Next

objFile.Writeline TBLClose
objFile.Writeline "<p></p>"
objFile.WriteLine TBLOpen
objFile.WriteLine "<tr><th colspan=2>Logon Type Description Table</th></tr>" & vbCrLf
objFile.writeline "<tr><th>Logon Type</th><th>Logon Type Description</th></tr>" &  vbCrLf
For i=0 to Ubound(LogonTypeDescriptor, 2)
tmpLine = tmpLine & "<tr><td align=center title=""" & LogonTypeDescriptor(1,i) & """>" & LogonTypeDescriptor(0,i) & "</td>"
tmpLine = tmpLine & "<td align=center>" & LogonTypedescriptor(1,i) & "</td>" & vbtab
tmpline = tmpLine & "</tr>" &  vbCrLf
objFile.Writeline tmpLine
tmpLine = ""
Next

objFile.Writeline TBLClose

objFile.Writeline "<p></p>"
objFile.WriteLine TBLOpen
objFile.WriteLine "<tr><th colspan=9>Logon / Logoff Events Detail Table</th></tr>" & vbCrLf
objFile.writeline "<tr><th>EventID</th><th>Time Generated</th><th>Logon Type</th><th>User Name</th><th>Workstation Name</th><th>Domain</th><th>Logon Process</th><th>Source Network Address</th><th>Source Port</th></tr>" &  vbCrLf
For i=0 to Ubound(Events, 2)
'EventID,Time Generated,Logon Type,username,workstation name,Domain,LogonProcess,Source Network Address,Source Network Port
tmpLine = tmpLine & "<tr><td align=center title=""" & ChooseDescriptor(Events(0,i), EventDescriptor) & """>" & Events(0,i) & "</td>"
tmpLine = tmpLine & "<td align=center>" & Events(10,i) & "</td>" & vbtab
tmpLine = tmpLine & "<td align=center title=""" & ChooseDescriptor(Events(6,i), LogonTypeDescriptor) & """>" & Events(6,i) & "</td>"
tmpLine = tmpLine & "<td align=center>" & Events(3,i) & "</td>" & vbtab
tmpLine = tmpLine & "<td align=center>" & Events(4,i) & "</td>" & vbtab
tmpLine = tmpLine & "<td align=center>" & Events(5,i) & "</td>" & vbtab
tmpLine = tmpLine & "<td align=center>" & Events(7,i) & "</td>" & vbtab
tmpLine = tmpLine & "<td align=center>" & Events(8,i) & "</td>" & vbtab
tmpLine = tmpLine & "<td align=center>" & Events(9,i) & "</td>" & vbtab
tmpline = tmpLine & "</tr>" &  vbCrLf
objFile.Writeline tmpLine
tmpLine = ""
Next

objFile.Writeline TBLClose
objFile.WriteLine HTMLClose
objFile.Close	
End Sub