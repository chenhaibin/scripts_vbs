Option Explicit
Dim strMDBFileName, strFilter, objFSO, chrOverwrite

'*************Configuration Variables
strFilter = "(&(objectCategory=person)(objectClass=user))" 'LDAP Query Filter
strMDBFileName = left(wscript.scriptfullname, InStrRev(wscript.scriptfullname, "\")) & "ADQuery02-18-10.mdb"
'************* End Configuration Variables

'*************DB Prep
Const Jet4x = 5
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strMDBFileName) Then
	wscript.echo "DB File: " & strMDBFileName & " already exists."
	wscript.echo "Delete File? (y/n)"
	chrOverwrite="y"
	Do While Not WScript.StdIn.AtEndOfLine
	chrOverwrite = Wscript.StdIn.Read(1)
	Loop
	if LCase(chrOverwrite) = "y" Then
		objFSO.DeleteFile(strMDBFileName)
		Wscript.Echo "Creating new DB File: " & strMDBFileName
		CreateNewMDB strMDBFileName, Jet4x
		CreateTables strMDBFileName
	Else
		wscript.echo "Exiting Script"
		wscript.quit
	End If
Else
	Wscript.Echo "Creating new DB File: " & strMDBFileName
	CreateNewMDB strMDBFileName, Jet4x
	CreateTables strMDBFileName
End If
'*************End DB Prep

SingleDCQuery strMDBFileName, strFilter
MultiDCQuery strMDBFileName, strFilter
UpdateAccountInfowithMultiDCData strMDBFileName
'***************************dtmInteger8
Private Function dtmInteger8(objToConvert)
Dim objWMIService, colTimeZones, objTimeZone, intTimeZoneBias
Dim intTemp, dtmConverted
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colTimeZones = objWMIService.ExecQuery("Select * From Win32_TimeZone")
For Each objTimeZone in colTimeZones
    intTimeZoneBias = objTimeZone.Bias
Next
If Not IsNull(objToConvert) Then
	If objToConvert.LowPart < 0 then
		intTemp = (objToConvert.HighPart + 1) * (2^32) + objToConvert.LowPart
	Else
		intTemp = objToConvert.HighPart * (2^32) + objToConvert.LowPart
	End If
	intTemp = intTemp / (60 * 10000000)
	intTemp = intTemp / 1440

	dtmConverted = intTemp + #1/1/1601#
	dtmConverted = DateAdd("n", intTimeZoneBias, dtmConverted)
	dtmInteger8 = dtmConverted
Else
dtmInteger8 = "NULL"
End If
End Function
'************************************GetMaxPasswordAge
Private Function GetMaxPWDAge()
Dim objRootDSE, oDOmain
Dim strDNSDomain
Dim maxPwdAge
Dim numDays

Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext") 
Set oDomain = GetObject("LDAP://" & strDNSDomain)
Set maxPwdAge = oDomain.Get("maxPwdAge")
numDays = CCur((maxPwdAge.HighPart * 2 ^ 32) + _
maxPwdAge.LowPart) / CCur(-864000000000)
GetMaxPWDAge = numDays
End Function
'************************************ParseUserAccountControl
Private Function ParseUserAccountControl(intDecimalValueofAccountControl, AttributeRequested)
Dim objHash
Set objHash = CreateObject("Scripting.Dictionary")
objHash.Add "SCRIPT", &h0001
objHash.Add "ACCOUNTDISABLE", &h0002
objHash.Add "HOMEDIR_REQUIRED", &h0008
objHash.Add "LOCKOUT", &h0010
objHash.Add "PASSWD_NOTREQD", &h0020
objHash.Add "PASSWD_CANT_CHANGE", &h0040
objHash.Add "ENCRYPTED_TEXT_PWD_ALLOWED", &h0080
objHash.Add "TEMP_DUPLICATE_ACCOUNT", &h0100
objHash.Add "NORMAL_ACCOUNT", &h0200
objHash.Add "INTERDOMAIN_TRUST_ACCOUNT", &h0800
objHash.Add "WORKSTATION_TRUST_ACCOUNT", &h1000
objHash.Add "SERVER_TRUST_ACCOUNT", &h2000
objHash.Add "DONT_EXPIRE_PASSWORD", &h10000
objHash.Add "MNS_LOGON_ACCOUNT", &h20000
objHash.Add "SMARTCARD_REQUIRED", &h40000
objHash.Add "TRUSTED_FOR_DELEGATION", &h80000
objHash.Add "NOT_DELEGATED", &h100000
objHash.Add "USE_DES_KEY_ONLY", &h200000
objHash.Add "DONT_REQ_PREAUTH", &h400000
objHash.Add "PASSWORD_EXPIRED", &h800000
objHash.Add "TRUSTED_TO_AUTH_FOR_DELEGATION", &h1000000

If objHash(AttributeRequested) And intDecimalValueofAccountControl Then
	ParseUserAccountControl = True
Else
	ParseUserAccountControl = False
End If
End Function
'********************ParsemsDSUserAccountControlComputed
Private Function ParsemsDSUserAccountControlComputed(intValue, AttributeRequested)
Dim objHash
Set objHash = CreateObject("Scripting.Dictionary")

objHash.Add "UF_LOCKOUT", &h0010
objHash.Add "UF_PASSWORD_EXPIRED", &h800000
objHash.Add "UF_PARTIAL_SECRETS_ACCOUNT", &h4000000
objHash.Add "UF_USE_AES_KEYS", &h8000000

If objHash(AttributeRequested) And intValue Then
	ParsemsDSUserAccountControlComputed = True
Else
	ParsemsDSUserAccountControlComputed = False
End If
End Function
'**************************CreateNewMDB
Sub CreateNewMDB(FileName, Format)
  Dim Catalog
  Set Catalog = CreateObject("ADOX.Catalog")
  Catalog.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
     "Jet OLEDB:Engine Type=" & Format & _
    ";Data Source=" & FileName
End Sub
'*************************CreateTables
Sub CreateTables(strMDBFile)
Dim objConnection
Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=" & strMDBFile

objConnection.Execute "CREATE TABLE AccountInfo(" & _
	"ObjectGUID GUID ," & _
	"DistinguishedName VarChar(250) ," & _
	"UserPrincipalName VarChar(100) ," & _
	"LastLogon DATETIME, " & _
	"ModifyTimestamp DATETIME, " & _
	"AccountDisabled YesNo ," & _
	"PWDNoExpire YesNo ," & _
	"PWDNotReq YesNo ," & _
	"PWDLastSet DATETIME ," & _
	"AccountLocked YesNo ," & _	
	"PWDExpired YesNo ," & _
	"PWDExpirationDate DATETIME," & _
	"LastLogonTimeStamp DATETIME" &_
	")"

objConnection.Execute "CREATE TABLE MultiDCAccountInfo(" & _
	"ObjectGUID GUID ," & _
	"LastLogon DATETIME ," & _
	"ModifyTimeStamp DATETIME" & _
	")"

objConnection.Close
End Sub
'*************************InsertData
Sub InsertData(strMDBFile,strTableName,objDictionary)
Dim objConnection, strInsertCommand, colFields, strField
'Define Insert Fields
Set objConnection = CreateObject("ADODB.Connection")

'Build Insert Statement
colFields = objDictionary.keys
strInsertCommand = "Insert Into " & strTableName & "("
For each strField in colFields
	'Build the SQL INSERT command fields based on the dictionary object keys
	strInsertCommand = strInsertCommand & "[" & strField & "],"
Next
'remove the trailing comma from the previous for loop and close the field section
strInsertCommand = left(strInsertCommand, LEN(strInsertCommand) -1) & ")"
'Begin Values section
strInsertCommand = strInsertCommand & " Values ("
For each strField in colFields
strInsertCommand = strInsertCommand & SQLFormat(objDictionary.item(strField)) & ","
Next
'remove the trailing comma from the previous for loop and close the value section
strInsertCommand = left(strInsertCommand, LEN(strInsertCommand) -1) & ")"

objConnection.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=" & strMDBFile
'Debug only
wscript.echo strInsertCommand
objConnection.Execute strInsertCommand
objConnection.Close
End Sub
'************************************GetDomainControllers
Private Function GetDomainControllers()
Dim objRootDSE, strConfigurationNC, objConnection, objCommand, objRecordSet, objParent, strDCList, arrDomainControllers
Const ADS_SCOPE_SUBTREE = 2

Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = "SELECT ADsPath FROM 'LDAP://" & strConfigurationNC & "' WHERE objectClass='nTDSDSA'"  
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    Set objParent = GetObject(GetObject(objRecordset.Fields("ADsPath")).Parent)
    strDCList = strDCList & objParent.DNSHostname & ","
    objRecordSet.MoveNext
Loop
strDCList = Left(strDCList, Len(strDCList) -1)
GetDomainControllers = Split(strDCList, ",")
End Function
'***************************SingleDCQuery
Private Sub SingleDCQuery(strMDBFileName, strFilter)
Dim adoCommand, adoConnection, strBase, strAttributes, intMaxPWDAgeDays, dtmPWDExpires
Dim objRootDSE, strDNSDomain, strQuery, adoRecordset, objRecordData
intMaxPWDAgeDays = GetMaxPWDAge
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

Set objRootDSE = GetObject("LDAP://RootDSE")

strDNSDomain = objRootDSE.Get("defaultNamingContext")
strBase = "<LDAP://" & strDNSDomain & ">"


strAttributes = "userPrincipalName,distinguishedName,objectGUID,pwdLastSet,userAccountControl,msDS-User-Account-Control-Computed,lastLogonTimestamp"

strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

Set adoRecordset = adoCommand.Execute
 
Do Until adoRecordset.EOF

	Set objRecordData = CreateObject("Scripting.Dictionary")
	objRecordData.Add "ObJectGUID", OctetToHexStr(adoRecordset.Fields("objectGUID").Value, 1)
	objRecordData.Add "DistinguishedName", adoRecordset.Fields("distinguishedName").Value
	objRecordData.Add "UserPrincipalName", adoRecordset.Fields("userPrincipalName").value
	objRecordData.Add "AccountDisabled", ParseUserAccountControl(adoRecordset.Fields("userAccountControl").Value,"ACCOUNTDISABLE")
	objRecordData.Add "PWDNoExpire", ParseUserAccountControl(adoRecordset.Fields("userAccountControl").Value,"DONT_EXPIRE_PASSWORD")
	objRecordData.Add "PWDNotReq", ParseUserAccountControl(adoRecordset.Fields("userAccountControl").Value,"PASSWD_NOTREQD")
	objRecordData.Add "PWDLastSet", dtmInteger8(adoRecordset.Fields("pwdLastSet").Value)
	objRecordData.Add "AccountLocked", ParsemsDSUserAccountControlComputed(adoRecordset.Fields("msDS-User-Account-Control-Computed").Value, "UF_LOCKOUT")
	objRecordData.Add "PWDExpired", ParsemsDSUserAccountControlComputed(adoRecordset.Fields("msDS-User-Account-Control-Computed").Value, "UF_PASSWORD_EXPIRED")
	If Not(ParseUserAccountControl(adoRecordset.Fields("userAccountControl").Value,"DONT_EXPIRE_PASSWORD")) Then
		objRecordData.Add "PWDExpirationDate", DateAdd("d", intMaxPWDAgeDays, dtmInteger8(adoRecordset.Fields("pwdLastSet").Value))
	Else
		objRecordData.Add "PWDExpirationDate", "NULL"
	End If
	objRecordData.Add "lastLogonTimestamp", dtmInteger8(adoRecordset.Fields("lastLogonTimestamp").Value)
	InsertData strMDBFileName,"AccountInfo", objRecordData
	objRecordData.Removeall
	adoRecordset.MoveNext
Loop
adoRecordset.Close
adoConnection.Close
End Sub
'***************************SQLFormat
Private Function SQLFormat(objData)
Dim objType
objType = VarType(objData)
Select Case objType

Case 0,1

	'Empty(Uninitialized), or Null
	SQLFormat = "NULL" 

Case 2,3,4,5,6,17

	'Variable is one of the following types:
	'vbInteger,vbLong,vbSingle,vbDouble,vbCurrency,vbByte
	SQLFormat = objData 
Case 11

	'Variable is of boolean type:
	SQLFormat = objData 

Case 7

	'Variable is of vbDate type:
	SQLFormat = "#" & FormatDateTime(objData, 0) & "#"

Case Else

	'String/Catchall
	If Left(objData,1) = "{" AND Right(objData,1) = "}" Then
		'Used as a quick and dirty match for DB style GUID, replace later with regex.
				SQLFormat = objData
	ElseIf LCase(objData)<>"null" Then
		SQLFormat = "'" & Replace(objData,"'","''") & "'"
	Else
		SQLFormat = "NULL"
	End If

End Select 
End Function
'***************************OctetToHexStr
Function OctetToHexStr(arrbytOctet, intReturnFormat)
' Function to convert OctetString (byte array) to Hex string,
' with bytes delimited by \ for an ADO filter.
'0 = LDAP Query Format \FF\FF\FF\FF\FF\FF\FF\FF\FF\FF\FF\FF\FF\FF\FF\FF format (LDAP compatible)
'1 = Access DB Insert Format {FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}

Dim k
OctetToHexStr = ""

If intReturnFormat = 0 Then
For k = 1 To Lenb(arrbytOctet)
	OctetToHexStr = OctetToHexStr & "\" & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
Next
Else
For k = 1 To Lenb(arrbytOctet)
	OctetToHexStr = OctetToHexStr & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
Next
OctetToHexStr = "{" & Left(OctetToHexStr, 8) & "-" & Mid(OctetToHexStr,9,4) & "-" & Mid(OctetToHexStr,13,4) &_
		"-" & Mid(OctetToHexStr,17,4) & "-" & right(OctetToHexStr,12) & "}"
End If
End Function
'***************************MultiDCQuery
Private Sub MultiDCQuery(strMDBFileName, strFilter)
Dim adoCommand, adoConnection, strBase, strAttributes, arrDomainControllers, strDC
Dim objRootDSE, strDNSDomain, strQuery, adoRecordset, objRecordData


Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

Set objRootDSE = GetObject("LDAP://RootDSE")
arrDomainControllers=GetDomainControllers()

For Each strDC in arrDomainControllers

strDNSDomain = objRootDSE.Get("defaultNamingContext")
strBase = "<LDAP://" & strDC & "/" & strDNSDomain & ">"

strAttributes = "objectGUID,lastLogon,modifyTimeStamp"

strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

Set adoRecordset = adoCommand.Execute

Do Until adoRecordset.EOF

	Set objRecordData = CreateObject("Scripting.Dictionary")
	objRecordData.Add "ObJectGUID", OctetToHexStr(adoRecordset.Fields("objectGUID").Value, 1)
	objRecordData.Add "LastLogon", dtmInteger8(adoRecordset.Fields("lastLogon").Value)
	objRecordData.Add "ModifyTimeStamp", adoRecordset.Fields("modifyTimeStamp").Value
	InsertData strMDBFileName,"MultiDCAccountInfo", objRecordData
	objRecordData.Removeall
	adoRecordset.MoveNext
Loop
adoRecordset.Close
Next
adoConnection.Close
End Sub
'***************************UpdateAccountInfowithMultiDCData
Sub UpdateAccountInfowithMultiDCData(strMDBFile)
Dim objConnection
Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open _
    "Provider= Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source=" & strMDBFile

objConnection.Execute _
"SELECT Distinct ObjectGUID, Max(LastLogon) as Last_Logon Into Last_Logon " & _
"FROM MultiDCAccountInfo " & _
"GROUP BY ObjectGUID;"

objConnection.Execute _
"SELECT Distinct ObjectGUID, Max(modifyTimeStamp) as Modify_TimeStamp Into Modify_TimeStamp " & _
"FROM MultiDCAccountInfo " & _
"GROUP BY ObjectGUID;"

objConnection.Execute _
"UPDATE AccountInfo " & _
"INNER JOIN Last_Logon ON AccountInfo.ObjectGUID = Last_Logon.ObjectGUID " & _
"SET AccountInfo.LastLogon = Last_Logon.Last_Logon;"

objConnection.Execute _
"UPDATE AccountInfo " & _
"INNER JOIN Modify_TimeStamp ON AccountInfo.ObjectGUID = Modify_TimeStamp.ObjectGUID " & _
"SET AccountInfo.ModifyTimeStamp = Modify_TimeStamp.Modify_TimeStamp;"

objConnection.Execute "DROP TABLE Last_Logon"
objConnection.Execute "DROP TABLE Modify_Timestamp"
objConnection.Execute "DROP TABLE MultiDCAccountInfo"
objConnection.Close
End Sub