Option Explicit
Dim ActiveNode
Dim Today

Const ClusterManagementIP = "192.168.17.1"
Const PrimaryNode = "Webnode1" 'Ensure that the server hosting the script can resolve the nodes via DNS.
Const SecondaryNode = "Webnode2"
Const SourceBackupPath = "C:\Metabase Backup\" 'Path to location to backup Metabase - Must Be A Local Directory!(IIScnf.vbs limitation)
Const UNCSourceBackupPath = "\C$\Metabase Backup\"'Path Used for backup consolidation, leave \\Node name out.
Const MetabaseImportPath = "\C$\Metabase Import\" 'Path used to search for the most recent Active metabase Backup.  Used to import to passive.  Leave \\Node name out.
Const MetabaseImportPathLocal = "C:\Metabase Import\"

'Discover which node is Active (1.)Works only from a remote server 2.)NLB must be configured in Single Host mode for Port rules)
Today = Now
ActiveNode = GetActiveNode(ClusterManagementIP)

'Backup The Metabase of Primary Node
BackupMetabase ActiveNode, PrimaryNode, SourceBackupPath, Today
'Backup The Metabase of Secondary node
BackupMetabase ActiveNode, SecondaryNode, SourceBackupPath, Today

ConsolidateBackupFiles PrimaryNode, SourceBackupPath,UNCSourceBackupPath
ConsolidateBackupFiles SecondaryNode, SourceBackupPath, UNCSourceBackupPath
If Lcase(ActiveNode) <> Lcase(PrimaryNode) Then
	ImportMetabase ActiveNode, PrimaryNode, MetabaseImportPath, SourceBackupPath, MetabaseImportPathLocal, Today
ElseIf Lcase(ActiveNode) <> Lcase(SecondaryNode) Then
	ImportMetabase ActiveNode, SecondaryNode, MetabaseImportPath, SourceBackupPath, MetabaseImportPathLocal, Today
End If

'**********************************GetActiveNode*******************************************
Private Function GetActiveNode(ClusterManagementIP)
Dim objWMIService, colItems, objItem

Set objWMIService = GetObject("winmgmts:\\" & ClusterManagementIP & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT CSName FROM Win32_OperatingSystem",,48) 
For Each objItem in colItems 
    GetActiveNode = objItem.CSName
Next
End Function
'*********************************BackupMetabase(ActiveNode, PrimaryNode, SecondaryNode)****
Private Sub BackupMetabase(ActiveNode, NodeToBackup, SourceBackupPath, Today)
Dim strFileName
Dim objShell
Dim objExec
Dim MetabaseBackupCommand
Dim objWMIService
Dim intProcessID
Dim errReturn
Dim stdErrLogFile
Dim stdOutLogFile
Dim ProcessRunning

'Construct the Backup FileName
Today = Now
If Lcase(ActiveNode) = Lcase(NodeToBackup) Then
strFileName = Month(Today)& "-" & Day(Today) & "-" & Year(today) & " IIS Metabase Backup(Active - " & NodeToBackup & ")"& ".xml"
stdErrLogFile = Month(Today)& "-" & Day(Today) & "-" & Year(today) & " IIS Metabase Backup(Active - " & NodeToBackup & ")"& "stdERR.log"
stdOutLogFile = Month(Today)& "-" & Day(Today) & "-" & Year(today) & " IIS Metabase Backup(Active - " & NodeToBackup & ")"& "stdOUT.log"
Elseif Lcase(ActiveNode) <> Lcase(NodeToBackup) Then
strFileName = Month(Today)& "-" & Day(Today) & "-" & Year(today) & " IIS Metabase Backup(Passive - " & NodeToBackup & ")"& ".xml"
stdErrLogFile = Month(Today)& "-" & Day(Today) & "-" & Year(today) & " IIS Metabase Backup(Passive - " & NodeToBackup & ")"& "stdERR.log"
stdOutLogFile = Month(Today)& "-" & Day(Today) & "-" & Year(today) & " IIS Metabase Backup(Passive - " & NodeToBackup & ")"& "stdOUT.log"
End If
MetabaseBackupCommand = "cmd.exe /c cscript C:\Windows\System32\iiscnfg.vbs /export /f """ & SourceBackupPath & strFileName & """ /sp / /children > """ & SourceBackupPath & stdOutLogFile & """ 2> """ & SourceBackupPath & stdErrLogFile & """" & "& net stop ""IIS Admin Service"" /Y & net start ""World Wide Web Publishing Service"" & net start ""Simple Mail Transfer Protocol (SMTP)"""

'Run IIScnfg.vbs
Set objWMIService = GetObject _
    ("winmgmts:\\" & NodeToBackup & "\root\cimv2:Win32_Process")
errReturn = objWMIService.Create _
    (MetabaseBackupCommand, Null, Null, intProcessID)
'Wscript.Echo "(If 0, then no error occured creating process): " & errReturn
ProcessRunning = ProcessMonitor(NodeToBackup, intProcessID)
Do While ProcessRunning = 0
'wscript.echo "ProcessID: " & intProcessID & " Active"
ProcessRunning = ProcessMonitor(NodeToBackup, intProcessID)
wscript.sleep(1000)
Loop
End Sub
'*************************************ConsolidateBackupFiles**********************
Private Sub ConsolidateBackupFiles(Node, SourceBackupPath, UNCSourceBackupPath)
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile "\\" & Node & UNCSourceBackupPath & "*", SourceBackupPath, true
objFSO.DeleteFile "\\" & Node & UNCSourceBackupPath & "*", true
End Sub
'*********************************ProcessMonitor**********************************
Private Function ProcessMonitor(Node, PID)
Dim objWMIService
Dim colItems
Dim objItem
Dim PIDPresent
PIDPresent = 1
Set objWMIService = GetObject("winmgmts:\\" & Node & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT ProcessId FROM Win32_Process",,48) 
For Each objItem in colItems 
	If PID = objItem.ProcessId Then
	PIDPresent = 0
	End If
Next
ProcessMonitor = PIDPresent
End Function
'****************************************ImportMetabase******************************
Private Sub ImportMetabase(ActiveNode, ImportTargetNode, MetabaseImportPath, SourceBackupPath, MetabaseImportPathLocal, Today)
Dim objFSO
Dim objWMIService
Dim errReturn
Dim ProcessRunning
Dim strFileName
Dim MetabaseImportCommand
Dim intProcessID

strFileName = Month(Today)& "-" & Day(Today) & "-" & Year(today) & " IIS Metabase Backup(Active - " & ActiveNode & ").xml"
MetabaseImportCommand = "cmd.exe /c cscript C:\Windows\System32\iiscnfg.vbs /import /f """ & MetabaseImportPathLocal & strFileName & """ /sp / /dp / /children > """ & "C:\Metabase Import\LastImport-stdOut.log" & """ 2> """ & "C:\Metabase Import\LastImport-stdErr.log" & """"& "& net stop ""IIS Admin Service"" /Y & net start ""World Wide Web Publishing Service"" & net start ""Simple Mail Transfer Protocol (SMTP)"""

Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile SourceBackupPath & strFileName, "\\" & ImportTargetNode & MetabaseImportPath, true
'wscript.echo MetabaseImportCommand
Set objWMIService = GetObject _
   ("winmgmts:\\" & ImportTargetNode & "\root\cimv2:Win32_Process")
errReturn = objWMIService.Create _
    (MetabaseImportCommand, Null, Null, intProcessID)
'Wscript.Echo "(If 0, then no error occured creating process): " & errReturn
ProcessRunning = ProcessMonitor(ImportTargetNode, intProcessID)
Do While ProcessRunning = 0
'wscript.echo "ProcessID: " & intProcessID & " Active"
ProcessRunning = ProcessMonitor(ImportTargetNode, intProcessID)
wscript.sleep(1000)
Loop
End Sub