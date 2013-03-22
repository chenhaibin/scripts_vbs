Option Explicit
Dim strServer, strAppPoolName, strOperation, objWMIService, objIISApplicationPool

strServer = Wscript.Arguments.Named("S")
strAppPoolName = "'W3SVC/AppPools/" & Wscript.Arguments.Named("A")& "'"
strOperation = Wscript.Arguments.Named("O")

If (len(strServer)=0 or len(strAppPoolName)=0 or len(strOperation)=0) Then
wscript.echo "Cscript " & wscript.scriptfullname & " [/S:[TargetServer]] [/A:[ApplicationPool]] [/O:[start|stop|recycle|]]"
Else

Set objWMIService = GetObject("WinMgmts:{authenticationLevel=pktPrivacy}!\\" & strServer & "\root\MicrosoftIISv2") 
Set objIISApplicationPool = objWMIService.Get("IIsApplicationPool.Name=" & strAppPoolName)

Select Case Lcase(strOperation)
Case "stop"
	objIISApplicationPool.stop
Case "start"
	objIISApplicationPool.start
Case "recycle"
	objIISApplicationPool.recycle
Case Else
	wscript.echo "The specified operation is not supported, please specify Start, Stop, or Recycle"
End Select

End If