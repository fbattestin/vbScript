'https://msdn.microsoft.com/en-us/library/aa394372(v=vs.85).aspx


strComputer = "."
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim objFile

'schedule start
dtmStart = Now
'dtmFinish, variavel que recebe tempo total de execucao. DateAdd soma ao tempo corrente "n" corresponde a minutos.
dtmFinish=DateAdd("n",WScript.arguments(0),dtmStart)
'MsgBox dtmFinish

'intervalo de coleta dos processos.
'por ex.: 	a variavel dtmFinish recebe o valor 5. Esse valor sera somado em minutos ao tempo corrente.()
'			a variavel intCollect recebe o valor 60, logo, durante 5 minutos (dtmFinish) a cada 60 segundos (intCollect) os processos serao coletados

intCollect = WScript.arguments(1) 

Call Scheduler()

Sub Scheduler()
	If Now > dtmFinish Then 
		WScript.Quit()
	End If
	If Now = dtmStart and dtmStart<= dtmFinish Then
		Call Collect()
	Else 
		If Now>=dtmStart and dtmStart<= dtmFinish Then 
			Call Collect()
		End If
	End If

End Sub


Sub Collect()
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)
	Set objFSO=CreateObject("Scripting.FileSystemObject")

	'arquivo de sa[ida]
	outFile="c:\temp\MemorypPocess.inf"

	If (objFSO.FileExists(outFile)) Then
		Set objFile = objFSO.OpenTextFile(outFile, ForAppending, True)

		For Each objItem in colItems
			 
			objFile.Write Now & ";"
			objFile.Write objItem.SessionId & ";"
			objFile.Write objItem.ProcessId & ";"
			objFile.Write objItem.ExecutionState & ";"
			objFile.Write objItem.Name & ";"
			objFile.Write objItem.CSName & ";"
			objFile.Write objItem.ExecutablePath  & ";"
			objFile.Write objItem.PageFaults  & ";"
			objFile.Write objItem.PageFileUsage & ";"
			objFile.Write objItem.PrivatePageCount & ";"
			objFile.Write objItem.VirtualSize & ";"
			objFile.Write objItem.WorkingSetSize & vbCrLf
		Next

	Else
		Set objFile = objFSO.CreateTextFile(outFile,True)
		objFile.Write "Date;SessionId;ProcessId;ExecutionState;ProccessName;CSName;ExecutablePath;PageFaults;PageFileUsage;PrivatePageCount;VirtualSize;WorkingSetSize" & vbCrLf
		For Each objItem in colItems
			
			objFile.Write Now & ";"
			objFile.Write objItem.SessionId & ";"
			objFile.Write objItem.ProcessId & ";"
			objFile.Write objItem.ExecutionState & ";"
			objFile.Write objItem.Name & ";"
			objFile.Write objItem.CSName & ";"
			objFile.Write objItem.ExecutablePath  & ";"
			objFile.Write objItem.PageFaults  & ";"
			objFile.Write objItem.PageFileUsage & ";"
			objFile.Write objItem.PrivatePageCount & ";"
			objFile.Write objItem.VirtualSize & ";"
			objFile.Write objItem.WorkingSetSize & vbCrLf
		Next

	End If

	objFile.Close
	
	WScript.Sleep intCollect * 1000
	Call Scheduler()
	
End Sub

WScript.Quit()
