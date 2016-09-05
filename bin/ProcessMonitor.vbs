strComputer = "."
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim objFile

'schedule start
dtmStart = Now
dtmFinish=DateAdd("n",5,dtmStart)
MsgBox dtmFinish
'intervarl collect imn seconds
intCollect = 60

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
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process",,48)
	Set objFSO=CreateObject("Scripting.FileSystemObject")

	outFile="c:\temp\process.inf"

	If (objFSO.FileExists(outFile)) Then
		Set objFile = objFSO.OpenTextFile(outFile, ForAppending, True)

		For Each objItem in colItems
			 
			objFile.Write Now & ";"
			objFile.Write objItem.Name & ";"
			objFile.Write objItem.IDProcess & ";"
			objFile.Write objItem.WorkingSet  & ";"
			objFile.Write objItem.PageFaultsPerSec & ";"
			objFile.Write objItem.PageFileBytes & ";"
			objFile.Write objItem.PoolPagedBytes & ";"
			objFile.Write objItem.PrivateBytes & ";"
			objFile.Write objItem.ThreadCount & ";"
			objFile.Write objItem.VirtualBytes & vbCrLf
		Next

	Else
		Set objFile = objFSO.CreateTextFile(outFile,True)
		objFile.Write "Date;IDProccess;ProccessName;WorkingSet;PageFaultsPerSec;PageFileBytes;PoolPagedBytes;PrivateBytes;ThreadCount;VirtualBytes" & vbCrLf
		For Each objItem in colItems
			
			objFile.Write Now & ";"
			objFile.Write objItem.Name & ";"
			objFile.Write objItem.IDProcess & ";"
			objFile.Write objItem.WorkingSet  & ";"
			objFile.Write objItem.PageFaultsPerSec & ";"
			objFile.Write objItem.PageFileBytes & ";"
			objFile.Write objItem.PoolPagedBytes & ";"
			objFile.Write objItem.PrivateBytes & ";"
			objFile.Write objItem.ThreadCount & ";"
			objFile.Write objItem.VirtualBytes & vbCrLf
		Next

	End If

	objFile.Close
	
	WScript.Sleep intCollect * 1000
	Call Scheduler()
	
End Sub

WScript.Quit()
