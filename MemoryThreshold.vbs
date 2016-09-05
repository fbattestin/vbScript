strComputer = "."
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim objFile

Call CoreMonitor()

Sub CoreMonitor()
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48)
	'Set objFSO=CreateObject("Scripting.FileSystemObject")
	
		For Each objItem in colItems
			 
			'Wscript.Echo "TotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize
			'Wscript.Echo "FreePhysicalMemory: " & objItem.FreePhysicalMemory
			Wscript.Echo objItem.FreePhysicalMemory / objItem.TotalVisibleMemorySize 
			Wscript.Echo 1 - (objItem.FreePhysicalMemory / objItem.TotalVisibleMemorySize)
		Next
End Sub