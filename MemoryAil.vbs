'links uteis: http://www.visualbasicscript.com/VBScript-memory-leaks-m40532.aspx
'https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
'http://www.activexperts.com/admin/scripts/wmi/vbscript/0407/
Class MyClass
      Public Foo
  End Class

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim X,Y,objFile,ceiling,intSemaphore

strComputer = "."
' ceiling threshold for stressing Memory
ceiling = 0.75
intSemaphore = 60 'seconds

Call CoreMonitor()

Sub CoreMonitor()

	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48)
	'Set objFSO=CreateObject("Scripting.FileSystemObject")
	
		For Each objItem in colItems
			 
			'Wscript.Echo objItem.FreePhysicalMemory / objItem.TotalVisibleMemorySize 
			'Wscript.Echo 1 - (objItem.FreePhysicalMemory / objItem.TotalVisibleMemorySize)
			'Wscript.Echo ceiling
			'Wscript.Echo ceiling * objItem.TotalVisibleMemorySize
			
			
			If (ceiling * objItem.TotalVisibleMemorySize) >= objItem.TotalVisibleMemorySize Then
				MsgBox "Total stress ceiling can not be greater (or equal) than the total amount of visible memory to Operation System. Total OS Memory:" & objItem.TotalVisibleMemorySize
				WScript.Quit()
			End If
			
			If (ceiling * objItem.TotalVisibleMemorySize) < objItem.TotalVisibleMemorySize Then ' ceiling smaller total of OS Mem
				If ceiling > (objItem.FreePhysicalMemory / objItem.TotalVisibleMemorySize) Then 'greater than the memory already consumed
					Call MemoryBusy()
				End If
			Else
				Call Semaphore()
			End If
			
		Next
End Sub


Sub MemoryBusy()

	  Set X = New MyClass
	  Set Y = New MyClass
	  
	  Set X.Foo = Y
	  Set Y.Foo = X
	  
	  Set X = Nothing
	  Set Y = Nothing

	  Call Semaphore()
	  
 End Sub
 
 Sub Semaphore()
		WScript.Sleep intSemaphore * 1000
		Call CoreMonitor()
End Sub