Class MyClass
      Public Foo
  End Class

Dim X,Y,objFile,ceiling,intSemaphore
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

ceiling = 0.75
strComputer = "."
intSemaphore = 1

Call CoreMonitor()
Call MemoryBusy()
Call CoreMonitor()

'1 ciclo = 1.6kb
'100000 ciclos = 160mb


Sub CoreMonitor()
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48)
	'Set objFSO=CreateObject("Scripting.FileSystemObject")
	
		For Each objItem in colItems
			 
			'Wscript.Echo "TotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize
			'Wscript.Echo "FreePhysicalMemory: " & objItem.FreePhysicalMemory
			Wscript.Echo objItem.FreePhysicalMemory 
			Wscript.Echo 1 - (objItem.FreePhysicalMemory / objItem.TotalVisibleMemorySize)
		Next
End Sub

Sub MemoryBusy()

	Do While intSemaphore< 100001
	  Set X = New MyClass
	  Set Y = New MyClass
	  
	  Set X.Foo = Y
	  Set Y.Foo = X
	  
	  Set X = Nothing
	  Set Y = Nothing

	  intSemaphore = intSemaphore + 1
	Loop
	  
 End Sub