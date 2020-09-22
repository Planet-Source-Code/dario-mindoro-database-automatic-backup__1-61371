Attribute VB_Name = "ModMain"
Public Type SHFILEOPSTRUCT
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Const FO_COPY = &H2
Public Const FOF_ALLOWUNDO = &H40
Public Counts As Long


Public Function FileExists(ByVal sPathName As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(sPathName) And vbNormal) = vbNormal
    On Error GoTo 0
End Function

Public Sub CopyFileWindowsWay(SourceFile As String, DestinationFile As String)

     Dim lngReturn As Long
     Dim typFileOperation As SHFILEOPSTRUCT

     With typFileOperation
        .hWnd = 0
        .wFunc = FO_COPY
        .pFrom = SourceFile & vbNullChar & vbNullChar 'source file
        .pTo = DestinationFile & vbNullChar & vbNullChar 'destination file
        .fFlags = FOF_ALLOWUNDO
     End With
     lngReturn = SHFileOperation(typFileOperation)
End Sub







