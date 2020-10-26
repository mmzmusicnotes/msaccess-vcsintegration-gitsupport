Attribute VB_Name = "VCS_Loader"
Option Compare Database

Option Explicit

Public Sub loadVCS(Optional ByVal SourceDirectory As String)
    If SourceDirectory = vbNullString Then
      SourceDirectory = CurrentProject.Path & "\MSAccess-VCS\source\"
    End If

'check if directory exists! - SourceDirectory could be a file or not exist
On Error GoTo Err_DirCheck
    If ((GetAttr(SourceDirectory) And vbDirectory) = vbDirectory) Then
        GoTo Fin_DirCheck
    Else
        'SourceDirectory is not a directory
        Err.Raise 60000, "loadVCS", "Source Directory specified is not a directory"
    End If

Err_DirCheck:
    
    If Err.Number = 53 Then 'SourceDirectory does not exist
        Debug.Print "Error: " & Err.Number & " | " & "File/Directory not found"
    Else
        Debug.Print "Error: " & Err.Number & " | " & Err.Description
    End If
    Exit Sub
Fin_DirCheck:

    'delete if modules already exist + provide warning of deletion?

    On Error GoTo Err_DelHandler

    Dim fileName As String
    'Use the list of files to import as the list to delete
    fileName = Dir$(SourceDirectory & "*.bas")
    Do Until Len(fileName) = 0
        'strip file type from file name
        fileName = Left$(fileName, InStrRev(fileName, ".bas") - 1)
        DoCmd.DeleteObject acModule, fileName
        fileName = Dir$()
    Loop

    GoTo Fin_DelHandler
    
Err_DelHandler:
    If Err.Number <> 7874 Then 'is not - can't find object
        Debug.Print "WARNING (" & Err.Number & ") | " & Err.Description
    End If
    Resume Next
    
Fin_DelHandler:
    fileName = vbNullString

'import files from specific dir? or allow user to input their own dir?
On Error GoTo Err_LoadHandler

    fileName = Dir$(SourceDirectory & "*.bas")
    Do Until Len(fileName) = 0
        'strip file type from file name
        fileName = Left$(fileName, InStrRev(fileName, ".bas") - 1)
        Application.LoadFromText acModule, fileName, SourceDirectory & fileName & ".bas"
        fileName = Dir$()
    Loop

    GoTo Fin_LoadHandler
    
Err_LoadHandler:
    Debug.Print "Error: " & Err.Number & " | " & Err.Description
    Resume Next

Fin_LoadHandler:
    displayFormVersion SourceDirectory

End Sub

Public Sub displayFormVersion(ByVal SourceDirectory As String)
On Error GoTo Err_FormVersion
    Dim versionPath As String, FormsVersion As String, textline As String, posLat As Integer, posLong As Integer
    versionPath = SourceDirectory & "\VERSION.txt"
    Open versionPath For Input As #1

    Do Until EOF(1)
        Line Input #1, textline
        FormsVersion = FormsVersion & textline
        
    Loop
    Close #1

    MsgBox "Form Version: " & FormsVersion & " loaded"

    GoTo Fin_FormVersion
    
Err_FormVersion:

    If Err.Number = 53 Then 'VERSION.txt does not exist
        Debug.Print "Error: " & Err.Number & " | " & "Path to VERSION.txt not found"
    Else
        Debug.Print "Error: " & Err.Number & " | " & Err.Description
    End If
    Exit Sub

Fin_FormVersion:
    Debug.Print "Done"

End Sub
