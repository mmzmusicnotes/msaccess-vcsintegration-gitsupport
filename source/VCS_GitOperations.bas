Attribute VB_Name = "VCS_GitOperations"
Option Compare Database

Option Explicit

' Peforms operations related to interrogating the status of Git
' Note: All of these operations make certain assumptions:
' 1) The database is in the root of the git repository.
' 2) Source code is in the source\ directory.

Private Const GetCommitDateCommand As String = "git show -s --format=%ci HEAD"
Private Const GetImportedFilesCommand As String = " git diff --name-only access-vcs-last-imported-commit..HEAD"

' Return the datestamp of the current head commit
Public Function GetCurrentCommitDate() As Date
    Dim GitDateString As String
    Dim AccessDate As Date
    
    GitDateString = ShellRun(GetCommitDateCommand)
        
    ' convert the result from ISO 8601 to Access,
    ' trimming off the timezone at the end (should always be -0500)
    ' see StackOverflow #38751429
    GitDateString = Split(GitDateString, " -")(0)
    AccessDate = CDate(GitDateString)
    
    GetCurrentCommitDate = AccessDate
End Function

' Returns the result of a shell command as a string
' Commands are always run in the current directory
' Based on StackOverflow #2784367
Public Function ShellRun(sCmd As String) As String
    
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    ' run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec("cmd.exe /c cd " & CurrentProject.Path & " & " & sCmd)

    ' handle the results as they are written to and read from the StdOut object
    ShellRun = oExec.StdOut.ReadAll

End Function

' Returns a list of all the files imported between the current HEAD
' and the commit carrying the last-imported-commit tag that are in the
' /source directory. Note: Last entry in array will be empty.
Public Function GetSourceFilesSinceLastImport() As Collection
    Dim FileListString As String
    Dim AllFilesArray As Variant
    Dim SourceFilesCollection As Collection
    Set SourceFilesCollection = New Collection
    Dim File As Variant
    
    
    FileListString = ShellRun(GetImportedFilesCommand)
    AllFilesArray = Split(FileListString, vbLf)
    
    For Each File In AllFilesArray
        If File <> "" And File Like "source/*" Then
            File = Replace(File, "/", "\")
            SourceFilesCollection.Add File
        End If
    Next
    
    Set GetSourceFilesSinceLastImport = SourceFilesCollection
End Function
