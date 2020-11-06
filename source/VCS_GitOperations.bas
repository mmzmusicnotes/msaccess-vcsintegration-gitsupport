Attribute VB_Name = "VCS_GitOperations"
Option Compare Database

Option Explicit

' Peforms operations related to interrogating the status of Git
' Note: All of these operations make certain assumptions:
' 1) The database is in the root of the git repository.
' 2) Source code is in the source\ directory.

Private Const GetCommitDateCommand As String = "git show -s --format=%ci HEAD"
Private Const GetCommittedFilesCommand As String = "git diff --name-status access-vcs-last-imported-commit..HEAD"
Private Const GetAllChangedFilesCommand As String = "git diff --name-status access-vcs-last-imported-commit"
Private Const GetUntrackedFilesCommand As String = "git ls-files . --exclude-standard --others"
Private Const SetTaggedCommitCommand As String = "git tag access-vcs-last-imported-commit HEAD -f"

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

' Returns a collcetion containing two lists:
' first, of all the objects to modify or re-import based on the state of the git repo
' second, of all the objects to delete based on the same
' if getUncommittedFiles is false, files list is all files between the current HEAD
' and the commit carrying the last-imported-commit tag that are in the
' /source directory. if it is true, file list includes any uncommitted changes
' Note: Last entries in file arrays will be empty.
Public Function GetSourceFilesSinceLastImport(getUncommittedFiles As Boolean) As Variant
    Dim FileListString As String
    Dim AllFilesArray As Variant
    Dim SourceFilesToImportCollection As Collection
    Dim SourceFilesToRemoveCollection As Collection
    Set SourceFilesToImportCollection = New Collection
    Set SourceFilesToRemoveCollection = New Collection
    Dim FileStatus As Variant
    Dim CommandToRun As String
    Dim File As Variant
    Dim Status As String
    Dim FileStatusSplit As Variant
    Dim ReturnArray(2) As Variant

    If getUncommittedFiles = True Then
        CommandToRun = GetAllChangedFilesCommand
    Else
        CommandToRun = GetCommittedFilesCommand
    End If
    
    ' get files already committed (and staged, if flag passed)
    FileListString = ShellRun(CommandToRun)

    ' sanitize paths, determine the operation type, and add to relevant collection
    For Each FileStatus In Split(FileListString, vbLf)
        If FileStatus = "" Then Exit For
        
        FileStatusSplit = Split(FileStatus, vbTab)
        Status = Left(FileStatusSplit(0), 1) ' only first character actually indicates status; the rest is "score"
        File = FileStatusSplit(1)
        
        If File <> "" And File Like "source/*" Then
            File = Replace(File, "/", "\")
            
            ' overwrite/add modified, copied, added
            If Status = "M" Or Status = "A" Or Status = "U" Then
                SourceFilesToImportCollection.Add File
            End If
    
            ' overwrite result of rename or copy
            If Status = "R" Or Status = "C" Then
                ' add the result to the collection of import files
                SourceFilesToImportCollection.Add Replace(FileStatusSplit(2), "/", "\")
            End If
    
            ' remove deleted objects and original renamed files
            If Status = "D" Or Status = "R" Then
                SourceFilesToRemoveCollection.Add File
            End If
        End If
    Next

    ' get and add untracked files
    If getUncommittedFiles = True Then
        FileListString = ShellRun(GetUntrackedFilesCommand)
        For Each File In Split(FileListString, vbLf)
            If File <> "" And File Like "source/*" Then
                File = Replace(File, "/", "\")
                SourceFilesToImportCollection.Add File
            End If
        Next
    End If

    Set ReturnArray(0) = SourceFilesToImportCollection
    Set ReturnArray(1) = SourceFilesToRemoveCollection
    GetSourceFilesSinceLastImport = ReturnArray
End Function

Public Sub SetLastImportedCommitToCurrent()
    ShellRun SetTaggedCommitCommand
End Sub
