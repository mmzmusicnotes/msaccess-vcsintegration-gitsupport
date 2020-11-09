Attribute VB_Name = "VCS_IE_Functions"
Option Compare Database

Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Private Const AggressiveSanitize As Boolean = True
Private Const StripPublishOption As Boolean = True

' Constants for Scripting.FileSystemObject API
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8
Public Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2

' Deletes an object of a given type
Public Function DeleteObject(objectName as String, objectType as String) As Variant
    Dim accessObjectType As AcObjectType
    Select Case objectType
        Case "forms"
            accessObjectType = acForm
        Case "reports"
            accessObjectType = acReport
        Case "macros"
            accessObjectType = acMacro
        Case "modules"
            accessObjectType = acModule
        Case "queries"
            accessObjectType = acQuery
        ' todo: are these the correct object types for relations and tbldef?
        Case "relations"
            accessObjectType = acTable
        Case "tables"
            accessObjectType = acTable
        Case "tbldef"
            accessObjectType = acTable
        Case Else
            Exit Function
    End Select
    DoCmd.DeleteObject accessObjectType, objectName
End Function

' Get the correct Modified Date of the passed object.  MSysObjects and DAO are not accurate for all object types.
' See StackOverflow #57103395. This is done because LastUpdated is not reliable.
' Based on a tip from Philipp Stiefel <https://codekabinett.com>
' Getting the last modified date with this line of code does indeed return incorrect results.
'   ? CurrentDb.Containers("Forms").Documents("Form1").LastUpdated
'
' But, that is not what we use to receive the last modified date, except for queries, where the above line is working correctly.
' What we use instead is:
'   ? CurrentProject.AllForms("Form1").DateModified
' LastUpdated is accurate for queries, tables, and relations only. See:
'   https://support.microsoft.com/hr-hr/help/299554 (yes, it's only available in Hungarian)
Public Function GetObjectModifiedDate(ByVal objectName As String, ByVal objectType As String) As Variant
    Select Case objectType
        Case "forms"
            GetObjectModifiedDate = CurrentProject.AllForms(objectName).DateModified
        Case "reports"
            GetObjectModifiedDate = CurrentProject.AllReports(objectName).DateModified
        Case "macros"
            GetObjectModifiedDate = CurrentProject.AllMacros(objectName).DateModified
        Case "modules"
            ' This will report the date that *ANY* module was last saved.
            ' The CurrentDb.Containers method and MSysObjects will report the date created.
            GetObjectModifiedDate = CurrentProject.AllModules(objectName).DateModified
        Case Else
            ' Do nothing.  Return Null.
    End Select
End Function

' Can we export without closing the form?

' Export a database object with optional UCS2-to-UTF-8 conversion.
Public Sub VCS_ExportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                    ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False)

    VCS_Dir.VCS_MkDirIfNotExist Left$(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        Application.SaveAsText obj_type_num, obj_name, tempFileName
        VCS_File.VCS_ConvertUcs2Utf8 tempFileName, file_path
    Else
        Application.SaveAsText obj_type_num, obj_name, file_path
    End If
End Sub

' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub VCS_ImportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                    ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False)
    
    If Not VCS_Dir.VCS_FileExists(file_path) Then Exit Sub
    
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        VCS_File.VCS_ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName
        
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        FSO.DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

'shouldn't this be SanitizeTextFile (Singular)?

' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Public Sub VCS_SanitizeTextFiles(ByVal Path As String, ByVal Ext As String, Optional ByVal CurrentCommitDate As Date)

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    '
    '  Setup Block matching Regex.
    Dim rxBlock As Object
    Set rxBlock = CreateObject("VBScript.RegExp")
    rxBlock.ignoreCase = False
    '
    '  Match PrtDevNames / Mode with or  without W
    Dim srchPattern As String
    srchPattern = "PrtDev(?:Names|Mode)[W]?"
    If (AggressiveSanitize = True) Then
      '  Add and group aggressive matches
      srchPattern = "(?:" & srchPattern
      srchPattern = srchPattern & "|GUID|""GUID""|NameMap|dbLongBinary ""DOL"""
      srchPattern = srchPattern & ")"
    End If
    '  Ensure that this is the beginning of a block.
    srchPattern = srchPattern & " = Begin"
'Debug.Print srchPattern
    rxBlock.Pattern = srchPattern
    '
    '  Setup Line Matching Regex.
    Dim rxLine As Object
    Set rxLine = CreateObject("VBScript.RegExp")
    srchPattern = "^\s*(?:"
    srchPattern = srchPattern & "Checksum ="
    srchPattern = srchPattern & "|BaseInfo|NoSaveCTIWhenDisabled =1"
    If (StripPublishOption = True) Then
        srchPattern = srchPattern & "|dbByte ""PublishToWeb"" =""1"""
        srchPattern = srchPattern & "|PublishOption =1"
    End If
    srchPattern = srchPattern & ")"
'Debug.Print srchPattern
    rxLine.Pattern = srchPattern
    Dim FileName As String
    FileName = Dir$(Path & "*." & Ext)
    Dim isReport As Boolean
    isReport = False
    
    Do Until Len(FileName) = 0
        DoEvents
        
        Dim obj_name As String
        obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
        
        Dim CurrentFile As File
        Set CurrentFile = FSO.GetFile(Path & obj_name & "." & Ext)
        If CurrentCommitDate > CurrentFile.DateLastModified Then GoTo ContinueLenFileName

        Dim InFile As Object
        Set InFile = FSO.OpenTextFile(Path & obj_name & "." & Ext, IOMode:=ForReading, create:=False, Format:=TristateFalse)
        Dim OutFile As Object
        Set OutFile = FSO.CreateTextFile(Path & obj_name & ".sanitize", overwrite:=True, Unicode:=False)
    
        Dim getLine As Boolean
        getLine = True
        
        Do Until InFile.AtEndOfStream
            DoEvents
            Dim txt As String
            '
            ' Check if we need to get a new line of text
            If getLine = True Then
                txt = InFile.ReadLine
            Else
                getLine = True
            End If
            '
            ' Skip lines starting with line pattern
            If rxLine.Test(txt) Then
                Dim rxIndent As Object
                Set rxIndent = CreateObject("VBScript.RegExp")
                rxIndent.Pattern = "^(\s+)\S"
                '
                ' Get indentation level.
                Dim matches As Object
                Set matches = rxIndent.Execute(txt)
                '
                ' Setup pattern to match current indent
                Select Case matches.Count
                    Case 0
                        rxIndent.Pattern = "^" & vbNullString
                    Case Else
                        rxIndent.Pattern = "^" & matches(0).SubMatches(0)
                End Select
                rxIndent.Pattern = rxIndent.Pattern + "\s"
                '
                ' Skip lines with deeper indentation
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If Not rxIndent.Test(txt) Then Exit Do
                Loop
                ' We've moved on at least one line so do get a new one
                ' when starting the loop again.
                getLine = False
            '
            ' skip blocks of code matching block pattern
            ElseIf rxBlock.Test(txt) Then
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            ElseIf InStr(1, txt, "Begin Report") = 1 Then
                isReport = True
                OutFile.WriteLine txt
            ElseIf isReport = True And (InStr(1, txt, "    Right =") Or InStr(1, txt, "    Bottom =")) Then
                'skip line
                If InStr(1, txt, "    Bottom =") Then
                    isReport = False
                End If
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        InFile.Close

        FSO.DeleteFile (Path & FileName)

        Dim thisFile As Object
        Set thisFile = FSO.GetFile(Path & obj_name & ".sanitize")
        
        ' Error Handling to deal with errors caused by Dropbox, VirusScan,
        ' or anything else touching the file.
        Dim ErrCounter As Integer
        On Error GoTo ErrorHandler
        thisFile.Move (Path & FileName)

ContinueLenFileName:
        FileName = Dir$()
    Loop
    
    Exit Sub
ErrorHandler:
    ErrCounter = ErrCounter + 1
    If ErrCounter = 20 Then  ' 20 attempts seems like a nice arbitrary number
        MsgBox "This file could not be moved: " & vbNewLine, vbCritical + vbApplicationModal, _
            "Error moving file..."
        Resume Next
    End If
    Select Case Err.Number
        Case 58    ' "File already exists" error.
            DoEvents
            Sleep 10
            Resume    ' Go back to what you were doing
        Case Else
            MsgBox "This file could not be moved: " & vbNewLine, vbCritical + vbApplicationModal, _
                "Error moving file..."
    End Select
    Resume Next
End Sub

