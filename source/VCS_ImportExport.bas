Option Compare Database

Option Explicit
' List of lookup tables that are part of the program rather than the
' data, to be exported with source code
' Set to "*" to export the contents of all tables
'Only used in ExportAllSource
Private Const INCLUDE_TABLES As String = ""

Private Const INCLUDE_TABLES_PFS As String = ""
' This is used in ImportAllSource
Private Const DebugOutput As Boolean = False
'this is used in ExportAllSource
'Causes the VCS_ code to be exported
Private Const ArchiveMyself As Boolean = False

'export/import all Queries as plain SQL text
Private Const HandleQueriesAsSQL As Boolean = True

'returns true if named module is NOT part of the VCS code
Private Function IsNotVCS(ByVal moduleName As String) As Boolean
    If moduleName <> "VCS_ImportExport" And _
      moduleName <> "VCS_IE_Functions" And _
      moduleName <> "VCS_File" And _
      moduleName <> "VCS_Dir" And _
      moduleName <> "VCS_String" And _
      moduleName <> "VCS_Loader" And _
      moduleName <> "VCS_Table" And _
      moduleName <> "VCS_Reference" And _
      moduleName <> "VCS_DataMacro" And _
      moduleName <> "VCS_Report" And _
      moduleName <> "VCS_Relation" And _
      moduleName <> "VCS_Query" And _
      moduleName <> "VCS_Button_Functions" And _
      moduleName <> "VCS_GitOperations" Then
        IsNotVCS = True
    Else
        IsNotVCS = False
    End If

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
Public Function GetObjectModifiedDate(objectName As String, objectType As String) As Variant
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


' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
Public Sub ExportSource(ByVal ExportReports As Boolean, ByVal ExportQueries As Boolean, ByVal ExportForms As Boolean, ByVal ExportMacros As Boolean, _
        ByVal ExportModules As Boolean, ByVal ExportTables As Boolean, ByVal ExportReferences As Boolean, ByVal isButton As Boolean, ByVal editedOnly As Boolean)
    Dim Db As Object ' DAO.Database
    Dim source_path As String
    Dim source_path_pfs As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim obj_data_count As Integer
    Dim ucs2 As Boolean
    Dim ExportTablesTemp As Boolean
    Dim CurrentCommitDate As Date
    
    Set Db = CurrentDb
    
    If isButton = True Then
        ExportTablesTemp = False
    Else
        ExportTablesTemp = ExportTables
    End If
    
    If editedOnly = True Then CurrentCommitDate = VCS_GitOperations.GetCurrentCommitDate()

    CloseFormsReports
    'InitVCS_UsingUcs2

    source_path = VCS_Dir.VCS_ProjectPath() & "source\"
    source_path_pfs = VCS_Dir.VCS_ProjectPath() & "pfs\"
    VCS_Dir.VCS_MkDirIfNotExist source_path

    Debug.Print

        If ExportQueries Then
                obj_path = source_path & "queries\"
                If editedOnly = False Then VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "bas"
                Debug.Print VCS_String.VCS_PadRight("Exporting queries...", 24);
                obj_count = 0
                For Each qry In Db.QueryDefs
                    DoEvents
                    Dim skipQueryNotEditedAfterCommit As Boolean
                    ' note that every module will always be exported, due to a quirk in
                    ' how modules' DateModified flag is set - but this is OK,
                    ' as modules are generally the fastest to export
                    skipQueryNotEditedAfterCommit = editedOnly And _
                        CurrentCommitDate > qry.LastUpdated ' accurate for queries
                    If skipQueryNotEditedAfterCommit = False And Left$(qry.name, 1) <> "~" Then
                        ' Replace characters unacceptable to Windows
                        ' note: this should not change the query's actual name;
                        ' that would require DoCmd or SQL
                        qry.name = Replace(qry.name, "\", "%backslash%")
                        qry.name = Replace(qry.name, "/", "%forwardslash%")
                        qry.name = Replace(qry.name, ":", "%colon%")
                        qry.name = Replace(qry.name, "*", "%asterisk%")
                        qry.name = Replace(qry.name, "?", "%questionmark%")
                        qry.name = Replace(qry.name, """", "%doublequotes%")
                        qry.name = Replace(qry.name, "<", "%leftarrow%")
                        qry.name = Replace(qry.name, ">", "%rightarrow%")
                        qry.name = Replace(qry.name, "|", "%pipe%")
                        If editedOnly = True Then VCS_Dir.VCS_DelIfExist (obj_path)
                        If HandleQueriesAsSQL Then
                            VCS_Query.ExportQueryAsSQL qry, obj_path & qry.name & ".bas", False
                        Else
                            VCS_IE_Functions.VCS_ExportObject acQuery, qry.name, obj_path & qry.name & ".bas", VCS_File.VCS_UsingUcs2
                        End If
                        obj_count = obj_count + 1
                    End If
                Next
                Debug.Print VCS_String.VCS_PadRight("Sanitizing...", 15);
                VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas", CurrentCommitDate
                Debug.Print "[" & obj_count & "]"
        End If

    
    For Each obj_type In Split( _
        "forms|Forms|" & acForm & "," & _
        "reports|Reports|" & acReport & "," & _
        "macros|Scripts|" & acMacro & "," & _
        "modules|Modules|" & acModule _
        , "," _
    )
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        obj_type_name = obj_type_split(1)
        obj_type_num = Val(obj_type_split(2))
        obj_path = source_path & obj_type_label & "\"
        obj_count = 0
                
                If (obj_type_label = "forms" And ExportForms) _
            Or (obj_type_label = "reports" And ExportReports) _
            Or (obj_type_label = "macros" And ExportMacros) _
            Or (obj_type_label = "modules" And ExportModules) Then
                        If editedOnly = False Then VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "bas"
                        Debug.Print VCS_String.VCS_PadRight("Exporting " & obj_type_label & "...", 24);
                        For Each doc In Db.Containers(obj_type_name).Documents
                                DoEvents
                                Dim skipObjectNotEditedAfterCommit As Boolean
                                ' note that every module will always be exported, due to a quirk in
                                ' how modules' DateModified flag is set - but this is OK,
                                ' as modules are generally the fastest to export
                                skipObjectNotEditedAfterCommit = editedOnly And _
                                    CurrentCommitDate > GetObjectModifiedDate(doc.name, obj_type_label)
                                If (Left$(doc.name, 1) <> "~") And _
                                   (IsNotVCS(doc.name) Or ArchiveMyself) And _
                                   skipObjectNotEditedAfterCommit = False Then
                                        If obj_type_label = "modules" Then
                                                ucs2 = False
                                        Else
                                                ucs2 = VCS_File.VCS_UsingUcs2
                                        End If
                                        Dim obj_full_path As String
                                        ' todo: should use .form for forms?
                                        obj_full_path = obj_path & doc.name & ".bas"
                                        If editedOnly = True Then VCS_Dir.VCS_DelIfExist (obj_full_path)
                                        VCS_IE_Functions.VCS_ExportObject obj_type_num, doc.name, obj_full_path, ucs2
                                        
                                        If obj_type_label = "reports" Then
                                                VCS_Report.VCS_ExportPrintVars doc.name, obj_path & doc.name & ".pv"
                                        End If
                                        
                                        obj_count = obj_count + 1
                                End If
                        Next

                        Debug.Print VCS_String.VCS_PadRight("Sanitizing...", 15);
                        If obj_type_label <> "modules" Then
                                VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas", CurrentCommitDate
                        End If
                        Debug.Print "[" & obj_count & "]"
                End If
    Next
    
    If ExportReferences = True Then
        Dim refCount As Integer
        Debug.Print VCS_String.VCS_PadRight("Exporting references...", 24);
        obj_count = VCS_Reference.VCS_ExportReferences(source_path)
        Debug.Print "[" & obj_count & "]"
    End If

'-------------------------table export------------------------
    If ExportTablesTemp Then
            obj_path = source_path & "tables\"
            VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
            VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "txt"
            
            Dim td As DAO.TableDef
            Dim tds As DAO.TableDefs
            Set tds = Db.TableDefs

            obj_type_label = "tbldef"
            obj_type_name = "Table_Def"
            obj_type_num = acTable
            obj_path = source_path & obj_type_label & "\"
            obj_count = 0
            obj_data_count = 0
            VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
            
            'move these into Table and DataMacro modules?
            ' - We don't want to determin file extensions here - or obj_path either!
            If editedOnly = False Then
                VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "sql"
                VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "xml"
                VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "LNKD"
            End If
            
            Dim IncludeTablesCol As Collection
            Set IncludeTablesCol = StrSetToCol(INCLUDE_TABLES, ",")
            
            Debug.Print VCS_String.VCS_PadRight("Exporting " & obj_type_label & "...", 24);
            
            For Each td In tds
                    Dim skipTableNotEditedAfterCommit As Boolean
                    ' note that every module will always be exported, due to a quirk in
                    ' how modules' DateModified flag is set - but this is OK,
                    ' as modules are generally the fastest to export
                    skipTableNotEditedAfterCommit = editedOnly And _
                        CurrentCommitDate > td.LastUpdated
                    ' This is not a system table
                    ' this is not a temporary table
                    If Left$(td.name, 4) <> "MSys" And _
                    Left$(td.name, 1) <> "~" And _
                    skipTableNotEditedAfterCommit = False Then
                            If editedOnly = True Then VCS_Dir.VCS_DelIfExist (obj_path)
                            If Len(td.connect) = 0 Then ' this is not an external table
                                    VCS_Table.VCS_ExportTableDef td.name, obj_path
                                    If INCLUDE_TABLES = "*" Then
                                            DoEvents
                                            VCS_Table.VCS_ExportTableData CStr(td.name), source_path & "tables\"
                                            If Len(Dir$(source_path & "tables\" & td.name & ".txt")) > 0 Then
                                                    obj_data_count = obj_data_count + 1
                                            End If
                                    ElseIf (Len(Replace(INCLUDE_TABLES, " ", vbNullString)) > 0) And INCLUDE_TABLES <> "*" Then
                                            DoEvents
                                            On Error GoTo Err_TableNotFound
                                            If InCollection(IncludeTablesCol, td.name) Then
                                                    VCS_Table.VCS_ExportTableData CStr(td.name), source_path & "tables\"
                                                    obj_data_count = obj_data_count + 1
                                            End If
Err_TableNotFound:
                                            
                                    'else don't export table data
                                    End If
                            Else
                                    VCS_Table.VCS_ExportLinkedTable td.name, obj_path
                            End If
                            
                            obj_count = obj_count + 1
                            
                    End If
            Next
            Debug.Print "[" & obj_count & "]"
            If obj_data_count > 0 Then
              Debug.Print VCS_String.VCS_PadRight("Exported data...", 24) & "[" & obj_data_count & "]"
            End If
            
            Set IncludeTablesCol = StrSetToCol(INCLUDE_TABLES_PFS, ",")
            
            Debug.Print VCS_String.VCS_PadRight("Exporting PFS tables...", 24);
            
            ' todo: this is a lot of duplicate code
            For Each td In tds
                    Dim skipPfsTableNotEditedAfterCommit As Boolean
                    ' note that every module will always be exported, due to a quirk in
                    ' how modules' DateModified flag is set - but this is OK,
                    ' as modules are generally the fastest to export
                    skipPfsTableNotEditedAfterCommit = editedOnly And _
                        CurrentCommitDate > td.LastUpdated
                    ' This is not a system table
                    ' this is not a temporary table
                    If Left$(td.name, 4) <> "MSys" And _
                    Left$(td.name, 1) <> "~" And _
                    skipPfsTableNotEditedAfterCommit = False Then
                            If editedOnly = True Then VCS_Dir.VCS_DelIfExist (obj_path)
                            If Len(td.connect) = 0 Then ' this is not an external table
                                    VCS_Table.VCS_ExportTableDef td.name, obj_path
                                    If INCLUDE_TABLES = "*" Then
                                            DoEvents
                                            VCS_Table.VCS_ExportTableData CStr(td.name), source_path_pfs & "tables\"
                                            If Len(Dir$(source_path_pfs & "tables\" & td.name & ".txt")) > 0 Then
                                                    obj_data_count = obj_data_count + 1
                                            End If
                                    ElseIf (Len(Replace(INCLUDE_TABLES, " ", vbNullString)) > 0) And INCLUDE_TABLES <> "*" Then
                                            DoEvents
                                            On Error GoTo Err_TablePFSNotFound
                                            If InCollection(IncludeTablesCol, td.name) Then
                                                    VCS_Table.VCS_ExportTableData CStr(td.name), source_path_pfs & "tables\"
                                                    obj_data_count = obj_data_count + 1
                                            End If
Err_TablePFSNotFound:
                                            
                                    'else don't export table data
                                    End If
                            Else
                                    VCS_Table.VCS_ExportLinkedTable td.name, obj_path
                            End If
                            
                            obj_count = obj_count + 1
                            
                    End If
            Next
            Debug.Print "[" & obj_count & "]"
            If obj_data_count > 0 Then
              Debug.Print VCS_String.VCS_PadRight("Exported data...", 24) & "[" & obj_data_count & "]"
            End If
            
            
            Debug.Print VCS_String.VCS_PadRight("Exporting Relations...", 24);
            obj_count = 0
            obj_path = source_path & "relations\"
            VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))

            VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "txt"

            Dim aRelation As DAO.Relation
            
            For Each aRelation In CurrentDb.Relations
                    ' Exclude relations from system tables and inherited (linked) relations
                    ' Skip if dbRelationDontEnforce property is not set. The relationship is already in the table xml file. - sean
                    If Not (aRelation.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" _
                                    Or aRelation.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups" _
                                    Or (aRelation.Attributes And DAO.RelationAttributeEnum.dbRelationInherited) = _
                                    DAO.RelationAttributeEnum.dbRelationInherited) _
                       And ((aRelation.Attributes And DAO.RelationAttributeEnum.dbRelationDontEnforce) = _
                                    DAO.RelationAttributeEnum.dbRelationDontEnforce) _
                       Then
                            VCS_Relation.VCS_ExportRelation aRelation, obj_path & aRelation.name & ".txt"
                            obj_count = obj_count + 1
                    End If
            Next
            Debug.Print "[" & obj_count & "]"
    End If
        
    Debug.Print "Done."
End Sub

Public Sub ExportAllReports()
    Call ExportSource(True, False, False, False, False, False, False, False, False)
End Sub

Public Sub ExportAllQueries()
    Call ExportSource(False, True, False, False, False, False, False, False, False)
End Sub

Public Sub ExportAllForms()
    Call ExportSource(False, False, True, False, False, False, False, False, False)
End Sub

Public Sub ExportAllMacros()
    Call ExportSource(False, False, False, True, False, False, False, False, False)
End Sub

Public Sub ExportAllModules()
    Call ExportSource(False, False, False, False, True, False, False, False, False)
End Sub

Public Sub ExportAllTables()
    Call ExportSource(False, False, False, False, False, True, False, False, False)
End Sub

Public Sub ExportAllReferences()
    Call ExportSource(False, False, False, False, False, False, True, False, False)
End Sub

Public Sub ExportAllSource(Optional ByVal isButton As Boolean)
    Call ExportSource(True, True, True, True, True, True, isButton, False, False)
End Sub

Public Sub ExportChanges()
    ' note: relations and modules will always have all content exported
    Call ExportSource(True, True, True, True, True, True, True, False, True)
End Sub

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportSource(ByVal ImportReports As Boolean, ByVal ImportQueries As Boolean, ByVal ImportForms As Boolean, ByVal ImportMacros As Boolean, _
        ByVal ImportModules As Boolean, ByVal ImportTables As Boolean, ByVal ImportReferences As Boolean, ByVal isButton As Boolean, ByVal editedOnly As Boolean)
    Dim FSO As Object
    Dim source_path As String
    Dim obj_path As String
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim fileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean
    
    Dim includeTables As Boolean
    
    If isButton = True Then
        includeTables = False
    Else
        includeTables = True
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")

    CloseFormsReports
    
    ' If this import is of changed files only,
    ' get the list of changed files and copy them to a separate
    ' source directory. Then, run import on that directory instead.
    ' todo: does this account for PFS files?
    If editedOnly = True Then
        source_path = VCS_Dir.VCS_ProjectPath() & "changes-only-source\"
        VCS_Dir.ClearAndMakeTemporaryChangesDirectory source_path
        Dim changedFiles As Collection
        Dim changedFile As Variant
        Dim changedFileName As String
        Set changedFiles = VCS_GitOperations.GetSourceFilesSinceLastImport()
        For Each changedFile In changedFiles
            Dim changedFileCopy As String
            changedFileCopy = Replace(changedFile, "source\", "")
            FSO.CopyFile VCS_Dir.VCS_ProjectPath() & changedFile, _
                VCS_Dir.VCS_ProjectPath() & "changes-only-source\" & changedFileCopy
        Next
    Else
        source_path = VCS_Dir.VCS_ProjectPath() & "source\"
    End If
        
    If Not FSO.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    Debug.Print
    
    If ImportReferences = True Then
        Dim refCount As Integer
        Debug.Print VCS_String.VCS_PadRight("Importing references... ", 24);
        refCount = VCS_Reference.VCS_ImportReferences(source_path)
        Debug.Print "[" & refCount & "]"
    End If
    
    obj_path = source_path & "queries\"
    fileName = Dir$(obj_path & "*.bas")
    
    Dim tempFilePath As String
    tempFilePath = VCS_File.VCS_TempFile()
    
    If Len(fileName) > 0 And ImportQueries = True Then
        Debug.Print VCS_String.VCS_PadRight("Importing queries...", 24)
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
            obj_name = Replace(obj_name, "%backslash%", "\")
            obj_name = Replace(obj_name, "%forwardslash%", "/")
            obj_name = Replace(obj_name, "%colon%", ":")
            obj_name = Replace(obj_name, "%asterisk%", "*")
            obj_name = Replace(obj_name, "%questionmark%", "?")
            obj_name = Replace(obj_name, "%doublequote%", """")
            obj_name = Replace(obj_name, "%leftarrow%", "<")
            obj_name = Replace(obj_name, "%rightarrow%", ">")
            obj_name = Replace(obj_name, "%pipe%", "|")
            'Check for plain sql export/import
                        If HandleQueriesAsSQL Then
                                VCS_Query.ImportQueryFromSQL obj_name, obj_path & fileName, False
                        Else
                                VCS_IE_Functions.VCS_ImportObject acQuery, obj_name, obj_path & fileName, VCS_File.VCS_UsingUcs2
                                VCS_IE_Functions.VCS_ExportObject acQuery, obj_name, tempFilePath, VCS_File.VCS_UsingUcs2
                                VCS_IE_Functions.VCS_ImportObject acQuery, obj_name, tempFilePath, VCS_File.VCS_UsingUcs2
                        End If
                        obj_count = obj_count + 1
            fileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    VCS_Dir.VCS_DelIfExist tempFilePath

    If includeTables = True And ImportTables = True Then
    ' restore table definitions
        obj_path = source_path & "tbldef\"
        fileName = Dir$(obj_path & "*.xml")
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.VCS_PadRight("Importing tabledefs...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                If DebugOutput Then
                    If obj_count = 0 Then
                        Debug.Print
                    End If
                    Debug.Print "  [debug] table " & obj_name;
                    Debug.Print
                End If
                VCS_Table.VCS_ImportTableDef CStr(obj_name), obj_path
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
        
        
        ' restore linked tables - we must have access to the remote store to import these!
        fileName = Dir$(obj_path & "*.LNKD")
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.VCS_PadRight("Importing Linked tabledefs...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                If DebugOutput Then
                    If obj_count = 0 Then
                        Debug.Print
                    End If
                    Debug.Print "  [debug] table " & obj_name;
                    Debug.Print
                End If
                VCS_Table.VCS_ImportLinkedTable CStr(obj_name), obj_path
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
        
        
        
        ' NOW we may load data
        obj_path = source_path & "tables\"
        fileName = Dir$(obj_path & "*.txt")
    
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.VCS_PadRight("Importing tables...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                DoEvents
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                VCS_Table.VCS_ImportTableData CStr(obj_name), obj_path
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
    
    
    
    ' load data for pfs
    
    'load Data Macros - not DRY!
        obj_path = source_path & "tbldef\"
        fileName = Dir$(obj_path & "*.dm")
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.VCS_PadRight("Importing Data Macros...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                DoEvents
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                'VCS_Table.VCS_ImportTableData CStr(obj_name), obj_path
                VCS_DataMacro.VCS_ImportDataMacros obj_name, obj_path
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
    End If
    

    ' import macros, forms, reports, modules
    Dim importObjectsString As String
    importObjectsString = ""
    
    If ImportForms = True Then importObjectsString = importObjectsString & "forms|" & acForm & ","
    If ImportReports = True Then importObjectsString = importObjectsString & "reports|" & acForm & ","
    If ImportMacros = True Then importObjectsString = importObjectsString & "macros|" & acReport & ","
    If ImportModules = True Then importObjectsString = importObjectsString & "modules|" & acModule
    
    For Each obj_type In Split(importObjectsString, ",")
    
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        obj_type_num = Val(obj_type_split(1))
        obj_path = source_path & obj_type_label & "\"
        
            
        fileName = Dir$(obj_path & "*.bas")
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.VCS_PadRight("Importing " & obj_type_label & "...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                ' DoEvents no good idea!
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = VCS_File.VCS_UsingUcs2
                End If
                If IsNotVCS(obj_name) Then
                    VCS_IE_Functions.VCS_ImportObject obj_type_num, obj_name, obj_path & fileName, ucs2
                    obj_count = obj_count + 1
                Else
                    If ArchiveMyself Then
                            MsgBox "Module " & obj_name & " could not be updated while running. Ensure latest version is included!", vbExclamation, "Warning"
                    End If
                End If
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        
        End If
    Next
    
    'import Print Variables
    If ImportReports = True Then
        Debug.Print VCS_String.VCS_PadRight("Importing Print Vars...", 24);
        obj_count = 0
        
        obj_path = source_path & "reports\"
        fileName = Dir$(obj_path & "*.pv")
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
            VCS_Report.VCS_ImportPrintVars obj_name, obj_path & fileName
            obj_count = obj_count + 1
            fileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    If includeTables = True And ImportTables = True Then
    'import relations
        Debug.Print VCS_String.VCS_PadRight("Importing Relations...", 24);
        obj_count = 0
        obj_path = source_path & "relations\"
        fileName = Dir$(obj_path & "*.txt")
        Do Until Len(fileName) = 0
            DoEvents
            VCS_Relation.VCS_ImportRelation obj_path & fileName
            obj_count = obj_count + 1
            fileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    DoEvents
    
    ' clean up source path
    If editedOnly = True And FSO.FolderExists(source_path) = True Then
        FSO.DeleteFolder Left(source_path, Len(source_path) - 1)
        VCS_GitOperations.SetLastImportedCommitToCurrent
    End If
    
    Debug.Print "Done."
End Sub

Public Sub ImportAllReports()
    Call ImportSource(True, False, False, False, False, False, False, False, False)
End Sub

Public Sub ImportAllQueries()
    Call ImportSource(False, True, False, False, False, False, False, False, False)
End Sub

Public Sub ImportAllForms()
    Call ImportSource(False, False, True, False, False, False, False, False, False)
End Sub

Public Sub ImportAllMacros()
    Call ImportSource(False, False, False, True, False, False, False, False, False)
End Sub

Public Sub ImportAllModules()
    Call ImportSource(False, False, False, False, True, False, False, False, False)
End Sub

Public Sub ImportAllTables()
    Call ImportSource(False, False, False, False, False, True, False, False, False)
End Sub

Public Sub ImportAllReferences()
    Call ImportSource(False, False, False, False, False, False, True, False, False)
End Sub

Public Sub ImportAllSource(Optional ByVal isButton As Boolean)
    Call ImportSource(True, True, True, True, True, True, False, isButton, False)
End Sub

Public Sub ImportChanges()
    ' note: relations and modules will always have all content exported
    Call ImportSource(True, True, True, True, True, True, True, False, True)
End Sub

' Main entry point for ImportProject.
' Drop all forms, reports, queries, macros, modules.
' execute ImportAllSource.
Public Sub ImportProject(Optional ByVal isButton As Boolean)
    On Error GoTo ErrorHandler

    Dim includeTables As Boolean
    
    If isButton = True Then
        includeTables = False
    Else
        includeTables = True
    End If
    
    If MsgBox("This action will delete all existing: " & vbCrLf & _
              vbCrLf & _
              IIf(includeTables, Chr$(149) & " Tables" & vbCrLf, "") & _
              Chr$(149) & " Forms" & vbCrLf & _
              Chr$(149) & " Macros" & vbCrLf & _
              Chr$(149) & " Modules" & vbCrLf & _
              Chr$(149) & " Queries" & vbCrLf & _
              Chr$(149) & " Reports" & vbCrLf & _
              vbCrLf & _
              "Are you sure you want to proceed?", vbCritical + vbYesNo, _
              "Import Project") <> vbYes Then
        Exit Sub
    End If

    Dim Db As DAO.Database
    Set Db = CurrentDb
    CloseFormsReports

    Debug.Print
    Debug.Print "Deleting Existing Objects"
    Debug.Print
    
    ' only delete tables & relations if var is true
    If includeTables = True Then
        Debug.Print "Deleting table relations"
        Dim rel As DAO.Relation
        For Each rel In CurrentDb.Relations
            If Not (rel.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or _
                    rel.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
                CurrentDb.Relations.DELETE (rel.name)
            End If
        Next
    End If
            
            ' First gather all Query Names.
            ' If you delete right away, the iterator loses track and only deletes every 2nd Query
    Dim toBeDeleted As Collection
    Set toBeDeleted = New Collection
    Dim qryName As Variant
    
    Debug.Print "Deleting queries"
    Dim dbObject As Object
    For Each dbObject In Db.QueryDefs
        DoEvents
        If Left$(dbObject.name, 1) <> "~" Then
                        toBeDeleted.Add dbObject.name
        End If
    Next

    
    For Each qryName In toBeDeleted
        Db.QueryDefs.DELETE qryName
    Next
        
        Set toBeDeleted = Nothing
    If includeTables = True Then
        Debug.Print "Deleting table defs"
        Dim td As DAO.TableDef
        For Each td In CurrentDb.TableDefs
            If Left$(td.name, 4) <> "MSys" And _
                Left$(td.name, 1) <> "~" Then
                CurrentDb.TableDefs.DELETE (td.name)
            End If
        Next
    End If

    Dim objType As Variant
    Dim objTypeArray() As String
    Dim doc As Object
    '
    '  Object Type Constants
    Const OTNAME As Byte = 0
    Const OTID As Byte = 1

    For Each objType In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
            , "," _
        )
        objTypeArray = Split(objType, "|")
        DoEvents
        For Each doc In Db.Containers(objTypeArray(OTNAME)).Documents
            DoEvents
            If (Left$(doc.name, 1) <> "~") And _
               (IsNotVCS(doc.name)) Then
'                Debug.Print doc.Name
                DoCmd.DeleteObject objTypeArray(OTID), doc.name
            End If
        Next
    Next
    
    Debug.Print "================="
    Debug.Print "Importing Project"
    ImportAllSource (isButton)
    
    Exit Sub

ErrorHandler:
    Debug.Print "VCS_ImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
                Err.Description
End Sub


'===================================================================================================================================
'-----------------------------------------------------------'
' Helper Functions - these should be put in their own files '
'-----------------------------------------------------------'

' Close all open forms.
Private Sub CloseFormsReports()
    On Error GoTo ErrorHandler
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).name
        DoEvents
    Loop
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).name
        DoEvents
    Loop
    Exit Sub

ErrorHandler:
    Debug.Print "VCS_ImportExport.CloseFormsReports: Error #" & Err.Number & vbCrLf & _
                Err.Description
End Sub


'errno 457 - duplicate key (& item)
Private Function StrSetToCol(ByVal strSet As String, ByVal delimiter As String) As Collection 'throws errors
    Dim strSetArray() As String
    Dim col As Collection
    
    Set col = New Collection
    strSetArray = Split(strSet, delimiter)
    
    Dim strPart As Variant
    For Each strPart In strSetArray
        col.Add strPart, strPart
    Next
    
    Set StrSetToCol = col
End Function


' Check if an item or key is in a collection
Private Function InCollection(col As Collection, Optional vItem, Optional vKey) As Boolean
    On Error Resume Next

    Dim vColItem As Variant

    InCollection = False

    If Not IsMissing(vKey) Then
        col.item vKey

        '5 if not in collection, it is 91 if no collection exists
        If Err.Number <> 5 And Err.Number <> 91 Then
            InCollection = True
        End If
    ElseIf Not IsMissing(vItem) Then
        For Each vColItem In col
            If vColItem = vItem Then
                InCollection = True
                GoTo Exit_Proc
            End If
        Next vColItem
    End If

Exit_Proc:
    Exit Function
Err_Handle:
    Resume Exit_Proc
End Function