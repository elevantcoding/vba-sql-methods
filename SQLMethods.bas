Attribute VB_Name = "SQLMethods"
Option Compare Database
Option Explicit

'--------------------------------
' enumerations used in functions
' in this module
'--------------------------------
Public Enum SQLActionType
    sqlInsert
    sqlUpdate
End Enum

Public Enum SQLSearchType
    sqlSearchAll = 0
    sqlSearchView
    sqlSearchProc
    sqlSearchFunction
    sqlSearchTrigger
    sqlSearchDefault
End Enum

Public Enum SQLObjectType
    sqlTable = 1
    sqlView
    sqlProc
    sqlFunction
    sqlTrigger
End Enum

Const ModName As String = "SQLMethods"
Public Function SQLCount(ByVal strSQL As String) As Long
    On Error GoTo Except
    
    ' get a record count directly from SQL Server
    ' helpful for tables or views not linked to an Access project
    Dim adors As ADODB.Recordset
    Dim lRecordCount As Long
    Dim blnSelectCount As Boolean
    Dim strDesc As String
            
    ' default record count
    lRecordCount = 0
    
    ' find out if call is using SELECT COUNT
    blnSelectCount = False
    If InStr(1, strSQL, "SELECT COUNT(", vbTextCompare) > 0 Then blnSelectCount = True
    
    If Not blnSelectCount Then
        MsgBox "SQLCount requires a SELECT COUNT(...) statement.", vbInformation, "SQLCount"
        Exit Function
    End If
    
    ' open the connection to SQL
    ' if no connection, exit
    If Not OpenSQL Then MsgBox "Could not open the SQL connection.", vbInformation, "SQLCount": Exit Function
    
    ' set the record set reference
    Set adors = New ADODB.Recordset
    
    ' open and retrieve results
    adors.open strSQL, SQLcn, adOpenForwardOnly, adLockReadOnly
    If Not adors.EOF Then lRecordCount = CLng(adors.Fields(0).Value)
    
    SQLCount = lRecordCount
       
Finally:
    Call CloseADORS(adors)
    ' optional, close connection to SQL if not persistent session connection
    Exit Function

Except:
    strDesc = Err.Description
    Select Case Err.Number
        Case -2147217887 ' mal-formed statement
            MsgBox "Please check the SQL statement passed to SQLCount.  SQL is unable to parse the statement.", vbInformation, "SQLCount: SQL Structure"
        Case -2147217865
            If InStr(1, strDesc, "Invalid object name", vbTextCompare) > 0 Then ' named object not found
                MsgBox "Cannot find table or view named in " & strSQL & ".", vbInformation, "SQLCount: Schema, Table or View Name"
            ElseIf InStr(1, strDesc, "permission was denied", vbTextCompare) > 0 Then ' permissions
                MsgBox "Please check permissions to the objects named in the SELECT COUNT statement.", vbInformation, "SQLCount: Permissions"
            End If
        Case 3146, 3151 ' odbc connectivity
            MsgBox "Please check your connection.", vbOKOnly + vbInformation, "SQLCount: ODBC Error"
    End Select
    
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "SQLCount", , ModName)
    Resume Finally

End Function
Public Function SQLResult(ByVal strSQL As String, Optional ByVal strResultColumn As String = "") As Variant
    On Error GoTo Except

    ' retrieve a one-row result from SQL Server directly
    ' optionally specify the column to return results from
    ' else, will return results from first column
    Dim adors As ADODB.Recordset
    Dim vResult As Variant
    Dim lFieldCount As Long
    Dim strDesc As String
    
    ' default value
    vResult = Null
    
    ' open global connection
    ' confirm connection is open
    If Not OpenSQL Then MsgBox "Could not open the SQL connection.", vbInformation, "SQLResult": Exit Function
    
    ' create the recordset
    Set adors = New ADODB.Recordset
    adors.CursorLocation = adUseServer
        
    ' open read-only recordset
    adors.open strSQL, SQLcn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' if valid recordset / count records
    If Not adors.EOF Then
    
        ' get count of columns
        lFieldCount = CLng(adors.Fields.count)
                
        ' if more than 1 column:
        ' if strResultColumn is blank, return result of first column, else return result of strResultColumn
        ' else return column result
        If lFieldCount > 1 And strResultColumn <> "" Then
            vResult = adors.Fields(strResultColumn).Value
        Else
            vResult = adors.Fields(0).Value
        End If
        
        'if more than one row, notify and clear result
        adors.MoveNext
        If Not adors.EOF Then
            MsgBox "SQLResult is designed to return a one-row result but the SELECT statement " & strSQL & " is resulting " & _
                "in more than one row.  Please check the SELECT statement.", vbOKOnly + vbInformation, "SQLResult"
            vResult = Null
        End If
    End If
    
    ' return result
    SQLResult = vResult
    
Finally:
    Call CloseADORS(adors)
    ' optional, close connection to SQL if not persistent session connection
    Exit Function

Except:
    strDesc = Err.Description
    Select Case Err.Number
        Case 3265 ' named column not found in the select statement
            MsgBox "Referenced column " & strResultColumn & " is not found in the SELECT statement: " & vbCrLf & vbCrLf & strSQL & ".", vbInformation, "SQLResult: " & strResultColumn

        Case -2147217887 ' invalid SQL statement
            MsgBox "Please check the SQL statement passed to SQLResult.  SQL is unable to parse the statement.", vbInformation, "SQLResult: SQL Structure"

        Case -2147217900 ' column does not exist in the SELECT statement
            MsgBox "Column name referenced in the SELECT statement does not exist in the table, or if you're referencing a function, make sure it exists in the schema.", vbInformation, "SQLResult: Column Name or Possibly Function"
        
        Case -2147217865
            If InStr(1, strDesc, "Invalid object name", vbTextCompare) > 0 Then ' named object not found
                MsgBox "Cannot find table or view named in " & strSQL & ".", vbInformation, "SQLResult: Schema, Table or View Name"
            ElseIf InStr(1, strDesc, "permission was denied", vbTextCompare) > 0 Then ' permissions
                MsgBox "Please check permissions to the objects named in the SELECT COUNT statement.", vbInformation, "SQLResult: Permissions"
            End If
    End Select
    
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "SQLResult", , ModName)
    Resume Finally
End Function
Public Function SQLInsertUpdate(ByVal strSQL As String, ByVal ActionType As SQLActionType) As Long
    On Error GoTo Except
    
    'this function uses the following declarations:
    'Public Enum SQLActionType
    '    sqlInsert
    '    sqlUpdate
    'End Enum
    
    'process a one-row update or insert directly to SQL
    'Update: return records affected
    'Insert: return SCOPE_IDENTITY()
    Dim adors As ADODB.Recordset
    Dim lFind As Long, lRecordsAffected As Long
    Dim strAction As String, strMsg As String, strDesc As String
    Dim blnTrans As Boolean 'marker: is True while the transaction state is open and uncommitted
    
    SQLInsertUpdate = 0
    
    Select Case ActionType
        Case sqlInsert
            strAction = "INSERT INTO "
        Case sqlUpdate
            strAction = "UPDATE "
        Case Else
            MsgBox "Undefined / unknown SQL action type.", vbInformation, "SQLInsertUpdate: " & ActionType
            Exit Function
    End Select
    lFind = InStr(1, strSQL, strAction, vbTextCompare)
    
    If lFind = 0 Then
        MsgBox strAction & " not found in the SQL statement.", vbInformation, "SQLInsertUpdate: " & strAction
        Exit Function
    End If

    ' confirm connection is open
    If Not OpenSQL Then MsgBox "Could not open the SQL connection.", vbInformation, "SQLInsertUpdate": Exit Function
    
    SQLcn.BeginTrans
    blnTrans = True
    SQLcn.Execute strSQL, lRecordsAffected, adExecuteNoRecords ' execute, but no resultset needed
    
    'if 0 or more than one, notify after rollback, clean up
    'if 1, if Insert, set function result as SCOPE_IDENTITY(), else, return records affected
    Select Case lRecordsAffected
        Case 0, Is > 1
            SQLcn.RollbackTrans
            blnTrans = False
            strMsg = IIf(lRecordsAffected = 0, "No rows affected.", "Rolled back: more than one row would be affected.")
            MsgBox strMsg, vbOKOnly + vbInformation, "SQLInsertUpdate"
            GoTo ExitProcessing
        Case 1
            If ActionType = sqlInsert Then
                Set adors = SQLcn.Execute("SELECT SCOPE_IDENTITY() As NewID")
                If Not adors.EOF Then SQLInsertUpdate = CLng(Nz(adors!NewID, 0)) 'scope identity of inserted record
            Else
                SQLInsertUpdate = lRecordsAffected ' records affected by update
            End If
    End Select
    
    SQLcn.CommitTrans
    blnTrans = False

ExitProcessing:
Finally:
    Call CloseADORS(adors)
    ' optional, close connection to SQL if not persistent session connection
    Exit Function

Except:
    If blnTrans Then
        SQLcn.RollbackTrans
        blnTrans = False
    End If
    
    strDesc = Err.Description
    Select Case Err.Number
        Case -2147217887 'parsing of SQL / unrecognized named column / etc.
            MsgBox "Please check the SQL statement passed to SQLInsertUpdate.  SQL is unable to parse the statement.", vbOKOnly + vbInformation, "SQLInsertUpdate: SQL Structure"
        
        Case -2147217865
            If InStr(1, strDesc, "Invalid object name", vbTextCompare) > 0 Then ' named object not found
                MsgBox "Cannot find table or view named in " & strSQL & ".", vbInformation, "SQLInsertUpdate: Schema, Table or View Name"
            ElseIf InStr(1, strDesc, "permission was denied", vbTextCompare) > 0 Then ' permissions
                MsgBox "Please check permissions to the objects named in the SQL statement.", vbInformation, "SQLInsertUpdate: Permissions"
            End If
        
        Case 3146, 3151 'odbc connectivity
            MsgBox "Please check your connection.", vbOKOnly + vbInformation, "ODBC Error"
    End Select
    
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "SQLInsertUpdate", , ModName)
    Resume Finally

End Function
Public Sub SQLInfo(ByVal strToFind As String, Optional ByVal strSchema As String = "", Optional ByVal ObjType As SQLSearchType = sqlSearchAll)
On Error GoTo Except
    
    ' search SQL module definition for strToFind
    ' return written text file
    
    ' this function uses the following declarations:
    ' Public Enum SQLSearchType
    '   sqlSearchAll
    '   sqlSearchView
    '   sqlSearchProc
    '   sqlSearchFunction
    '   sqlSearchTrigger
    '   sqlSearchDefault
    ' End Enum

    Dim adors As ADODB.Recordset

    Dim stream As Object
    
    Dim strSQL As String
    Dim strFileName As String
    Dim strFilePath As String
    Dim strHeading As String
    Dim strObjType As String
    Dim strObjTypeName As String
    
    Dim lResults As Long
    
    Dim blnStream As Boolean: blnStream = False
    
    Dim GetType() As Variant
    Dim GetAbb() As Variant
    
    Const ProcName As String = "SQLInfo"
    
    ' array order must match SQLSearchType enum order
    GetType = Array("All Objects", "Views", "Procedures", "Functions", "Triggers", "Default Values")
    GetAbb = Array("ALL", "V", "P", "FN", "TR", "DEFAULT")
    
    ' if is out of bounds, exit
    If ObjType < LBound(GetType) Or ObjType > UBound(GetType) Then
        MsgBox ObjType & " is not a valid / defined option.", vbInformation, ProcName & ": " & ObjType
        Exit Sub
    End If
    
    strObjTypeName = GetType(ObjType)
    strObjType = GetAbb(ObjType)
    strToFind = Replace(strToFind, "'", "''")

    Select Case ObjType
        Case sqlSearchDefault
            strSQL = "SELECT TABLE_NAME As [name], COLUMN_NAME As [column_name], COLUMN_DEFAULT As [default_value] " & _
                "FROM INFORMATION_SCHEMA.COLUMNS "
            
            If Len(strSchema) > 0 Then
                strSQL = strSQL & "WHERE TABLE_SCHEMA = '" & strSchema & "' AND COLUMN_DEFAULT Is Not Null AND COLUMN_DEFAULT Like '%" & strToFind & "%'"
            Else
                strSQL = strSQL & "WHERE COLUMN_DEFAULT Is Not Null AND COLUMN_DEFAULT Like '%" & strToFind & "%'"
            End If
        
        Case sqlSearchView, sqlSearchProc, sqlSearchFunction, sqlSearchTrigger, sqlSearchAll
            strSQL = "SELECT * FROM (SELECT o.object_id, s.name As [schema_name], o.name As object_name, o.type, o.type_desc As object_type, m.definition " & _
                "FROM sys.sql_modules m INNER JOIN sys.objects o ON m.object_id = o.object_id " & _
                "INNER JOIN sys.schemas s ON o.schema_id = s.schema_id " & _
                "UNION ALL SELECT t.object_id, NULL As [schema_name], t.[name] As object_name, t.type, t.[type_desc] As object_type, m.[definition] " & _
                "FROM sys.sql_modules m INNER JOIN sys.triggers t ON m.object_id = t.object_id WHERE t.parent_class_desc = 'DATABASE') moddefinitions "
            
            If ObjType = sqlSearchAll Then
                If Len(strSchema) > 0 Then
                    strSQL = strSQL & "WHERE schema_name = '" & strSchema & "' AND [definition] LIKE '%" & strToFind & "%'"
                Else
                    strSQL = strSQL & "WHERE [definition] LIKE '%" & strToFind & "%'"
                End If
            Else
                If Len(strSchema) > 0 Then
                    strSQL = strSQL & "WHERE schema_name = '" & strSchema & "' AND [type] ='" & strObjType & "' AND [definition] LIKE '%" & strToFind & "%'"
                Else
                    strSQL = strSQL & "WHERE [type] ='" & strObjType & "' AND [definition] LIKE '%" & strToFind & "%'"
                End If
                
            End If
    End Select
    
    ' exit if no connection
    If Not OpenSQL Then MsgBox "Could not open the SQL connection.", vbInformation, "SQLInfo": Exit Sub

    strHeading = "SQL Server " & strObjTypeName & ", Containing: " & strToFind
    strFileName = "SQLView" & Format(Now(), "mdyyyy_hhnnss") & ".txt"
    strFilePath = MyFileLocation & strFileName

    Set adors = New ADODB.Recordset
    adors.open strSQL, SQLcn, adOpenForwardOnly, adLockReadOnly

    lResults = 0
    If Not adors.EOF Then
        
        Set stream = SetFso.CreateTextFile(strFilePath, True, False) 'file, overwrite, ansi
        blnStream = True
        stream.WriteLine strHeading
        stream.WriteLine "-----------------------------------------"

        ' display notification while process is occurring
        Call DisplayMsg(, , "Reviewing", "Find: " & strToFind)
        DoCmd.RepaintObject acForm, MsgFrm
        DoEvents
            
        Do Until adors.EOF
            
            lResults = lResults + 1

            If strObjType = Default Then
                stream.WriteLine lResults & " - " & adors.Fields("object_name") & " - " & adors.Fields("column_name")
            Else
                stream.WriteLine lResults & " - " & adors.Fields("object_name")
            End If

            adors.MoveNext
        Loop
        stream.Close
        blnStream = False
    End If
    
    ' display notification of results for 2 seconds
    Call DisplayMsg(, , "Found", lResults & " Results", , , 2)
    DoCmd.RepaintObject acForm, MsgFrm
    DoEvents
    
    If lResults = 0 Then
        MsgBox "No results found for " & strToFind & ".", vbInformation, "SQLInfo"
    Else
        If MsgBox("Do you wish to view the results?", vbYesNo + vbQuestion, "Text File Results") = vbYes Then
            Call OpenAnyFileType(strFilePath, strFileName)
        End If
    End If

Finally:
    Call CloseMsgFrm
    Call CloseADORS(adors)
    If Not stream Is Nothing Then
        If blnStream Then stream.Close
        Set stream = Nothing
    End If
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, ProcName, , ModName)
    Resume Finally
End Sub
Private Function IsValidSQLObject(ByVal strSchema As String, ByVal strObjName As String, _
    Optional ByVal SQLObj As SQLObjectType = 0) As Boolean
    
    On Error GoTo Except

    ' query sql sys.objects, sys.triggers for name of object and optional type
    ' this function uses RemBrackets to clear left and right bracktes from schema, objname, if exists
    ' this function uses SQLResult
    ' this function uses the following enumeration, specified in this module's declarations:
    ' Public Enum SQLObjectType
    ' sqlTable = 1
    ' sqlView
    ' sqlProc
    ' sqlFunction
    ' sqlTrigger
    ' End Enum
    ' enumeration deliberately starts at 1 so 0 can be used
    ' to indicate non-specified object type

    Dim strSQL As String, strSQLType As String
    Dim GetSQLType() As Variant
    Const ProcName As String = "IsValidSQLObject"
    
    ' escape ' and remove brackets
    strSchema = Replace(strSchema, "'", "''")
    strSchema = RemBrackets(strSchema)
    
    strObjName = Replace(strObjName, "'", "''")
    strObjName = RemBrackets(strObjName)
    
    ' include "schema" for DDL triggers
    strSchema = "'" & strSchema & "', 'DDL'"
    
    ' if specified a type, get appropriate search terms for sql
    ' array order must match SQLObjectType enum order
    ' enum is 1-based so for 0, use blank string
    GetSQLType = Array("", "'U'", "'V'", "'P'", "'FN', 'IF'", "'TR'")
    If SQLObj < LBound(GetSQLType) Or SQLObj > UBound(GetSQLType) Then
        MsgBox SQLObj & " is not a valid / defined option.", vbInformation, ProcName & ": " & SQLObj
        Exit Function
    End If
    
    strSQLType = GetSQLType(SQLObj)
    
    strSQL = "SELECT 1 FROM (SELECT o.name, o.type, s.name As [schema_name] " & _
        "FROM sys.objects o INNER JOIN sys.schemas s ON o.schema_id = s.schema_id " & _
        "UNION All SELECT t.name, t.type, 'DDL' As [schema_name] " & _
        "FROM sys.triggers t WHERE t.parent_class_desc = 'DATABASE') allobjects " & _
        "WHERE [name] = '" & strObjName & "' AND [schema_name] IN (" & strSchema & ")"
    
    If SQLObj > 0 Then
        strSQL = strSQL & " AND [type] IN (" & strSQLType & ")"
    End If
    
    ' use one-row result SQLResult (variant) to return results
    IsValidSQLObject = (SQLResult(strSQL) > 0)
    
Finally:
    Exit Function
    
Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, ProcName, , ModName)
    Resume Finally
End Function
Public Function BuildParameter(ByVal cmd As ADODB.Command, ByVal Value As Variant, ByVal datatype As ADODB.DataTypeEnum, _
    Optional ByVal size As Variant = -1) As ADODB.Parameter
    On Error GoTo Except

    Dim p As ADODB.Parameter
    Dim direction As ADODB.ParameterDirectionEnum
    
    direction = adParamInput  ' default direction

    If IsEmpty(size) Or IsNull(size) Or size = -1 Then
        Set p = cmd.CreateParameter(, datatype, direction, , Value)
    Else
        Set p = cmd.CreateParameter(, datatype, direction, size, Value)
    End If
    
    If datatype = adNumeric Then
        p.Precision = 18
        p.NumericScale = 2
    End If
    
    Set BuildParameter = p
    
Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "BuildParameter", , ModName)
    Resume Finally
    
End Function
Public Function SQLExportSPToXL(ByVal strFullFilePath As String, ByVal strSPName As String, Optional ByVal arrParams As Variant, Optional ByVal arrTypes As Variant, _
    Optional ByVal arrSizes As Variant, Optional blnNotify As Boolean = False, Optional ByVal blnDisplay As Boolean = False) As Boolean
    
    On Error GoTo Except

    Dim adors As ADODB.Recordset
    
    Dim strMsg As String, strFilePath As String, strFileName As String, strSchema As String
    
    Dim i As Long, col As Long
    
    Dim appXL As Object
    Dim wb As Object
    Dim wksht As Object
    Dim rng As Object
    
    Const defaultschema As String = "dbo"
    
    If Not IsValidExcelFile(strFullFilePath, strMsg, strFilePath, strFileName) Then
        MsgBox strMsg, vbInformation, "SQLExportSPToXL"
        Exit Function
    End If
    
    If SetFso.FileExists(strFilePath + strFileName) Then
        If MsgBox("The file " & strFileName & " already exists in " & strFilePath & ".  Overwrite?", vbYesNo + vbQuestion, "SQLExportSPToXL") = vbNo Then
            Exit Function
        Else
            SetFso.DeleteFile (strFilePath + strFileName)
        End If
    End If
    
    ' parse the proc name: look for schema, use default if not a schema specified
    ' pass to IsValidSQLObject to ensure the proc exists / is accessible
    ' IsValidSQLObject will remove the brackets to parse
    ' adding brackets here ensures when calling the procedure it will still parse in case of spaces in proc name
    i = InStr(1, strSPName, ".", vbTextCompare)
        
    If i = 0 Then
        strSchema = defaultschema
        strSPName = strSPName
    Else
        strSchema = Left(strSPName, i - 1)
        strSPName = Mid(strSPName, i + 1)
    End If
        
    strSchema = AddBrackets(strSchema)
    strSPName = AddBrackets(strSPName)
    
    If Not IsValidSQLObject(strSchema, strSPName, sqlProc) Then
        If Trim(strSPName) = "" Then
            MsgBox "Procedure name cannot be blank.", vbInformation, "SQLExportSPToXL"
        Else
            MsgBox "Specified procedure: " & strSPName & " not found.", vbInformation, "SQLExportSPToXL"
        End If
        GoTo ExitProcessing
    End If
    
    ' use SQL command for stored procedures - if no params, simply execute
    ' input params only
    Call OpenSQLCmd
    SQLcommand.CommandText = strSchema & "." & strSPName
    If Not IsMissing(arrParams) And IsArray(arrParams) Then
        If (Not IsEmpty(arrParams) And UBound(arrParams)) >= 0 Then
            If UBound(arrParams) <> UBound(arrTypes) Or UBound(arrParams) <> UBound(arrSizes) Then
                MsgBox "Parameter, type, and size arrays must be the same length.", vbInformation, "SQLExportSPToXL"
                GoTo ExitProcessing
            End If
            For i = 0 To UBound(arrParams)
                SQLcommand.Parameters.Append BuildParameter(SQLcommand, arrParams(i), arrTypes(i), arrSizes(i))
            Next
        End If
    End If
        
    ' execute
    Set adors = SQLcommand.Execute
    
    ' if no records, exit
    If adors.EOF Then
        MsgBox "The procedure returned zero records.", vbInformation, "SQLExportSPToXL"
        GoTo ExitProcessing
    End If
    
    ' display notification
    Call DisplayMsg(, "Creating Excel File")
    DoCmd.RepaintObject acForm, MsgFrm
    DoEvents
    
    ' set up excel
    Set appXL = CreateObject("Excel.Application")
    appXL.Visible = blnDisplay
    Set wb = appXL.Workbooks.add
    wb.SaveAs strFilePath + strFileName
    
    Set wksht = wb.ActiveSheet
    
    ' write column names
    For col = 0 To adors.Fields.count - 1
        wksht.Cells(1, col + 1).Value = adors.Fields(col).Name
    Next
    
    ' write recordset
    wksht.Range("A2").CopyFromRecordset adors
    
    ' set font
    With wksht
        Set rng = .Cells
        rng.Font.Name = "Arial"
        rng.Font.size = "10"
    End With
    wb.Save
    
    ' proc is true
    SQLExportSPToXL = True
    
    ' if excel not already displayed and blnNotify is true, confirm open file
    If Not blnDisplay And blnNotify Then
        If MsgBox(strFileName & " successfully created at " & strFilePath & ".  Open File?", vbYesNo + vbQuestion, "Open?") = vbYes Then
            Call OpenAnyFileType(strFilePath + strFileName, "SQLExportSPToXL")
        End If
    End If

ExitProcessing:
Finally:
    Call CloseMsgFrm
    Call CloseADORS(adors)
    If Not rng Is Nothing Then Set rng = Nothing
    If Not blnDisplay Then
        If Not wksht Is Nothing Then Set wksht = Nothing
        If Not wb Is Nothing Then wb.Close SaveChanges:=True
        Set wb = Nothing
        If Not appXL Is Nothing Then
            On Error Resume Next
            appXL.Quit
            Set appXL = Nothing
        End If
        On Error GoTo Except
    End If
    Exit Function

Except:
    Select Case Err.Number
        Case 70 ' file is open
            MsgBox "Please close the file first.", vbInformation, "SQLExportSPToXL"
        Case Else
            MsgBox "Err Number " & Err.Number & " with description " & Err.Description & " on line " & Erl
    End Select
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "SQLExportSPToXL", , ModName)
    Resume Finally
    
End Function
Public Function SQLExportToXL(ByVal strFullFilePath As String, ByVal strSQL As String, Optional ByVal blnAddWorksheet As Boolean = False, _
    Optional ByVal strWkShtName As String = "", Optional blnSQLServer As Boolean = True, Optional ByVal blnNotify As Boolean = False, _
    Optional blnDisplay As Boolean = False) As Boolean
    
    On Error GoTo Except
    
    Dim adors As ADODB.Recordset
    
    Dim appXL As Object
    Dim wb As Object
    Dim wksht As Object
    Dim rng As Object
    
    Dim blnWkBookExists As Boolean
    Dim blnChkWksht As Boolean
    
    Dim i As Long
    Dim col As Long
    
    Dim strMsg As String
    Dim strFileName As String
    Dim strFilePath As String
    Dim strRep As String
    
    SQLExportToXL = False
    
    If Not IsValidExcelFile(strFullFilePath, strMsg, strFilePath, strFileName) Then
        MsgBox strMsg, vbInformation, "SQLExportToXL"
        Exit Function
    End If
    
    Call DisplayMsg(, "Just a Moment")
    DoCmd.RepaintObject acForm, MsgFrm
    DoEvents
    
    ' ensure SQL statement has been indicated
    strSQL = Trim(strSQL)
    If strSQL = "" Then
        MsgBox "Please indicate an SQL statement for the export.", vbInformation, "SQLExportToXL"
        GoTo ExitProcessing
    End If
        
    ' make sure SELECT, not insert, etc.
    If InStr(1, strSQL, "SELECT ", vbTextCompare) = 0 Then
        MsgBox "SELECT not found in SQL statement: " & strSQL & ".", vbInformation, "SQLExportToXL"
        GoTo ExitProcessing
    End If
    
    ' ensure valid worksheet name
    If strWkShtName <> "" Then
        If Not IsValidSheetName(strWkShtName, strMsg) Then
            MsgBox strMsg, vbInformation, "SQLExportToXL"
            GoTo ExitProcessing
        End If
    End If
    
    ' determine if file exists
    blnWkBookExists = False
    If SetFso.FileExists(strFilePath & strFileName) Then blnWkBookExists = True

    If blnAddWorksheet Then
        If Not blnWkBookExists Then
            MsgBox strFilePath & strFileName & " does not exist.  Cannot add a worksheet to non-existent file.", vbInformation, "SQLExportToXL"
            GoTo ExitProcessing
        End If
        If strWkShtName <> "" Then blnChkWksht = True
    Else
        If blnWkBookExists Then
            If MsgBox("The file " & strFileName & " already exists in " & strFilePath & ".  Overwrite?", vbYesNo + vbQuestion, "SQLExportToXL") = vbNo Then
                Exit Function
            Else
                SetFso.DeleteFile (strFilePath + strFileName)
            End If
        End If
    End If
    
    ' set the ado record set object
    Set adors = New ADODB.Recordset
        
    If blnSQLServer Then
        ' ensure proper tokens
        i = SearchPredicate(strSQL, "#", True, "'", strRep)
        If i > 1 Then strSQL = strRep
            
        i = SearchPredicate(strSQL, "*", True, "%", strRep)
        If i > 0 Then strSQL = strRep
            
        ' use SQL connection
        Call OpenSQL
        adors.open strSQL, SQLConnect, adOpenForwardOnly, adLockReadOnly
    Else
        ' ensure proper tokens
        i = SearchPredicate(strSQL, "'", True, "#", strRep)
        If i > 1 Then strSQL = strRep
            
        i = SearchPredicate(strSQL, "%", True, "*", strRep)
        If i > 0 Then strSQL = strRep
            
        ' use current project connection / Access linked or local tables
        adors.open strSQL, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
    End If
    
    If adors.EOF Then
        MsgBox "There are no records in the specified SQL statement.", vbOKOnly + vbInformation, "Records Availability"
        GoTo ExitProcessing
    End If
    
    Call DisplayMsg(, "Creating Excel File")
    DoCmd.RepaintObject acForm, MsgFrm
    DoEvents
    
    Set appXL = CreateObject("Excel.Application")
             
    'set display or do not display excel file while exporting
    appXL.Visible = blnDisplay
    
    If blnAddWorksheet Then
            
        Call DisplayMsg(, "Updating Excel File")
        DoCmd.RepaintObject acForm, MsgFrm
        DoEvents
            
        appXL.Workbooks.open (strFilePath + strFileName)
        Set wb = appXL.ActiveWorkbook
            
        'make sure no duplicate worksheet name
        If blnChkWksht Then
            For Each wksht In wb.Worksheets
                If wksht.Name = strWkShtName Then
                    MsgBox "There is already a worksheet named " & strWkShtName & "  Export cannot continue.", vbOKOnly + vbInformation, "Please Note"
                    GoTo ExitProcessing
                End If
            Next
        End If
            
        appXL.Worksheets.add
        wb.Save
    Else
        Set wb = appXL.Workbooks.add
        wb.SaveAs strFilePath + strFileName
    End If
    
    Call DisplayMsg(, "Creating Excel File")
    DoCmd.RepaintObject acForm, MsgFrm
    DoEvents
    
    Set wksht = wb.ActiveSheet
        
    'write column names first
    For col = 0 To adors.Fields.count - 1
        wksht.Cells(1, col + 1).Value = adors.Fields(col).Name
    Next
    
    'write entire recordset to Excel beginning at A2
    wksht.Range("A2").CopyFromRecordset adors
    
    'set font and size of font to entire worksheet
    With wksht
        Set rng = .Cells
        rng.Font.Name = "Arial"
        rng.Font.size = "10"
        
        'if worksheet name specified, edit name of wksht
        If strWkShtName <> "" Then .Name = strWkShtName
    End With
    
    SQLExportToXL = True
    
    'complete all automation ops before displaying notification
    If Not blnDisplay And blnNotify Then
        If MsgBox(strFileName & " successfully created at " & strFilePath & ".  Open File?", vbYesNo + vbQuestion, "Open?") = vbYes Then
            Call OpenAnyFileType(strFilePath + strFileName, "SQLExportToXL")
        End If
    End If
    
ExitProcessing:
Finally:
    Call CloseMsgFrm
    Call CloseADORS(adors)
    If Not rng Is Nothing Then Set rng = Nothing
    If Not wksht Is Nothing Then Set wksht = Nothing
    If Not blnDisplay Then
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        Set wb = Nothing
        If Not appXL Is Nothing Then
            On Error Resume Next
            appXL.Quit
            Set appXL = Nothing
        End If
    End If
        
Except:
    Select Case Err.Number
        Case 70 'named wbk is already open
            MsgBox "Please close the spreadsheet before exporting again.", vbOKOnly + vbInformation, "Spreadsheet Open"
    
        Case -2147217865, 3078 'no table named
            MsgBox "Table name specified in the SQL statement is not found.", vbOKOnly + vbInformation, "No Such Table"
    
        Case -2147217900, 3061 'sql statement parsing (whether sql or access)
            MsgBox "The statement " & strSQL & " could not be parsed or may have an invalid column reference.", vbOKOnly + vbInformation, "SQL"
        Case Else
            Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "SQLExportToXL", , ModName)
    End Select
    Resume Finally
    
End Function
Private Function AddBrackets(ByVal strInput As String) As String
    On Error GoTo Except

    If Len(strInput) = 0 Then Exit Function

    strInput = Trim(strInput)
    
    If Right(strInput, 1) <> "]" Then strInput = strInput & "]"
    If Left(strInput, 1) <> "[" Then strInput = "[" & strInput
    
    AddBrackets = strInput
    
Finally:
    Exit Function
    
Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "AddBrackets", , ModName)
    Resume Finally
    
End Function
Private Function RemBrackets(ByVal strInput As String) As String
    On Error GoTo Except
    
    If Len(strInput) = 0 Then Exit Function
    
    strInput = Trim(strInput)
    
    If Right(strInput, 1) = "]" Then strInput = Left(strInput, Len(strInput) - 1)
    If Left(strInput, 1) = "[" Then strInput = Right(strInput, Len(strInput) - 1)
    
    RemBrackets = strInput

Finally:
    Exit Function
    
Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "RemBrackets", , ModName)
    Resume Finally

End Function
Private Function SearchPredicate(ByVal strSQL As String, ByVal findChar As String, Optional ByVal blnRevise As Boolean = False, _
    Optional ByVal strReplaceWith As String = "", Optional ByRef strRevisedSQL As String = "") As Long
    On Error GoTo Except
    
    ' this function performs context-aware predicate scanning rather than blind string replacement.
    
    ' replace chars in SQL string predicate:
    ' intended for toggling between Access and SQL: * or %, ' or #.
    ' alternatively, will replace other strings in a predicate without prejudice
    ' note: this function avoids counting or replacing characters in bracketed columns
    ' note: does not account for bracketed column names containing escaped brackets
    ' note: could be extended to account for insert, update and delete statements
    
    Dim i As Long, strLen As Long, findLen As Long, inWhe As Long, inHav As Long, countFind As Long, countApos As Long, firstPos As Long
    Dim lastPos As Long, firstCharPos As Long
    
    Dim findWhe As String, findHav As String, findSel As String, findUni As String, findGrp As String, findOrd As String
    Dim char As String, nextChar As String, prevChar As String, nextChars2 As String, strVal As String
    Dim blnColumn As Boolean, blnString As Boolean
    
    Const Sel As String = " SELECT "
    Const Uni As String = " UNION "
    Const Whe As String = " WHERE "
    Const Grp As String = " GROUP BY "
    Const Hav As String = " HAVING "
    Const Ord As String = " ORDER BY "
    
    ' default value
    SearchPredicate = 0
    
    ' make sure contains select statement (or if add insert / update / delete, validate)
    If InStr(1, strSQL, LTrim(Sel), vbTextCompare) = 0 Then
        MsgBox "SELECT not found.", vbInformation, "SearchPredicate"
        Exit Function
    End If
    
    ' if using revision option, strReplaceWith must contain a value
    If blnRevise And strReplaceWith = "" Then
        MsgBox "Must indicate a character(s) to replace with.", vbInformation, "SearchPredicate"
        Exit Function
    End If
    
    ' initialize
    strLen = Len(strSQL)
    findLen = Len(findChar)
    blnColumn = False
    char = "": nextChar = "": prevChar = "": nextChars2 = ""
    inWhe = 0: inHav = 0: countFind = 0: countApos = 0: firstPos = 0: lastPos = 0: firstCharPos = 0
    strRevisedSQL = strSQL
    
    ' main loop
    For i = 1 To strLen
    
        ' char for i in strSQL
        char = Mid(strSQL, i, findLen)
        
        ' char for i - 1 in strSQL
        If i > 1 Then
            prevChar = Mid(strSQL, i - 1, findLen)
        End If
        
        ' char for i + 1
        If i < strLen - 1 Then
            nextChar = Mid(strSQL, i + 1, findLen)
        Else
            nextChar = ""
        End If
        
        ' next two chars
        If i < strLen - (findLen + 1) Then
            nextChars2 = Mid(strSQL, i + 1, findLen + 1)
        Else
            nextChars2 = ""
        End If
        
        ' skip read of escaped apostrophe inside of string
        If blnString And nextChars2 = "''" Then i = i + 2
        
        ' find various query clauses
        If i >= Len(Whe) Then findWhe = Mid(strSQL, i, Len(Whe))
        If i >= Len(Hav) Then findHav = Mid(strSQL, i, Len(Hav))
        If i >= Len(Sel) Then findSel = Mid(strSQL, i, Len(Sel))
        If i >= Len(Uni) Then findUni = Mid(strSQL, i, Len(Uni))
        If i >= Len(Grp) Then findGrp = Mid(strSQL, i, Len(Grp))
        If i >= Len(Ord) Then findOrd = Mid(strSQL, i, Len(Ord))
                
        ' when find where or having, set read range
        ' based on i + length of either
        If findWhe = Whe Then inWhe = i + Len(Whe)
        If findHav = Hav Then inHav = i + Len(Hav)
        
        ' when clause is select, union, group by or order by, clear read range
        If findSel = Sel Or findUni = Uni Or findGrp = Grp Or findOrd = Ord Then
            inWhe = 0
            inHav = 0
        End If
        
        ' when predicate (where or having) has been defined range,
        ' only count findChar when outside of possible
        ' bracketed columns
        If inWhe > 0 Or inHav > 0 Then
            
            ' bracketed columns, where brackets not within a string
            If Not blnString Then
                If char = "[" Then blnColumn = True
                If char = "]" Then blnColumn = False
            End If
            
            If Not blnColumn Then ' if not a bracketed column
                If char = "'" Or char = "#" Then ' count apostrophes or # (# functions like apostrophe for a date)
                    countApos = countApos + 1
                        
                    If countApos = 1 Then
                        blnString = True ' beginning of string
                        firstPos = i ' position of found '
                        firstCharPos = i + 1 ' position of first char in string after '
                    Else
                        blnString = False ' end of string
                        countApos = 0 ' restart count of apostrophes
                        lastPos = i ' position of found '
                    End If
                End If
            End If
            
            If char = findChar Then
                If char = "*" Or char = "%" Then
                    If blnString Then
                        If prevChar = "'" Or nextChar = "'" Then
                            countFind = countFind + 1 ' count * or % inside a string only, next to an apos
                            If blnRevise Then ' if revising, be certain exchange is between % and *
                                If (strReplaceWith = "%" And findChar = "*") Or (strReplaceWith = "*" And findChar = "%") Then
                                    strRevisedSQL = ReplaceCharAtIndex(strRevisedSQL, i, strReplaceWith)
                                End If
                            End If
                        End If
                    End If
                    
                ElseIf char = "'" Or char = "#" Then
                    If Not blnColumn And Not blnString And lastPos > 0 Then ' when encounter either apos char, determine if value between is a date
                        strVal = Mid(strSQL, firstCharPos, lastPos - firstCharPos)
                        If IsValidDate(strVal) Then
                            countFind = countFind + 2 ' increment count by 2 for both chars
                            If blnRevise Then
                                strRevisedSQL = ReplaceCharAtIndex(strRevisedSQL, firstPos, strReplaceWith)
                                strRevisedSQL = ReplaceCharAtIndex(strRevisedSQL, lastPos, strReplaceWith)
                            End If
                        End If
                    End If
                    
                Else
                    If Not blnColumn Then
                        countFind = countFind + 1
                        If blnRevise Then strRevisedSQL = ReplaceCharAtIndex(strRevisedSQL, i, strReplaceWith)
                    End If
                End If
            End If ' char = findChar
            
        End If ' inWhe, inHav
    Next
    
    ' return result of count
    SearchPredicate = countFind

Finally:
    Exit Function
    
Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "SearchPredicate", , ModName)
    Resume Finally
End Function





