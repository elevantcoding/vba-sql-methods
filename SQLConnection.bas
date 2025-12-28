Attribute VB_Name = "SQLConnection"
Option Compare Database
Option Explicit

' global reference to the SQL connection and command
Public SQLcn As ADODB.Connection
Public SQLcommand As ADODB.Command

Const ModName As String = "SQLConnection"
Public Function OpenSQL(Optional ByVal lTimeout As Long = 90) As Boolean
    On Error GoTo Except
    
    OpenSQL = False
    
    If Not SQLcn Is Nothing Then
        If SQLcn.State = adStateOpen Then
            OpenSQL = True
            Exit Function
        End If
    End If
    
    Set SQLcn = New ADODB.Connection
    SQLcn.ConnectionTimeout = lTimeout
    SQLcn.open ADOConnect
    OpenSQL = True
    
Finally:
    Exit Function
    
Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "OpenSQL", , ModName)
    Resume Finally
End Function
Public Sub OpenSQLCmd(Optional ByVal lTimeout As Long = 90, Optional ByVal lCmdType As ADODB.CommandTypeEnum = adCmdStoredProc)
    On Error GoTo Except
    
    If SQLcn Is Nothing Then Call OpenSQL
    
    If SQLcommand Is Nothing Then
        Set SQLcommand = New ADODB.Command
    Else
        With SQLcommand
            If .Parameters.count > 0 Then
                Do Until .Parameters.count = 0
                    .Parameters.Delete 0
                Loop
            End If
            .CommandText = vbNullString
        End With
    End If
    
    With SQLcommand
        .ActiveConnection = SQLcn
        .CommandType = lCmdType
        .CommandTimeout = lTimeout
        If lCmdType = adCmdStoredProc Then .NamedParameters = True
    End With
    
Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "OpenSQLCmd", , ModName)
    Resume Finally
    
End Sub
Public Sub CloseADORS(ByRef adors As ADODB.Recordset)
    On Error GoTo Except
    
    'clean up any named ado recordset
    'use ByRef to clear the calling procedure's object reference
    If Not adors Is Nothing Then
        If adors.State = adStateOpen Then adors.Close
        Set adors = Nothing
    End If

Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "CloseADORS", , ModName)
    Resume Finally

End Sub
Public Sub CloseSQL()
    On Error GoTo Except
    
    ' close the glboal SQL connection and ref
    If Not SQLcn Is Nothing Then
        If SQLcn.State = adStateOpen Then
            SQLcn.Close
        End If
        Set SQLcn = Nothing
    End If
    
Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "CloseSQL", , ModName)
    Resume Finally
End Sub


