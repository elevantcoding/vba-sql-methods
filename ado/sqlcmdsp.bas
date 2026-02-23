' publicly-available vars for SQL connection
Public SQLConnect As ADODB.Connection
Public SQLCmd As ADODB.Command
Public SQLPrm As ADODB.Parameter

' mSQL module connection string
Private Function ADOConnect() As String    
    ADOConnect = "YourConnectionString"
End Function

' global SQL connection
Public Sub OpenSQL(Optional ByVal lTimeout As Long = 90)
    On Error GoTo Except
    
    Const ProcName As String = "OpenSQL"
    
    If Not SQLConnect Is Nothing Then
        If SQLConnect.State = adStateOpen Then Exit Sub
        Set SQLConnect = Nothing
    End If
    
    Set SQLConnect = New ADODB.Connection
    SQLConnect.ConnectionTimeout = lTimeout
    SQLConnect.Open ADOConnect
    
Finally:
    Exit Sub

Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Sub

-- the type of command, defaults to stored procedure
Sub SQLCmdAsType(Optional ByVal lTimeout As Long = 90, Optional ByVal CmdType As ADODB.CommandTypeEnum = adCmdStoredProc)
    On Error GoTo Except
    
    Const ProcName As String = "SQLCmdAsType"
    If SQLConnect Is Nothing Then OpenSQL
    
    If SQLCmd Is Nothing Then
        Set SQLCmd = New ADODB.Command
    Else
        With SQLCmd
            If .Parameters.Count > 0 Then
                Do Until .Parameters.Count = 0
                    .Parameters.Delete 0
                Loop
            End If
        End With
    End If
    
    With SQLCmd
        .ActiveConnection = SQLConnect
        .CommandType = CmdType
        .CommandTimeout = lTimeout
        .NamedParameters = (CmdType = adCmdStoredProc)
    End With
    
Finally:
    Exit Sub

Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Sub

Public Function SQLCmdSP(ByVal cmdText As String, ParamArray SP() As Variant) As Boolean
    On Error GoTo Except

    Const ProcName As String = "SQLCmdSP"
    Dim i As Long, p As Variant, MsgDetail As String: MsgDetail = "When calling SP " & cmdText
    
    ' designed to be called using SPParam if sp has params
    
    ' default value
    SQLCmdSP = False
    
    ' ensure stored proc name not zero length string
    cmdText = Trim(cmdText)
    If Len(cmdText) = 0 Then Exit Function
    
    ' use global SQLCmd
    SQLCmdAsType 30, adCmdStoredProc
    
    ' set procedure name
    With SQLCmd
        .CommandText = cmdText
        
        ' if SP called with params
        If UBound(SP) >= 0 Then
            
            ' loop through SP param array
            For i = LBound(SP) To UBound(SP)
                
                ' loop each array
                p = SP(i)
                
                ' make sure is array
                If Not IsArray(p) Then RaiseCustomMsg SysArray, ProcName, MsgDetail
                
                ' expect between five and 7 params, lower bound 0
                If Not IsBetween(UBound(p), 4, 6) Then RaiseCustomMsg SysSQLSPParams, ProcName, MsgDetail
                
                ' if prm type is string and prm size is 0, raise err
                If IsIn(p(1), adVarChar, adVarWChar) And p(3) = 0 Then RaiseCustomMsg SysSQLSPParamsSize, ProcName, MsgDetail
                
                ' set parameters
                Set SQLPrm = .CreateParameter(p(0), p(1), p(2), p(3), p(4))
                
                ' if is decimal param, set precision and scale
                If p(1) = adDecimal Then
                    SQLPrm.Precision = p(5)
                    SQLPrm.NumericScale = p(6)
                End If
                
                ' append
                .Parameters.Append SQLPrm
            Next
        End If
        ' exec proc
        .Execute
    End With
    
    ' function successfully exec
    ' results are available in SQLCmd.Parameters(varname)
    SQLCmdSP = True
    
Finally:
    Exit Function

Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function

-- use to call SQLCmdSP
Public Function SPParam(ByVal PrmName As String, ByVal PrmType As ADODB.DataTypeEnum, ByVal PrmDir As ADODB.ParameterDirectionEnum, _
          ByVal PrmSize As Long, ByVal PrmVal As Variant, Optional DecPrecision As Long = 0, Optional DecScale As Long = 0) As Variant
    On Error GoTo Except
    
    Const ProcName As String = "SPParam"
    
    ' if decimal,
    ' precision must not be less than or equal to zero,
    ' scale must not be less than 0,
    ' scale must not be greater than precision
    If PrmType = adDecimal Then
        If DecPrecision <= 0 _
            Or DecScale < 0 _
            Or DecScale > DecPrecision Then
            RaiseCustomMsg SysSQLSPParamsDecimal, ProcName
        End If
        SPParam = Array(PrmName, PrmType, PrmDir, PrmSize, PrmVal, DecPrecision, DecScale)
    Else
        SPParam = Array(PrmName, PrmType, PrmDir, PrmSize, PrmVal)
    End If

Finally:
    Exit Function

Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function

-- like T-SQL IN statement for comparing variable types, used in SQLCmdSP
Public Function IsIn(ByVal ValComp As Variant, ParamArray Vals() As Variant) As Boolean
    On Error GoTo Except
    
    Const ProcName As String = "IsIn"
    Dim I As Long
    
    For I = LBound(Vals) To UBound(Vals)
        If VarType(Vals(I)) = VarType(ValComp) Then
            If Vals(I) = ValComp Then
                IsIn = True
                Exit Function
            End If
        End If
    Next
    
    IsIn = False
    
Finally:
    Exit Function
Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function

-- order-agnostics value comparision
Public Function IsBetween(ByVal evalNum As Double, ByVal valOne As Double, ByVal valTwo As Double) As Boolean
    Dim val As Double
    If valOne > valTwo Then
        val = valOne
        valOne = valTwo
        valTwo = val
    End If
        
    IsBetween = (evalNum >= valOne And evalNum <= valTwo)

End Function
