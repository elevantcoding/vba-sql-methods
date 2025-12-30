'--------------------------
'dlgfrmMessage
'--------------------------
Option Compare Database
Option Explicit

Dim vSec As Variant
Dim strMsg As String
Dim blnExpectResponse As Boolean
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Except

    Dim ctl As Control
    Dim defaultBtn As Variant
    
    Const ProcName As String = "Form_Open"
    
    'if no args found, exit
    'if not multiple args, exit
    If IsNull(Me.OpenArgs) Or Nz(Me.OpenArgs, "") = "" Or Not IsMultipleValues(Me.OpenArgs) Then
        Call RaiseCustomErr(OpenArgs_ExpVal, ProcName)
    End If
            
    ' 0 to 7: FormTitle, Prompt, BtnCaption1, BtnCaption2, BtnCaption3, Input, DiplaySeconds, DefaultButton
    'use arg values to set captions and button visibility
    Me.Caption = GetOpenArgValue(Me.OpenArgs, 0)
    Me.Main.Caption = GetOpenArgValue(Me.OpenArgs, 1)
    
    'set each btn, if not a blank string, visible = True
    Set ctl = Me.Controls("btn1")
    ctl.Caption = GetOpenArgValue(Me.OpenArgs, 2)
    ctl.Visible = Len(ctl.Caption) > 0
    
    Set ctl = Me.Controls("btn2")
    ctl.Caption = GetOpenArgValue(Me.OpenArgs, 3)
    ctl.Visible = Len(ctl.Caption) > 0
    
    Set ctl = Me.Controls("btn3")
    ctl.Caption = GetOpenArgValue(Me.OpenArgs, 4)
    ctl.Visible = Len(ctl.Caption) > 0
    
    blnExpectResponse = CBool(GetOpenArgValue(Me.OpenArgs, 5))
        
    'if seconds passed are greater than 0, set timer and caption of displaytimer label for timeout if no response received
    vSec = GetOpenArgValue(Me.OpenArgs, 6)
    If IsNumeric(vSec) Then
        If vSec > 0 Then
            Me.TimerInterval = 1000
            If blnExpectResponse Then
                strMsg = vSec & " seconds remaining for response."
            Else
                strMsg = "This notification will automatically close in " & vSec & " seconds."
            End If
            Me.DisplayTimer.Caption = strMsg
        End If
    End If
    
    'get value of arg 7, this will be the default response option
    defaultBtn = val(Nz(GetOpenArgValue(Me.OpenArgs, 7), 1))
    
    Select Case defaultBtn
        Case 1: If Me.btn1.Visible Then Me.btn1.SetFocus
        Case 2: If Me.btn2.Visible Then Me.btn2.SetFocus
        Case 3: If Me.btn3.Visible Then Me.btn3.SetFocus
    End Select

    DoCmd.Beep
    
Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, , , Me.Name)
    Cancel = True
    Resume Finally

End Sub
Private Sub Form_Timer()
    On Error GoTo Except
    
    'if timer interval has been set, show timeout reminder
    'if timeout, then set response as such and close form (time out when less than 0)
    If blnExpectResponse Then
        strMsg = vSec & " seconds remaining for response."
    Else
        strMsg = "This notification will automatically close in " & vSec & " seconds."
    End If
    Me.DisplayTimer.Caption = strMsg
    vSec = vSec - 1
    If vSec < 0 Then
        MsgResponse = "Timeout"
        Me.Form.Visible = False
    End If
    
Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, , , Me.Name)
    Resume Finally

End Sub
Private Sub btn1_Click()
    On Error GoTo Except

    'for each button, public var msgresponse will contain the button's caption to know which response
    'was selected
    MsgResponse = Me.btn1.Caption
    Me.Form.Visible = False

Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, , CtlType, Me.Name)
    Resume Finally

End Sub
Private Sub btn2_Click()
    On Error GoTo Except

    MsgResponse = Me.btn2.Caption
    Me.Form.Visible = False

Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, , CtlType, Me.Name)
    Resume Finally

End Sub
Private Sub btn3_Click()
    On Error GoTo Except

    MsgResponse = Me.btn3.Caption
    Me.Form.Visible = False

Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, , CtlType, Me.Name)
    Resume Finally

End Sub


