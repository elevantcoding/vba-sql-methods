Attribute VB_Name = "Interaction"
Option Compare Database
Option Explicit

' public reference to dlgfrmNotification for brevity
Public Const MsgFrm As String = "dlgfrmNotification"

'used with RespondMsg
Public MsgResponse As Variant

Const ModName As String = "Interaction"
Public Sub DisplayMsg(Optional ByVal MainCap As String = " ", Optional ByVal Label1Cap As String = "", Optional ByVal Label2Cap As String = "", Optional ByVal Label3Cap As String = "", _
          Optional ByVal lFont As Long = 10, Optional ByVal lFormColor As Long = White, Optional ByVal lDisplaySeconds As Long = 0)
    On Error GoTo Except

    ' uses a 1 x 4 form without buttons for user notifications
    ' see dlgMessage.bas for the form module code
    ' form contains three (3) labels for which captions can be set
    ' at runtime
    ' options: form caption, three lines of information, font size, form color, number of seconds to display
    ' if using lDisplaySeconds, form will automatically close after n seconds
    ' else, the form will remain open and need to be closed after use
    
    ' global/public-level declaration for MsgFrm in this module
    Const MsgLabel1 As String = "One"
    Const MsgLabel2 As String = "Two"
    Const MsgLabel3 As String = "Three"
    
    Dim frm As Form, ctl1 As Control, ctl2 As Control, ctl3 As Control, LabelForeColor As Long
    
    LabelForeColor = IIf(IsDarkColor(lFormColor), vbWhite, vbBlack)
    ' open form:
    ' if displaying for n seconds, if form is loaded, close and re-open with displayseconds in open args
    ' else, only open the form if it is not already open
    If lDisplaySeconds > 0 Then
        If FormIsLoaded(MsgFrm) Then DoCmd.Close acForm, MsgFrm
        DoCmd.OpenForm MsgFrm, acNormal, , , , , lDisplaySeconds
    Else
        If Not FormIsLoaded(MsgFrm) Then DoCmd.OpenForm MsgFrm, acNormal
    End If
   
    ' set the form and each label ctl
    Set frm = Forms(MsgFrm)
    Set ctl1 = Forms(MsgFrm).Controls(MsgLabel1)
    Set ctl2 = Forms(MsgFrm).Controls(MsgLabel2)
    Set ctl3 = Forms(MsgFrm).Controls(MsgLabel3)
    
    ' set form caption and form color
    ' set label captions and label forecolors
    frm.Caption = MainCap: frm.Detail.BackColor = lFormColor
    ctl1.Caption = Label1Cap: ctl1.ForeColor = LabelForeColor
    ctl2.Caption = Label2Cap: ctl2.ForeColor = LabelForeColor
    ctl3.Caption = Label3Cap: ctl3.ForeColor = LabelForeColor
    
    ' set font size for each label
    ctl1.FontSize = lFont
    ctl2.FontSize = lFont
    ctl3.FontSize = lFont
        
Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "DisplayMsg", , ModName)
    Resume Finally
    
End Sub
Public Sub CloseMsgFrm()
    On Error GoTo Except
    
    ' global/public-level declaration for MsgFrm in this module
    If FormIsLoaded(MsgFrm) Then DoCmd.Close acForm, MsgFrm
    
Finally:
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "CloseMsgFrm", , ModName)
    Resume Finally

End Sub
Public Function RespondMsg(ByVal FormTitle As String, ByVal Prompt As String, ByVal strCap1 As String, Optional ByVal strCap2 As String = "", _
    Optional ByVal strCap3 As String = "", Optional ByVal blnExpectResponse As Boolean = False, Optional ByVal lDisplaySeconds As Long = 0, _
    Optional ByVal DefaultBtnIndex As Integer = 1) As Variant
    On Error GoTo Except

    ' uses a 1 1/2 x 4 form containing three command buttons
    ' Call RespondMsg in code to display a custom message box
    ' FormTitle - shows at top of dialog form
    ' Prompt - form main caption within label of a form
    ' strCap1 through strCap3 - button captions
    ' blnExpectResponse - whether response / input is expected
    ' lDisplaySeconds - if greater than zero, dialog form will display for lDisplaySeconds
    ' DefaultBtnIndex - which button will be given the focus / is the default response option
    ' uses MsgResponse, public variant
    
    Dim strArgs As String, C As Long, vCaptions As Variant, cCount As Long
    
    Const DialogFrm As String = "dlgfrmMessage"
    Const ProcName As String = "RespondMsg"
    Const sep As String = ";"
    Const noresponse As String = "N/A"
    Const maxOpenArgLen As Long = 2048
    
    ' initial function return value
    RespondMsg = noresponse
    
    ' public variant defined in this module
    MsgResponse = noresponse
    
    ' trim trailing / leading spaces
    strCap1 = Trim(strCap1): strCap2 = Trim(strCap2): strCap3 = Trim(strCap3)
    
    ' determine which captions have non-blank strings: strCap1, strCap2, strCap3
    vCaptions = Array(strCap1, strCap2, strCap3)
    cCount = 0
    For C = LBound(vCaptions) To UBound(vCaptions)
        If Len(vCaptions(C)) > 0 Then
            cCount = cCount + 1
        End If
    Next
    
    ' ensure proper count of captions indicated for method selected
    ' define main caption prefix
    If blnExpectResponse Then
        If cCount < 2 Then
            MsgBox "If requesting a response, at least two response options must be specified.", vbInformation, "ResponseMsg"
            Exit Function
        End If
        FormTitle = "Response Requested: " & FormTitle
    Else
        If cCount > 1 Then
            MsgBox "If not request input, more than one response option is not needed.", vbInformation, "ResponseMsg"
            Exit Function
        End If
        FormTitle = "Information: " & FormTitle
    End If
    
    ' information to pass to form open args
    strArgs = FormTitle & sep & _
        Prompt & sep & _
        strCap1 & sep & _
        strCap2 & sep & _
        strCap3 & sep & _
        CStr(blnExpectResponse) & sep & _
        lDisplaySeconds & sep & _
        DefaultBtnIndex
    
    ' if length of created args is too long, is err
    If Len(strArgs) > maxOpenArgLen Then Call RaiseCustomErr(OpenArgs_Len, ProcName, Len(strArgs) & " characters in OpenArgs.")
    
    ' open form as dialog
    DoCmd.OpenForm DialogFrm, , , , , acDialog, strArgs
    
    'set response
    RespondMsg = MsgResponse
    
    'close the form, which was hidden
    If IsLoaded(DialogFrm) Then DoCmd.Close acForm, DialogFrm
    
Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "RespondMsg", , ModName)
    Resume Finally
    
End Function
Public Sub OpenAnyFileType(ByVal strPath As String, ByVal strOrigin As String)
    On Error GoTo Except
  
    Dim strFileName As String
    Dim i As Long
    
    ' normalize path sep
    strPath = Replace(strPath, "/", "\")
        
    ' validate existence
    If Not SetFso.FileExists(strPath) Then
        i = InStrRev(strPath, "\")
        strFileName = Mid(strPath, i + 1)
        MsgBox "The file " & strFileName & " could not be found at " & Left(strPath, i), vbOKOnly + vbInformation, "Open File"
        Exit Sub
    End If
    
    ' open file
    CreateObject("WScript.Shell").Run """" & strPath & """", 1, False
    
Finally:
    Exit Sub

Except:
    MsgBox "OpenAnyFileType command encountered an error while attempting to open: " & vbCrLf & strPath, vbInformation, "OpenAnyFileType"
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "OpenAnyFileType", , ModName, , strOrigin)
    Resume Finally

End Sub
Public Function MyFileLocation() As String
    On Error GoTo Except
    
    Dim strPath As String
    
    ' set location to store text file reports from this system
    strPath = Environ("USERPROFILE") & "\Desktop\MyFiles\"
    
    If Not SetFso.FolderExists(strPath) Then
        SetFso.CreateFolder (strPath)
    End If

    MyFileLocation = strPath
    
Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "MyFileLocation", , ModName)
    Resume Finally
    
End Function
Public Function IsValidFileOrPath(ByVal strPath As String) As Boolean
    On Error GoTo Except
    
    Dim p As String
    p = Replace(strPath, "/", "\")
    IsValidFileOrPath = (SetFso.FileExists(p) Or SetFso.FolderExists(p))
    
Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "IsValidFileOrPath", , ModName)
    Resume Finally

End Function



