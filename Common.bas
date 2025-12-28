Attribute VB_Name = "Common"
Option Compare Database
Option Explicit

Private m_fso As Object
    
Const ModName As String = "Common"
Public Function SetFso() As Object
    On Error GoTo Except
    
    'global file system reference, late binding
    If m_fso Is Nothing Then
        Set m_fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    Set SetFso = m_fso

Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "SetFso", , ModName)
    Resume Finally
End Function
Private Function IsValidAccessColor(ByVal lColor As Long) As Boolean
    On Error GoTo Except

    IsValidAccessColor = (lColor >= 0 And lColor <= &HFFFFFF)
    
Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "IsValidAccessColor", , ModName)
    Resume Finally
End Function
Public Function IsDarkColor(ByVal lAccessColor As Long) As Boolean
    On Error GoTo Except

    ' convert Access color code to R, G, B and return val for IsDarkColorRGB
    ' false as default for invalid colors
    Dim R As Long, G As Long, B As Long
    
    IsDarkColor = False
    
    If Not IsValidAccessColor(lAccessColor) Then Exit Function
    
    R = lAccessColor Mod 256
    G = (lAccessColor \ 256) Mod 256
    B = (lAccessColor \ 65536) Mod 256
    
    IsDarkColor = IsDarkColorRGB(R, G, B)
    
Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "IsDarkColor", , ModName)
    Resume Finally
End Function
Public Function IsDarkColorRGB(ByVal R As Long, ByVal G As Long, ByVal B As Long) As Boolean
    On Error GoTo Except

    'Perceived brightness formula (ITU-R BT.601)
    Dim luminance As Double
    luminance = (0.299 * R) + (0.587 * G) + (0.114 * B)
    IsDarkColorRGB = (luminance < 128)
    
Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "IsDarkColorRGB", , ModName)
    Resume Finally
End Function
Public Function FormIsLoaded(ByVal strFormName As String) As Boolean
    On Error GoTo Except

    ' detect if named-form is open: top-level, non-nested forms only
    Dim obj As AccessObject, db As Object
    
    Set db = Application.CurrentProject
    
    FormIsLoaded = False
    
    Set obj = db.AllForms(strFormName)
    
    If obj.IsLoaded Then
        If obj.CurrentView <> 0 Then
            FormIsLoaded = True
        End If
    End If

Finally:
    Exit Function

Except:
    Select Case Err.Number
        Case 2467
            MsgBox strFormName & " does not exist.", vbInformation, "FormIsLoaded"
        Case Else
            Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "FormIsLoaded", , ModName)
    End Select
    Resume Finally
End Function
Function IsValidExcelFile(ByVal strFullFilePath As String, Optional ByRef returnMsg As String = "", Optional ByRef strFilePath As String = "", _
    Optional ByRef strFileName As String = "") As Boolean
    On Error GoTo Except

    Dim strFileType As String
    Dim pos As Long

    IsValidExcelFile = False

    ' get path: if not path detected, return msg
    pos = InStrRev(strFullFilePath, "\")
    If pos = 0 Then
        returnMsg = "Invalid file path."
        Exit Function
    End If
    
    ' if no file name is specified, return msg
    If Right(strFullFilePath, 1) = "\" Then
        returnMsg = "File name missing from path."
        Exit Function
    End If
    
    ' detect valid path
    strFilePath = Left(strFullFilePath, pos)
    If Not fso.FolderExists(strFilePath) Then
        returnMsg = "The path " & strFilePath & " could not be found."
        Exit Function
    End If
    
    ' get file name
    pos = Len(strFullFilePath) - pos
    If pos = 0 Then
        returnMsg = "Invalid file name."
        Exit Function
    End If
    
    ' get file type
    strFileName = Right(strFullFilePath, pos)
    strFileType = GetFileType(strFileName)
    Select Case LCase(strFileType)
        Case "xls", "xlsx", "xlsm", "xlsb", "csv", "xml"
            ' okay
        Case Else
            returnMsg = "Not a valid Excel file type."
            Exit Function
    End Select
    
    IsValidExcelFile = True
    
Finally:
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "IsValidExcelFile", , ModName)
    Resume Finally
End Function
Public Function IsValidSheetName(ByVal strWksheetName As String, ByRef strReason As String) As Boolean
On Error GoTo Except

    Dim i As Long, strCharacters As String

    IsValidSheetName = False
    
    ' too long
    If Len(strWksheetName) > 31 Then
        strReason = "Worksheet Name specified in Parameters is too long."
        Exit Function
    End If
    
    ' too short
    If Len(strWksheetName) = 0 Then
        strReason = "Worksheet Name specified in Parameters is zero characters in length."
        Exit Function
    End If
    
    ' contains reserved word
    If strWksheetName = "History" Then
        strReason = "Excel worksheet cannot be 'History'."
        Exit Function
    End If
    
    ' cannot end a worksheet name with '
    If Right(strWksheetName, 1) = "'" Then
        strReason = "Worksheet name cannot end with an apostrophe."
        Exit Function
    End If
            
    ' contains invalid characters
    strCharacters = "\/?*[]:"
    For i = 1 To Len(strCharacters)
        If InStr(strWksheetName, Mid(strCharacters, i, 1)) > 0 Then
            strReason = "Invalid character in worksheet name: " & Mid(strCharacters, i, 1) & "."
            Exit Function
        End If
    Next

    IsValidSheetName = True

Finally:
    Exit Function
    
Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "IsValidSheetName", , ModName)
    Resume Finally

End Function

