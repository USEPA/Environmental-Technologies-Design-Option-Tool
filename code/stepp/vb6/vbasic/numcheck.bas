Attribute VB_Name = "NumCheckMod"
Global Temp_Text As String

Function HaveNumber(Value As Double) As Integer

    If Value > 0# Then HaveNumber = True Else HaveNumber = False

End Function

Function HaveNumber2(Value As Double) As Integer

' This subroutine was added to allow negative values to be entered in the
' user input section for the following properties :

     HaveNumber2 = True

End Function

Function HaveTemp(Value As Double) As Integer
   If Value > -250# Then HaveTemp = True Else HaveTemp = False
End Function

Sub NumberCheck(KeyAscii As Integer)
    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And KeyAscii <> Asc(".") And KeyAscii <> 8 And KeyAscii <> Asc("E") And KeyAscii <> Asc("e") And KeyAscii <> Asc("-") And KeyAscii <> Asc("+") Then
       KeyAscii = 0
       Beep
    End If

End Sub

Sub TextGetFocus(txt As TextBox, Temp_Text As String)
    Temp_Text = txt.Text
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)

End Sub

Sub TextHandleError(IsError As Integer, txt As TextBox, Temp_Text As String)
    Dim Dummy As Double
    Dim I As Integer

    IsError = False

    On Error GoTo ErrorHandler
       Dummy = CDbl(txt.Text)
'       If Dummy < 0# Then GoTo NegativeNumberError
       If IsError Then txt.SetFocus
       GoTo ContinueSub

ErrorHandler:
    IsError = True
    'frmAirWaterProperties.Print "Error Occurred"
    MsgBox "Incorrect Value Will Be Replaced By Previous Value", , "Invalid Data Error"
    txt.Text = Temp_Text
    If txt.Text = "" Then
       txt.SetFocus
       Exit Sub
    Else
       Resume
    End If


NegativeNumberError:
    IsError = True
    txt.Text = Temp_Text
    txt.SetFocus

ContinueSub:

End Sub

Sub TextNumberChanged(ValueChanged As Integer, txt As TextBox, Temp_Text As String)
    Dim Dummy1 As Double, Dummy2 As Double

    ValueChanged = True
    If Temp_Text = "" Then Exit Sub
    Dummy1 = CDbl(txt.Text)
    Dummy2 = CDbl(Temp_Text)
    If txt.Text = Temp_Text Then ValueChanged = False
    If Abs(Dummy1 - Dummy2) < TOLERANCE Then ValueChanged = False

End Sub

Sub TextStringChanged(ValueChanged As Integer, txt As TextBox, Temp_Text As String)
    
    ValueChanged = True
    If txt.Text = Temp_Text Then ValueChanged = False

End Sub

