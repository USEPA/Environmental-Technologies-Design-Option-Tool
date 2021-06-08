Attribute VB_Name = "MiscMod"
Option Explicit

Global Const POSITIONFORM_CENTER = 0
Global Const POSITIONFORM_UR = 1

Sub CenterThisForm(ThisForm As Form)

  Call PositionThisForm(ThisForm, POSITIONFORM_CENTER)

End Sub

Function GetLogDateTime() As String
Dim s As String
Dim s2 As String
Dim NowDateTime

  NowDateTime = Now
  s = Format$(NowDateTime, "ddd mmm")
  s2 = Format$(NowDateTime, "d")
  If (Len(s2) = 1) Then s2 = " " & s2
  s = s & " " & s2
  s2 = Format$(NowDateTime, "h")
  If (Len(s2) = 1) Then s2 = " " & s2
  s = s & " " & s2
  s = s & ":" & Format$(NowDateTime, "nn") & ":" & Format$(NowDateTime, "ss")
  s = s & " " & Format$(NowDateTime, "yyyy")

  GetLogDateTime = s

End Function

Function IsFormLoaded(FormToCheck As Form) As Integer
    Dim Y As Integer
    
    For Y = 0 To Forms.Count - 1
        If Forms(Y) Is FormToCheck Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
    IsFormLoaded = False
End Function

Sub PositionAForm(MainForm As Form, ThisForm As Form, Pos As Integer)
Dim x As Long
Dim Y As Long
Dim CORNER_MARGIN_TWIPS As Long

  CORNER_MARGIN_TWIPS = 200

  Select Case Pos
    Case POSITIONFORM_CENTER:
      x = MainForm.Left + (MainForm.Width - ThisForm.Width) / 2
      Y = MainForm.Top + (MainForm.Height - ThisForm.Height) / 2
    Case POSITIONFORM_UR:
      x = MainForm.Left + MainForm.Width - (ThisForm.Width + CORNER_MARGIN_TWIPS)
      Y = MainForm.Top + (MainForm.Height - MainForm.ScaleHeight) + CORNER_MARGIN_TWIPS
  End Select

  ThisForm.Move x, Y

End Sub

Sub PositionThisForm(ThisForm As Form, Pos As Integer)

  ThisForm.WindowState = 0
  If IsFormLoaded(frmptadscreen1) Then
    Call PositionAForm(frmptadscreen1, ThisForm, Pos)
  ElseIf IsFormLoaded(frmPTADScreen2) Then
    Call PositionAForm(frmPTADScreen2, ThisForm, Pos)
  ElseIf IsFormLoaded(frmBubble) Then
    Call PositionAForm(frmBubble, ThisForm, Pos)
  ElseIf IsFormLoaded(frmsurface) Then
    Call PositionAForm(frmsurface, ThisForm, Pos)
  End If

End Sub

Sub system_log(info As String)
  'Call System_Log0("debug.log", info)
End Sub

Sub System_Log0(fn_log As String, info As String)

  Exit Sub

Dim f As Integer
Dim s As String
Dim fn_debug As String

  f = FreeFile
  fn_debug = App.Path & "\" & fn_log
  Open fn_debug For Append As #f
  s = GetLogDateTime()
  Print #f, s & " : " & info
  Close #f
  Kill fn_debug

End Sub

