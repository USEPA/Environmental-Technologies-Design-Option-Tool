Attribute VB_Name = "ValidateMod"
Option Explicit

Global Const vfSendFocus = 1
Global Const vfReturnFocus = 2

Global glBackColor As Long
Global glForeColor As Long





Const ValidateMod_declarations_end = True


Sub GotFocus_Handle(frm As Form, Ctl As Control, OriginalValue As String)
  If TypeOf Ctl Is TextBox Then
    '
    ' Select entire text string
    '
    OriginalValue = Ctl.Text
    Ctl.SelStart = 0
    Ctl.SelLength = Len(OriginalValue)
    '
    ' Set light blue background
    '
    glBackColor = Ctl.BackColor
    glForeColor = Ctl.ForeColor
    Ctl.BackColor = &HFFFF00
    Ctl.ForeColor = &H80000008
  End If
'  If (frmMainMenu!VFocus1.ActiveControl = 0) Then
'    frmMainMenu!VFocus1.ActiveControl = ctl.hWnd
'    If TypeOf ctl Is TextBox Then
'      '-- Select entire text string
'      OriginalValue = ctl.Text
'      ctl.SelStart = 0
'      ctl.SelLength = Len(OriginalValue)
'
'      '-- Set light blue background
'      glBackColor = ctl.BackColor
'      glForeColor = ctl.ForeColor
'      ctl.BackColor = &HFFFF00
'      ctl.ForeColor = &H80000008
'    End If
'  End If
End Sub

Sub LostFocus_Handle(frm As Form, Ctl As Control, ValidationOK As Integer)
  If TypeOf Ctl Is TextBox Then
    'Ctl.BackColor = glBackColor
    'Ctl.ForeColor = glForeColor
    'FORCE TO BLACK TEXT ON WHITE BACKGROUND.
    Ctl.BackColor = QBColor(15)
    Ctl.ForeColor = QBColor(0)
  End If
'  If (ValidationOK) Then
'    frmMainMenu!VFocus1.FocusAction = vfSendFocus
'  Else
'    frmMainMenu!VFocus1.FocusAction = vfReturnFocus
'  End If
'  frmMainMenu!VFocus1.ActiveControl = 0
'
'  If TypeOf Ctl Is TextBox Then
'    Ctl.BackColor = glBackColor
'    Ctl.ForeColor = glForeColor
'  End If
End Sub

Function LostFocus_IsEvil(frm As Form, Ctl As Control)
  LostFocus_IsEvil = False
'  If (frmMainMenu!VFocus1.ActiveControl = Ctl.hWnd) Then
'    LostFocus_IsEvil = False
'  Else
'    LostFocus_IsEvil = True
'  End If
End Function


'Sub GotFocus_Handle(frm As Form, Ctl As Control, OriginalValue As String)
'
'  If (contam_prop_form!VFocus1.ActiveControl = 0) Then
'    contam_prop_form!VFocus1.ActiveControl = Ctl.hWnd
'    If TypeOf Ctl Is TextBox Then
'      '-- Select entire text string
'      OriginalValue = Ctl.Text
'      Ctl.SelStart = 0
'      Ctl.SelLength = Len(OriginalValue)
'
'      '-- Set light blue background
'      glBackColor = Ctl.BackColor
'      glForeColor = Ctl.ForeColor
'      Ctl.BackColor = &HFFFF00
'      Ctl.ForeColor = &H80000008
'    End If
'  End If
'
'End Sub
'
'Sub LostFocus_Handle(frm As Form, Ctl As Control, ValidationOK As Integer)
'
'  If (ValidationOK) Then
'    contam_prop_form!VFocus1.FocusAction = vfSendFocus
'  Else
'    contam_prop_form!VFocus1.FocusAction = vfReturnFocus
'  End If
'  contam_prop_form!VFocus1.ActiveControl = 0
'
'  If TypeOf Ctl Is TextBox Then
'    Ctl.BackColor = glBackColor
'    Ctl.ForeColor = glForeColor
'  End If
'
'End Sub
'
'Function LostFocus_IsEvil(frm As Form, Ctl As Control)
'
'  If (contam_prop_form!VFocus1.ActiveControl = Ctl.hWnd) Then
'    LostFocus_IsEvil = False
'  Else
'    LostFocus_IsEvil = True
'  End If
'
'End Function

