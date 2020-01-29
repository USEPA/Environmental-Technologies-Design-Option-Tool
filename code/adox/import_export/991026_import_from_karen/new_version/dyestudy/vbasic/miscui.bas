Attribute VB_Name = "MiscUI"
Option Explicit




Const MiscUI_declarations_end = True


Sub Global_DirtyStatus_Set( _
    Frm As Form, _
    DirtyFlag As Boolean, _
    NewSetting As Boolean)
  DirtyFlag = NewSetting
  If (NewSetting) Then
    Frm.sspanel_Dirty = "Data Changed"
    Frm.sspanel_Dirty.ForeColor = QBColor(12)
  Else
    Frm.sspanel_Dirty = "Unchanged"
    Frm.sspanel_Dirty.ForeColor = QBColor(0)
  End If
End Sub

Sub Global_GenericStatus_Set( _
    Frm As Form, _
    NewString As String)
  Frm.sspanel_Status = NewString
End Sub


Sub frmMain_Close_All_Windows()
Dim ifc%
Dim i%
  On Error Resume Next
  ifc% = Forms.Count - 1
  For i% = ifc% To 0 Step -1
    'If (Forms(i%).name <> "frmMain") And _
       (Forms(i%).name <> "frmProgress") Then
    If (Forms(i%).Name <> "frmMain") Then
      Unload Forms(i%)
    End If
  Next i%
End Sub


Sub CenterOnScreen(frm_to_center As Form)
  frm_to_center.Left = (Screen.Width - frm_to_center.Width) / 2
  frm_to_center.Top = (Screen.Height - frm_to_center.Height) / 2
End Sub


Sub CenterOnForm(frm_to_center As Form, Frm As Form)
  frm_to_center.Left = Frm.Left + (Frm.Width - frm_to_center.Width) / 2
  frm_to_center.Top = Frm.Top + (Frm.Height - frm_to_center.Height) / 2
End Sub


Sub Show_Message(msg As String, flags As Integer, WinTitle As String)
  MsgBox msg, vbExclamation, WinTitle
End Sub

Sub Show_Error(msg As String)
  Beep
  Call Show_Message(msg, vbExclamation, App.Title)
End Sub

Sub Show_Trapped_Error(subname As String)
  Beep
  Call Show_Message("An error #" & Trim$(Str$(Err)) & _
      " has occurred in routine " & Trim$(subname) & _
      ": `" & Trim$(Error$) & "`.  Ending this operation.", _
      vbExclamation, App.Title)
End Sub


Sub Launch_Notepad(fn_edit As String)
Dim CmdLine As String
Dim RetVal As Integer
  CmdLine = "notepad " & fn_edit
  RetVal = 0 * Shell(CmdLine, 3)
End Sub

