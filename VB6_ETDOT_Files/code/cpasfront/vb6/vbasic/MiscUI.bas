Attribute VB_Name = "MiscUI"
Option Explicit


Sub Show_Message0(msg As String, flags As Integer)
  MsgBox msg, flags, App.Title
End Sub
Sub Show_Message(msg As String)
  'MsgBox msg, vbCritical, App.Title
  Call Show_Message0(msg, vbInformation)
End Sub
Sub Show_Error(msg As String)
  Beep
  Call Show_Message0(msg, vbCritical)
End Sub
Sub Show_Trapped_Error(subname As String)
  Call Show_Error("An error #" & Trim$(Str$(Err)) & _
      " has occurred in routine " & Trim$(subname) & _
      ": `" & Trim$(Error$) & "`.  Ending this operation.")
End Sub


Sub StatusInfo_Display(sspanelX As Control, _
    msg As String, _
    qbc_back As Integer, _
    qbc_fore As Integer)
Dim use_msg As String
  use_msg = Trim$(msg)
  sspanelX.Caption = use_msg
  If (use_msg = "") Then
    sspanelX.BackColor = QBColor(7)
    sspanelX.ForeColor = QBColor(0)
  Else
    sspanelX.BackColor = QBColor(qbc_back)
    sspanelX.ForeColor = QBColor(qbc_fore)
  End If
End Sub


Sub Close_All_Windows()
Dim ifc%
Dim i%
  On Error Resume Next
  ifc% = Forms.Count - 1
  For i% = ifc% To 0 Step -1
    'If (Forms(i%).Name <> "frmMain") And _
    '   (Forms(i%).Name <> "frmProgress") Then
      Unload Forms(i%)
    'End If
  Next i%
End Sub


Sub CenterOnScreen(frm_to_center As Form)
  frm_to_center.Left = (Screen.Width - frm_to_center.Width) / 2
  frm_to_center.Top = (Screen.Height - frm_to_center.Height) / 2
End Sub
Sub CenterOnForm(frm_to_center As Form, frm As Form)
  frm_to_center.Left = frm.Left + (frm.Width - frm_to_center.Width) / 2
  frm_to_center.Top = frm.Top + (frm.Height - frm_to_center.Height) / 2
End Sub


'NOTE: THIS FUNCTION WORKS EQUALLY WELL ON
'EITHER FILES OR DIRECTORIES.
Function File_IsExists(fn As String) As Boolean
Dim Dummy As Long
  On Error GoTo err_File_IsExists
  Dummy = GetAttr(fn)
  File_IsExists = True
  Exit Function
exit_err_File_IsExists:
  File_IsExists = False
  Exit Function
err_File_IsExists:
  Resume exit_err_File_IsExists
End Function
Function FileExists(fn As String) As Boolean
  FileExists = File_IsExists(fn)
End Function


Sub Launch_Notepad(fn_edit As String)
Dim CmdLine As String
Dim RetVal As Integer
  CmdLine = "notepad " & fn_edit
  RetVal = 0 * Shell(CmdLine, 3)
End Sub

