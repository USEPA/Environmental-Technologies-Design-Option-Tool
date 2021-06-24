Attribute VB_Name = "MiscUI"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long




Const MiscUI_declarations_end = True


'Sub CalcStatus_Set(newVal As Boolean)
'  If (newVal) Then
'    Call GenericStatus_Set("Calculating -- please wait.")
'  Else
'    Call GenericStatus_Set("")
'  End If
'End Sub
'Sub GenericStatus_Set(fn_Text As String)
'  frmMain.sspanel_Status = fn_Text
'End Sub
'Sub DirtyStatus_Set(newVal As Boolean)
'  If (newVal) Then
'    frmMain.sspanel_Dirty = "Data Changed"
'    frmMain.sspanel_Dirty.ForeColor = QBColor(12)
'  Else
'    frmMain.sspanel_Dirty = "Unchanged"
'    frmMain.sspanel_Dirty.ForeColor = QBColor(0)
'  End If
'End Sub
'Sub DirtyStatus_Set_Current()
'  Call DirtyStatus_Set(Project_Is_Dirty)
'End Sub
'Sub DirtyStatus_Throw()
'  Project_Is_Dirty = True
'  Call DirtyStatus_Set_Current
'End Sub
'Sub DirtyStatus_Clear()
'  Project_Is_Dirty = False
'  Call DirtyStatus_Set_Current
'End Sub


Sub Global_DirtyStatus_Set( _
    Frm As Form, _
    DirtyFlag As Boolean, _
    NewSetting As Boolean)
  DirtyFlag = NewSetting
  If (NewSetting) Then
    Frm.sspanel_Dirty = "Data Changed"
    Frm.sspanel_Dirty.ForeColor = QBColor(12)
    Calculated_OK = False
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


Sub Show_Message00(msg As String, flags As Integer, WinTitle As String)
  MsgBox msg, flags, WinTitle
End Sub
Sub Show_Message0(msg As String, flags As Integer)
  Call Show_Message00(msg, vbInformation, App.Title)
End Sub
Sub Show_Message(msg As String)
  Call Show_Message0(msg, vbInformation)
End Sub
Sub Show_Error(msg As String)
  Beep
  Call Show_Message0(msg, vbExclamation)
End Sub
Sub Show_Trapped_Error(subname As String)
  Call Show_Error("An error #" & Trim$(Str$(Err)) & _
      " has occurred in routine " & Trim$(subname) & _
      ": `" & Trim$(Error$) & "`.  Ending this operation.")
End Sub


Sub Launch_Notepad(fn_edit As String)
Dim cmdline As String
Dim RetVal As Integer
  cmdline = "notepad " & fn_edit
  RetVal = 0 * Shell(cmdline, 3)
End Sub


Sub Handle_Change_Of_Units( _
    IN_NewUnitType As Integer)
Dim strNewUnits As String
  NowProj.UnitType = IN_NewUnitType
  Select Case IN_NewUnitType
    Case UnitType___ENGLISH: strNewUnits = "lbs/d"
    Case UnitType___SI: strNewUnits = "kg/d"
  End Select
  Call unitsys_set_units(frmMain.txtOutput(1), strNewUnits)
  Call unitsys_set_units(frmMain.txtOutput(2), strNewUnits)
  Call unitsys_set_units(frmMain.txtOutput(3), strNewUnits)
  Call frmMain_Refresh
End Sub

'
Sub ShellExecute_LocalFile( _
    in_Filename As String)
  Call ShellExecute(0&, vbNullString, in_Filename, vbNullString, _
vbNullString, vbNormalFocus)
End Sub
Sub ShellExecute_URL( _
    in_URL As String)
  Call ShellExecute(0&, vbNullString, in_URL, vbNullString, _
    vbNullString, vbNormalFocus)
End Sub



