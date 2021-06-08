Attribute VB_Name = "LaunchApp"
Option Explicit





Const LaunchApp_declarations_end = 0


Sub Launch_Application( _
    proj As ProjectType, _
    tab_in As Integer, _
    group_in As Integer, _
    icon_in As Integer)
Dim ta As TabType
Dim gr As GroupType
Dim ic As IconType
Dim CmdLine_0 As String
Dim CmdLine As String
Dim LaunchDir As String
Dim msg_base As String
Dim msg As String
Dim RetVal As Integer
Dim RetValBool As Boolean
  On Error GoTo err_Launch_Application_OtherErrors
  ta = proj.Tabs(tab_in)
  gr = ta.Groups(group_in)
  ic = gr.Icons(icon_in)
  'MsgBox "Testing! " & Trim$(ic.Name), vbInformation
  CmdLine_0 = ic.fn_ApplicationLink
  CmdLine = String_PrepareForApplicationLaunch(CmdLine_0)
  LaunchDir = String_PrepareForApplicationLaunch(ic.fn_ApplicationLink_Dir)
  msg_base = "Application Information:" & vbCrLf & vbCrLf & _
      "Tab Name: `" & ta.Name & "`" & vbCrLf & _
      "Group Name: `" & gr.Name & "`" & vbCrLf & _
      "Icon Name: `" & ic.Name & "`" & vbCrLf & _
      "Launch Directory: `" & LaunchDir & "`" & vbCrLf & _
      "Command Line: `" & CmdLine & "`" & vbCrLf
  If (frmMain_MODE = frmMain_MODE_DESIGN) Then
    'DESIGN MODE: DON'T ACTUALLY LAUNCH APPLICATION;
    'DISPLAY DEBUG INFO INSTEAD.
    msg = msg_base & _
        vbCrLf & _
        "Note: Switch to User Mode in order to launch applications"
    Call Show_Message(msg)
    Exit Sub
  End If
  'HIDE THE POPUP WINDOW (IF SHOWN).
  Call frmMain.Main_PopupWindow_HideIt
  'LAUNCH THE APP.
  On Error GoTo err_Launch_Application
  If (NowProj.MinimizeOnApplicationExecution) Then
    frmMain.WindowState = 1
  End If
  If (LaunchDir <> "") Then
    ChDir LaunchDir
  End If
  If (1 = 1) Then
    RetValBool = LaunchFile_General(LaunchDir, CmdLine)
    If (Not RetValBool) Then
      GoTo resume_err_Launch_Application
    End If
  Else
    RetVal = 0 * Shell(CmdLine, 1)
  End If
  On Error Resume Next
  Exit Sub
  
resume_err_Launch_Application:
  If (NowProj.MinimizeOnApplicationExecution) Then
    frmMain.WindowState = 0
  End If
  msg = "Error: Unable to launch application.  There are several " & _
      "possible reasons why this error may occur: the " & _
      "application has not been installed yet, or the drive " & _
      "it is installed on is not accessible, or your computer " & _
      "is low on available resources." & vbCrLf & vbCrLf
  msg = msg & msg_base
  Call Show_Error(msg)
  Exit Sub
err_Launch_Application:
  Resume resume_err_Launch_Application

exit_err_Launch_Application_OtherErrors:
  Exit Sub
err_Launch_Application_OtherErrors:
  Resume exit_err_Launch_Application_OtherErrors
End Sub



Function String_ConvertStr1ToStr2( _
    str_in As String, _
    str1 As String, _
    str2 As String) As String
Dim str_now As String
Dim temp As String
Dim temp2 As String
Dim idx As Integer
Dim Chars_Remaining As Integer
  str_now = str_in
  Do While (1 = 1)
    idx = InStr(str_now, str1)
    If (idx = 0) Then Exit Do
    temp = IIf(idx = 1, "", Left$(str_now, idx - 1))
    Chars_Remaining = Len(str_now) - idx - Len(str1) + 1
    temp2 = IIf(Chars_Remaining <= 0, "", Right$(str_now, Chars_Remaining))
    str_now = temp & str2 & temp2
  Loop
  String_ConvertStr1ToStr2 = str_now
End Function


Function String_PrepareForApplicationLaunch( _
    str_in As String) As String
  String_PrepareForApplicationLaunch = _
      String_ConvertStr1ToStr2(str_in, "<C>", App.Path)
End Function


