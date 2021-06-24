Attribute VB_Name = "Launch_Plugin"
Option Explicit

Global Const LaunchFunc_PPMS_Calc = 1




Const Launch_Plugin_decl_end = True


Function Launch_Plugin_Go( _
    LaunchFunc As Integer) _
    As Boolean
Dim idx_This As Integer
Dim ThisClass As String
Dim ThisPath As String
Dim objPlugin As Object
Dim out_Raise_Dirty_Flag As Boolean
Dim out_Overall_Validity As Boolean
Dim i As Integer
Dim j As Integer
  On Error GoTo err_ThisSub
  '
  ' SET UP TO MAKE THE LAUNCH.
  '
  Select Case LaunchFunc
    Case LaunchFunc_PPMS_Calc:
      ThisClass = "PPMS_Calc.Plugin_PPMS_Calc"
  End Select
  ';idx_This = All_Plugins_Lookup_Name(Name_Plugin)
  'T'hisClass = All_Plugins(idx_This).Class
  'ThisPath = All_Plugins(idx_This).Path
  '
  ' CALL THE APPROPRIATE ACTIVEX OBJECT.
  '
  Set objPlugin = CreateObject(ThisClass)
  '#If IsOutsideDebugEnv Then
  '  ''''MsgBox "About to call this class: " & ThisClass
  '  Set objPlugin = CreateObject(ThisClass)
  '#Else
  '  'Select Case Trim$(UCase$(Name_Plugin))
  '  '  Case Trim$(UCase$("DORT")):
  '  '    Set objPlugin = New plugin_dort
  '  '  Case Trim$(UCase$("EFRAT")):
  '  '    Set objPlugin = New plugin_efrat
  '  'End Select
  '#End If
  Call objPlugin.PPMS_Calc_Go( _
      MAIN_APP_PATH)
  '
  ' REFRESH WINDOW.
  '
  Call frmMain_Refresh
  '
  ' EXIT OUTTA HERE.
  '
exit_normally_ThisSub:
  Launch_Plugin_Go = True
  Exit Function
exit_err_ThisSub:
  Launch_Plugin_Go = False
  Exit Function
err_ThisSub:
  Call Show_Trapped_Error("Launch_Plugin_Go")
  Resume exit_err_ThisSub
End Function







