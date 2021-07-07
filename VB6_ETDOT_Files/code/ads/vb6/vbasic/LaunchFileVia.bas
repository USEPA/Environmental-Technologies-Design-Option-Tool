Attribute VB_Name = "LaunchFileVia"
Option Explicit

Global Const OSTYPE_WIN95 = 1
Global Const OSTYPE_WINNT = 2

Global Const LAUNCHFILEVIA_IS_DEBUG_MODE_ON = False






Const LaunchFileVia_declarations_end = True


'RETURNS:
'    TRUE = SUCCEEDED.
'    FALSE = FAILED.
Function LaunchFileViaStartMethod_0( _
    fn_Dir As String, _
    fn_File As String, _
    OSTYPE As Integer) As Boolean
Dim RetVal As Integer
Dim CmdLine As String
    
  On Error GoTo err_LaunchFileViaStartMethod_0
  
  If (Trim$(fn_Dir) <> "") Then
    ChDir Trim$(fn_Dir)
  End If
  Select Case OSTYPE
    Case OSTYPE_WIN95:
      CmdLine = "command.com /c start " & Trim$(fn_File)
    Case OSTYPE_WINNT:
      CmdLine = "cmd /c start " & Trim$(fn_File)
  End Select
  If (LAUNCHFILEVIA_IS_DEBUG_MODE_ON) Then
    MsgBox "CmdLine = `" & CmdLine & "`"
  End If
  RetVal = 0 * Shell(CmdLine, 1)
  
  LaunchFileViaStartMethod_0 = True
  Exit Function
    
exit_err_LaunchFileViaStartMethod_0:
  LaunchFileViaStartMethod_0 = False
  Exit Function
err_LaunchFileViaStartMethod_0:
  Resume exit_err_LaunchFileViaStartMethod_0
End Function


'RETURNS:
'    TRUE = SUCCEEDED.
'    FALSE = FAILED.
Function LaunchFileViaExecMethod( _
    fn_Dir As String, _
    fn_File As String) As Boolean
Dim RetVal As Integer
Dim CmdLine As String
    
  On Error GoTo err_LaunchFileViaExecMethod
  
  If (Trim$(fn_Dir) <> "") Then
    ChDir Trim$(fn_Dir)
    On Error Resume Next
    ChDrive Trim$(fn_Dir)
    On Error GoTo err_LaunchFileViaExecMethod
  End If
  CmdLine = Trim$(fn_File)
  If (LAUNCHFILEVIA_IS_DEBUG_MODE_ON) Then
    MsgBox "CmdLine = `" & CmdLine & "`"
  End If
  'CmdLine = Dir("*.exe")
  
  RetVal = 0 * Shell(CmdLine, 1)
  
  LaunchFileViaExecMethod = True
  Exit Function
    
exit_err_LaunchFileViaExecMethod:
  LaunchFileViaExecMethod = False
  Exit Function
err_LaunchFileViaExecMethod:
  Resume exit_err_LaunchFileViaExecMethod
End Function


'RETURNS:
'    TRUE = SUCCEEDED.
'    FALSE = FAILED.
Function LaunchFileViaStartMethod( _
    fn_Dir As String, _
    fn_File As String) As Boolean
Dim RetValBool As Boolean
  RetValBool = LaunchFileViaStartMethod_0( _
      Trim$(fn_Dir), _
      Trim$(fn_File), _
      OSTYPE_WINNT)
  If (Not RetValBool) Then
    RetValBool = LaunchFileViaStartMethod_0( _
        Trim$(fn_Dir), _
        Trim$(fn_File), _
        OSTYPE_WIN95)
  End If
  LaunchFileViaStartMethod = RetValBool
End Function


'RETURNS:
'    TRUE = SUCCEEDED.
'    FALSE = FAILED.
Function LaunchFile_General( _
    fn_Dir As String, _
    fn_File As String) As Boolean
Dim RetValBool As Boolean
  If (Right$(Trim$(UCase$(fn_File)), 4) = ".EXE") Or _
      (Right$(Trim$(UCase$(fn_File)), 4) = ".COM") Or _
      (Right$(Trim$(UCase$(fn_File)), 4) = ".BAT") Then
    RetValBool = LaunchFileViaExecMethod(fn_Dir, fn_File)
  Else
    RetValBool = LaunchFileViaStartMethod(fn_Dir, fn_File)
  End If
  LaunchFile_General = RetValBool
End Function




