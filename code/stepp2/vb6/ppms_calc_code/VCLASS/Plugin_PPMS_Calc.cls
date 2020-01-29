VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin_PPMS_Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit






Const Plugin_PPMS_Calc_declarations_end = True


Private Sub Class_Initialize()
'
' INITIALIZATION CODE.
'
End Sub
Private Sub Class_Terminate()
'
' TERMINATION CODE.
'
End Sub


Public Function PPMS_Calc_Go( _
    fpath_dir_StartIn As String) _
    As Boolean
  On Error GoTo err_ThisSub
  '
  ' STORE THE START-IN PATH.
  '
  Global_fpath_dir_StartIn = fpath_dir_StartIn
  '
  ' START THE PROGRAM.
  '
  Call Main
  '
  ' EXIT OUT OF HERE.
  '
  GoTo exit_normally_ThisSub
exit_normally_ThisSub:
  PPMS_Calc_Go = True
  Exit Function
exit_err_ThisSub:
  PPMS_Calc_Go = False
  Exit Function
err_ThisSub:
  Call Show_Trapped_Error("PPMS_Calc_Go")
  Resume exit_err_ThisSub
End Function


