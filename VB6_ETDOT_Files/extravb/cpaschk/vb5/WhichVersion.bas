Attribute VB_Name = "WhichVersion"
Option Explicit

'Global Const CPASCHK_VERSION = "1.0"
Global Const CPASCHK_VERSION = "1.1"





Const WhichVersion_declarations_end = True


Sub Do_Create_File_vX(arg_CpasDir As String)
  Select Case CPASCHK_VERSION
    Case "1.0":
      Call Do_Create_File_v10(arg_CpasDir)
    Case "1.1":
      Call Do_Create_File_v11(arg_CpasDir)
  End Select
End Sub
Sub Do_Display_All_vX(arg_CpasDir As String)
  Select Case CPASCHK_VERSION
    Case "1.0":
      Call Do_Display_All_v10(arg_CpasDir)
    Case "1.1":
      Call Do_Display_All_v11(arg_CpasDir)
  End Select
End Sub
Sub Do_Get_Info_vX(arg_CpasDir As String, arg_ProgramKey As String, arg_ResultsDir As String)
  Select Case CPASCHK_VERSION
    Case "1.0":
      Call Do_Get_Info_v10(arg_CpasDir, arg_ProgramKey, arg_ResultsDir)
    Case "1.1":
      Call Do_Get_Info_v11(arg_CpasDir, arg_ProgramKey, arg_ResultsDir)
  End Select
End Sub
