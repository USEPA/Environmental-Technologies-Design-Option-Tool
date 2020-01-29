Attribute VB_Name = "MainModule"
Option Explicit

Global MAIN_APP_PATH As String


'NOTE: THIS FUNCTION WORKS EQUALLY WELL ON
'EITHER FILES OR DIRECTORIES.
Function File_IsExists(fn As String) As Boolean
Dim Dummy As Long
  On Error GoTo err_File_IsExists
  Dummy = GetAttr(fn)   'TRIGGERS ERROR IF FILE DOES NOT EXIST.
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


Sub ChangeDir_Exes()
  ChDrive MAIN_APP_PATH
  ChDir MAIN_APP_PATH & "\EXES"
End Sub
Sub ChangeDir_Main()
  ChDrive MAIN_APP_PATH
  ChDir MAIN_APP_PATH
End Sub


Sub Main()
  'SET UP MAIN APP PATH VARIABLE.
  If (File_IsExists(App.Path & "\debug_in_vb5.txt")) Then
    'FOR DEBUGGING IN THE VB5 ENVIRONMENT.
    MAIN_APP_PATH = "X:\etdot10\license\snum\cpaslib_vbtest1\vb5"
    ChDrive MAIN_APP_PATH
    ChDir MAIN_APP_PATH
  Else
    'DO NOTHING.
    MAIN_APP_PATH = App.Path
  End If
  
  Form1.Show
End Sub


