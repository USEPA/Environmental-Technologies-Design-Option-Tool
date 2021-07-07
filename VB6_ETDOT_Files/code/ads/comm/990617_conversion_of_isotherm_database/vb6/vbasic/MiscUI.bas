Attribute VB_Name = "MiscUI"
Option Explicit




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


Sub LogOutput(f As Integer, sOut As String)
  Print #f, Now & ": " & sOut
End Sub


Sub Launch_Notepad(fn_edit As String)
Dim CmdLine As String
Dim RetVal As Integer
  CmdLine = "notepad " & fn_edit
  RetVal = 0 * Shell(CmdLine, 3)
End Sub

