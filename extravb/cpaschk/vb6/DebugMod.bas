Attribute VB_Name = "DebugMod"
Option Explicit

'Global Const DebugMode = True
Global Const DebugMode = False





Const DebugMod_declarations_end = True


Sub Debug_Output(s As String)
Dim f As Integer
  f = FreeFile
  Open "c:\bug.txt" For Append As #f
  Write #f, "cpaschk", Date$ & " " & Time$ & " -- " & s
  Close #f
End Sub

