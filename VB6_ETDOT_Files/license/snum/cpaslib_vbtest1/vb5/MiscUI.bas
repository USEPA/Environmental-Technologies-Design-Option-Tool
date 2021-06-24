Attribute VB_Name = "MiscUI"
Option Explicit




Const MiscUI_declarations_end = True


Sub Parser_GetArg(sepchar As String, inline As String, ArgNum As Integer, retstr As String)
Dim i As Integer
Dim j As Integer
  retstr = ""
  j = 1
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      j = j + 1
      If (j > ArgNum) Then Exit For
    Else
      If (j = ArgNum) Then
        retstr = retstr + Mid$(inline, i, 1)
      End If
    End If
  Next i
End Sub


Function Parser_GetNumArgs(sepchar As String, inline As String) As Integer
Dim NumArgs As Integer
Dim i As Integer
  NumArgs = 1     'between chr #1 and first separator char.
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      NumArgs = NumArgs + 1
    End If
  Next i
  Parser_GetNumArgs = NumArgs
End Function



