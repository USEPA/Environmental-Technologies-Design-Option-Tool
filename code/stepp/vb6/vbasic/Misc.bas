Attribute VB_Name = "MiscMod"
Option Explicit

Global Const STEPPLINK_STATUS_INACTIVE = 1
Global Const STEPPLINK_STATUS_ACTIVE = 2
Global SteppLink_Status As Integer

Global SteppLink_ClientProgram As String     'ASAP or ADSIM
Global SteppLink_SpecifiedPressure As String
Global SteppLink_SpecifiedTemperature As String
Global SteppLink_fn_done_waitfile As String
Global SteppLink_fn_loadup_waitfile As String
Global SteppLink_fn_properties As String

Global commandparam_numargs As Integer

Sub centerform_relative(x_parent As Form, x_child As Form)

  'Don't attempt if form is minimized or maximized
  If (x_child.WindowState = 0) Then
    x_child.Left = x_parent.Left + (x_parent.Width - x_child.Width) / 2
    x_child.Top = x_parent.Top + (x_parent.Height - x_child.Height) / 2
  End If

End Sub

'Create a temporary file in the path {use_path}.
'Returns the filename {fn_temp}.
'Note: Does not return the path of the temporary file in {fn_temp}!
Sub GetTempFilename(use_path As String, fn_temp As String)
Dim temp As String
Dim trycount As Integer
Dim I As Integer
Dim c As String
Dim nowtime As String

Dim save_path As String
Dim f As Integer

  save_path = CurDir$
  ChDir use_path
  ChDrive use_path

  nowtime = Time$
  temp = Left$(Time$, 2) + Mid$(Time$, 4, 2) + Right$(Time$, 2) + ".___"
  trycount = 0
  I = 1
  Do While (1 = 1)
    If (Dir(temp) = "") Then Exit Do
    trycount = trycount + 1
    'if (trycount > 40) then
    I = I + 1
    If (I >= 7) Then
      I = 1
    End If
    c = Mid$(temp, I, 1)
    If ((c >= "0") And (c <= "8")) Then
      Mid$(temp, I, 1) = Chr$(Asc(c) + 1)
    ElseIf ((c >= "A") And (c <= "Y")) Then
      Mid$(temp, I, 1) = Chr$(Asc(c) + 1)
    ElseIf (c = "9") Then
      Mid$(temp, I, 1) = "A"
    ElseIf (c = "Z") Then
      Mid$(temp, I, 1) = "0"
    End If
  Loop

  fn_temp = temp

  f = FreeFile
  Open fn_temp For Output As #f
  Close #f
  ChDir save_path
  ChDrive save_path

End Sub

Sub parsedargs_getarg(sepchar As String, inline As String, ArgNum As Integer, RetStr As String)
Dim I As Integer
Dim J As Integer

  RetStr = ""
  J = 1
  For I = 1 To Len(inline)
    If (Mid$(inline, I, 1) = sepchar) Then
      J = J + 1
      If (J > ArgNum) Then Exit For
    Else
      If (J = ArgNum) Then
        RetStr = RetStr + Mid$(inline, I, 1)
      End If
    End If
  Next I

End Sub

Function ParsedArgs_GetNum(sepchar As String, inline As String) As Integer
Dim NumArgs As Integer
Dim I As Integer

  NumArgs = 1     'between chr #1 and first separator char.
  For I = 1 To Len(inline)
    If (Mid$(inline, I, 1) = sepchar) Then
      NumArgs = NumArgs + 1
    End If
  Next I

  ParsedArgs_GetNum = NumArgs

End Function

Sub SteppLink_AddItemToClipboard(StrDesc As String, StrData As String, cliptext As String)
Dim vb3CrLf As String
  vb3CrLf = Chr$(13) & Chr$(10)
  cliptext = cliptext & StrDesc
  cliptext = cliptext & vb3CrLf
  cliptext = cliptext & StrData
  cliptext = cliptext & vb3CrLf
End Sub

Function SteppLink_GetPropertyForOutput(pnum As Integer) As String
Dim S As String
  S = contam_prop_form.lblContaminantProperties(pnum)
  If (Trim$(UCase$(S)) = "NOT AVAILABLE") Then
    S = "UNAVAILABLE"
  Else
    'DO NOTHING.
  End If
  SteppLink_GetPropertyForOutput = S
End Function

Sub SteppLink_OutputProperty(fp As Integer, pnum As Integer, pname As String, punits As String)
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim S As String
  s1 = pname
  s2 = punits
  S = SteppLink_GetPropertyForOutput(pnum)
  Write #fp, s1, s2, S
End Sub

