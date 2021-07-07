Attribute VB_Name = "GridFunc"
Option Explicit




Const GridFunc_declarations_end = True


Sub GridFunc_GetTempFilename(out_fn_Temp As String)
Dim fn_This As String
Dim fn_Path As String
Dim i As Integer
Dim Found As Boolean
Dim This_SimCode As String
  '''''This_SimCode = NowProj.code
  '''''fn_Path = AppPath & "\sims\" & This_SimCode
  ''''fn_Path = MAIN_EXE_PATH
  fn_Path = MAIN_APP_PATH & "\exes"
  Found = False
  For i = 1 To 1000
    fn_This = fn_Path & "\" & "temp" & Trim$(Str$(i)) & ".tmp"
    If (FileExists(fn_This) = False) Then
      Found = True
      Exit For
    End If
  Next i
  If (Found = False) Then
    Call Show_Error("Unable to create temporary file in " & _
        "file-path `" & fn_Path & "`!  Grid data may become " & _
        "corrupted.  Recommendation: Make a backup copy of " & _
        "all simulation data as soon as possible.")
    Exit Sub
  End If
  'RETURN TEMP FILENAME.
  out_fn_Temp = fn_This
End Sub
Sub GridFunc_CopyGrid( _
    foFrom As Control, _
    foTo As Control)
Dim fn_Temp As String
Dim out_FileType As Integer     'NOTE: out_FileType IS IGNORED.
  Call GridFunc_GetTempFilename(fn_Temp)
  foFrom.Write fn_Temp, F1FileFormulaOne    'F1FileFormulaOne3
  foTo.Read fn_Temp, out_FileType
  Kill fn_Temp
  foTo.MaxRow = foFrom.MaxRow
End Sub
Sub GridFunc_CopyGrid_ViaClipboard( _
    foFrom As Control, _
    foFrom_Sheet As Integer, _
    foTo As Control, _
    foTo_Sheet As Integer)
'Dim fn_Temp As String
'Dim out_FileType As Integer     'NOTE: out_FileType IS IGNORED.
  foFrom.Sheet = foFrom_Sheet
  foFrom.SelStartRow = 1
  foFrom.SelStartCol = 1
  foFrom.SelEndRow = foFrom.MaxRow
  foFrom.SelEndCol = foFrom.MaxCol
  foFrom.EditCopy
  foTo.Sheet = foTo_Sheet
  foTo.EditClear F1ClearAll
  foTo.SelStartRow = 1
  foTo.SelStartCol = 1
  foTo.SelEndRow = 1
  foTo.SelEndCol = 1
  foTo.EditPaste
End Sub


'
'SUBROUTINE GridFunc_Convert_CommasToTabs
'
'CONVERTS ALL COMMA (,) CHARACTERS TO TAB CHARACATERS (CHR#9).
'
Sub GridFunc_Convert_CommasToTabs_OneLine( _
    OldLine As String, _
    NewLine As String _
    )
Dim WorkingStr As String
Dim NextStr As String
Dim NextPos As Integer
Dim ThisIter As Integer
  WorkingStr = OldLine
  ThisIter = 0
  Do While (1 = 1)
    NextPos = InStr(WorkingStr, ",")
    If (NextPos = 0) Then Exit Do
    If (NextPos > 1) Then
      NextStr = Left$(WorkingStr, NextPos - 1)
    Else
      NextStr = ""
    End If
    NextStr = NextStr & Chr$(9)     'tab character
    If (NextPos < Len(WorkingStr)) Then
      NextStr = NextStr & Right$(WorkingStr, Len(WorkingStr) - NextPos)
    End If
    WorkingStr = NextStr
    ThisIter = ThisIter + 1
    If (ThisIter > 100) Then Exit Do
  Loop
  NewLine = WorkingStr
End Sub
Sub GridFunc_Convert_CommasToTabs( _
  fn_In As String, _
  fn_Out As String)
Dim f1 As Integer
Dim f2 As Integer
Dim OldLine As String
Dim NewLine As String
  f1 = FreeFile
  Open fn_In For Input As #f1
  f2 = FreeFile
  Open fn_Out For Output As #f2
  Do While (1 = 1)
    If (EOF(f1)) Then Exit Do
    Line Input #f1, OldLine
    Call GridFunc_Convert_CommasToTabs_OneLine(OldLine, NewLine)
    Print #f2, NewLine
  Loop
  Close #f1
  Close #f2
End Sub
Sub GridFunc_ImportCommaDelimited( _
    foTo As Control, _
    fn_CommaDelimited As String)
Dim fn_Temp As String
Dim out_FileType As Integer     'IGNORED.
  Call GridFunc_GetTempFilename(fn_Temp)
  Call GridFunc_Convert_CommasToTabs(fn_CommaDelimited, fn_Temp)
  foTo.Sheet = 1
  foTo.NumSheets = 1
  foTo.Read fn_Temp, out_FileType
  Kill fn_Temp
End Sub


'
'SUBROUTINE GridFunc_Convert_SpacesToTabs_OneLine
'
'CONVERTS MULTIPLE OCCURRENCES OF SPACE CHARACTERS (CHR#32)
'TO A SINGLE TAB CHARACATER (CHR#9).
'
Sub GridFunc_Convert_SpacesToTabs_OneLine( _
    OldLine As String, _
    NewLine As String _
    )
Dim WorkingStr As String
Dim Part1 As String
Dim Part2 As String
Dim Part2Trimmed As String
Dim NextStr As String
Dim NextPos As Integer
Dim ThisIter As Integer
'123456789012345678901234567890123456789012345678901234567890
'    1.2345e+00      1.2345e+00    1.2345e+00    1.2345e+00
'
'1234567
'AB  EFG
'
'len(w) = 7
'np = 3
'7-3 = 4
'
  WorkingStr = Trim$(OldLine)
  ThisIter = 0
  Do While (1 = 1)
    NextPos = InStr(WorkingStr, Chr$(32))
    If (NextPos = 0) Then Exit Do
    If (NextPos > 1) Then
      Part1 = Left$(WorkingStr, NextPos - 1)
    Else
      Part1 = ""
    End If
    Part2 = Right$(WorkingStr, Len(WorkingStr) - NextPos + 1)
    Part2Trimmed = Trim$(Part2)
    WorkingStr = Part1 & Chr$(9) & Part2Trimmed
    ThisIter = ThisIter + 1
    If (ThisIter > 1000) Then Exit Do
  Loop
  NewLine = WorkingStr
End Sub
Sub GridFunc_Convert_SpacesToTabs( _
  fn_In As String, _
  fn_Out As String, _
  Do_Percent_Report As Boolean)
Dim f1 As Integer
Dim f2 As Integer
Dim OldLine As String
Dim NewLine As String
Dim Now_Percent As Double
Dim ReportPercent_Interval As Integer
Dim ReportPercent_Counter As Integer
  f1 = FreeFile
  Open fn_In For Input As #f1
  f2 = FreeFile
  Open fn_Out For Output As #f2
  ReportPercent_Interval = 100
  ReportPercent_Counter = 0
  Do While (1 = 1)
    If (EOF(f1)) Then Exit Do
    If (Do_Percent_Report) Then
      ReportPercent_Counter = ReportPercent_Counter + 1
      If (ReportPercent_Counter >= ReportPercent_Interval) Then
        ReportPercent_Counter = 0
        Now_Percent = 100# * CDbl(Loc(f1)) * 128# / CDbl(LOF(f1))
        'frmExcelCurvesProgress.lblProgress(2).Caption = _
        '    Trim$(Str$(CInt(Now_Percent))) & "% " & _
        '    "Complete"
        'DoEvents
      End If
    End If
    'loc(f1)
    Line Input #f1, OldLine
    Call GridFunc_Convert_SpacesToTabs_OneLine(OldLine, NewLine)
    Print #f2, NewLine
  Loop
  Close #f1
  Close #f2
End Sub
Sub GridFunc_ImportSpaceDelimited( _
    foTo As Control, _
    fn_SpaceDelimited As String, _
    Do_Percent_Report As Boolean)
Dim fn_Temp As String
Dim out_FileType As Integer     'IGNORED.
  Call GridFunc_GetTempFilename(fn_Temp)
  Call GridFunc_Convert_SpacesToTabs( _
      fn_SpaceDelimited, _
      fn_Temp, _
      Do_Percent_Report)
  foTo.Sheet = 1
  foTo.NumSheets = 1
  If (Do_Percent_Report) Then
    'frmExcelCurvesProgress.lblProgress(2).Caption = _
    '    "99% Complete"
    'DoEvents
  End If
  foTo.Read fn_Temp, out_FileType
  Kill fn_Temp
  If (Do_Percent_Report) Then
    'frmExcelCurvesProgress.lblProgress(2).Caption = ""
    'DoEvents
  End If
End Sub



