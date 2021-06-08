Option Strict Off
Option Explicit On
Module GridFunc
	
	
	
	
	Const GridFunc_declarations_end As Boolean = True
	
	
	Sub GridFunc_GetTempFilename(ByRef out_fn_Temp As String)
		Dim fn_This As String
		Dim fn_Path As String
		Dim i As Short
		Dim Found As Boolean
		Dim This_SimCode As String
		'''''This_SimCode = NowProj.code
		'''''fn_Path = AppPath & "\sims\" & This_SimCode
		''''fn_Path = MAIN_EXE_PATH
		fn_Path = MAIN_APP_PATH & "\exes"
		Found = False
		For i = 1 To 1000
			fn_This = fn_Path & "\" & "temp" & Trim(Str(i)) & ".tmp"
			If (FileExists(fn_This) = False) Then
				Found = True
				Exit For
			End If
		Next i
		If (Found = False) Then
			Call Show_Error("Unable to create temporary file in " & "file-path `" & fn_Path & "`!  Grid data may become " & "corrupted.  Recommendation: Make a backup copy of " & "all simulation data as soon as possible.")
			Exit Sub
		End If
		'RETURN TEMP FILENAME.
		out_fn_Temp = fn_This
	End Sub
	Sub GridFunc_CopyGrid(ByRef foFrom As VCIF1Lib.F1Book, ByRef foTo As VCIF1Lib.F1Book)
		Dim F1FileFormulaOne As Object
		Dim fn_Temp As String
		Dim out_FileType As Short 'NOTE: out_FileType IS IGNORED.
		Call GridFunc_GetTempFilename(fn_Temp)
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.Write. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foFrom.Write(fn_Temp, F1FileFormulaOne) 'F1FileFormulaOne3
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.Read. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.Read(fn_Temp, out_FileType)
		Kill(fn_Temp)
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.MaxRow = foFrom.MaxRow
	End Sub
	Sub GridFunc_CopyGrid_ViaClipboard(ByRef foFrom As VCIF1Lib.F1Book, ByRef foFrom_Sheet As Short, ByRef foTo As VCIF1Lib.F1Book, ByRef foTo_Sheet As Short)
		Dim F1ClearAll As Object
		'Dim fn_Temp As String
		'Dim out_FileType As Integer     'NOTE: out_FileType IS IGNORED.
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foFrom.Sheet = foFrom_Sheet
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foFrom.SelStartRow = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foFrom.SelStartCol = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foFrom.SelEndRow = foFrom.MaxRow
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foFrom.SelEndCol = foFrom.MaxCol
		'UPGRADE_WARNING: Couldn't resolve default property of object foFrom.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foFrom.EditCopy()
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.Sheet = foTo_Sheet
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.EditClear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.EditClear(F1ClearAll)
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.SelStartRow = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.SelStartCol = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.SelEndRow = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.SelEndCol = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.EditPaste. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.EditPaste()
	End Sub


	'
	'SUBROUTINE GridFunc_Convert_CommasToTabs
	'
	'CONVERTS ALL COMMA (,) CHARACTERS TO TAB CHARACATERS (CHR#9).
	'
	Sub GridFunc_Convert_CommasToTabs_OneLine(ByRef OldLine As String, ByRef NewLine As String)
		Dim WorkingStr As String
		Dim NextStr As String
		Dim NextPos As Short
		Dim ThisIter As Short
		WorkingStr = OldLine
		ThisIter = 0
		Do While (1 = 1)
			NextPos = InStr(WorkingStr, ",")
			If (NextPos = 0) Then Exit Do
			If (NextPos > 1) Then
				NextStr = Left(WorkingStr, NextPos - 1)
			Else
				NextStr = ""
			End If
			NextStr = NextStr & Chr(9) 'tab character
			If (NextPos < Len(WorkingStr)) Then
				NextStr = NextStr & Right(WorkingStr, Len(WorkingStr) - NextPos)
			End If
			WorkingStr = NextStr
			ThisIter = ThisIter + 1
			If (ThisIter > 100) Then Exit Do
		Loop 
		NewLine = WorkingStr
	End Sub
	Sub GridFunc_Convert_CommasToTabs(ByRef fn_In As String, ByRef fn_Out As String)
		Dim f1 As Short
		Dim f2 As Short
		Dim OldLine As String
		Dim NewLine As String
		f1 = FreeFile
		FileOpen(f1, fn_In, OpenMode.Input)
		f2 = FreeFile
		FileOpen(f2, fn_Out, OpenMode.Output)
		Do While (1 = 1)
			If (EOF(f1)) Then Exit Do
			OldLine = LineInput(f1)
			Call GridFunc_Convert_CommasToTabs_OneLine(OldLine, NewLine)
			PrintLine(f2, NewLine)
		Loop 
		FileClose(f1)
		FileClose(f2)
	End Sub
	Sub GridFunc_ImportCommaDelimited(ByRef foTo As VCIF1Lib.F1Book, ByRef fn_CommaDelimited As String)
		Dim fn_Temp As String
		Dim out_FileType As Short 'IGNORED.
		Call GridFunc_GetTempFilename(fn_Temp)
		Call GridFunc_Convert_CommasToTabs(fn_CommaDelimited, fn_Temp)
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.Sheet = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.NumSheets = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.Read. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.Read(fn_Temp, out_FileType)
		Kill(fn_Temp)
	End Sub


	'
	'SUBROUTINE GridFunc_Convert_SpacesToTabs_OneLine
	'
	'CONVERTS MULTIPLE OCCURRENCES OF SPACE CHARACTERS (CHR#32)
	'TO A SINGLE TAB CHARACATER (CHR#9).
	'
	Sub GridFunc_Convert_SpacesToTabs_OneLine(ByRef OldLine As String, ByRef NewLine As String)
		Dim WorkingStr As String
		Dim Part1 As String
		Dim Part2 As String
		Dim Part2Trimmed As String
		Dim NextStr As String
		Dim NextPos As Short
		Dim ThisIter As Short
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
		WorkingStr = Trim(OldLine)
		ThisIter = 0
		Do While (1 = 1)
			NextPos = InStr(WorkingStr, Chr(32))
			If (NextPos = 0) Then Exit Do
			If (NextPos > 1) Then
				Part1 = Left(WorkingStr, NextPos - 1)
			Else
				Part1 = ""
			End If
			Part2 = Right(WorkingStr, Len(WorkingStr) - NextPos + 1)
			Part2Trimmed = Trim(Part2)
			WorkingStr = Part1 & Chr(9) & Part2Trimmed
			ThisIter = ThisIter + 1
			If (ThisIter > 1000) Then Exit Do
		Loop 
		NewLine = WorkingStr
	End Sub
	Sub GridFunc_Convert_SpacesToTabs(ByRef fn_In As String, ByRef fn_Out As String, ByRef Do_Percent_Report As Boolean)
		Dim f1 As Short
		Dim f2 As Short
		Dim OldLine As String
		Dim NewLine As String
		Dim Now_Percent As Double
		Dim ReportPercent_Interval As Short
		Dim ReportPercent_Counter As Short
		f1 = FreeFile
		FileOpen(f1, fn_In, OpenMode.Input)
		f2 = FreeFile
		FileOpen(f2, fn_Out, OpenMode.Output)
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
					''    Trim$(Str$(CInt(Now_Percent))) & "% " & _
					''    "Complete"
					'DoEvents
				End If
			End If
			'loc(f1)
			OldLine = LineInput(f1)
			Call GridFunc_Convert_SpacesToTabs_OneLine(OldLine, NewLine)
			PrintLine(f2, NewLine)
		Loop 
		FileClose(f1)
		FileClose(f2)
	End Sub
	Sub GridFunc_ImportSpaceDelimited(ByRef foTo As VCIF1Lib.F1Book, ByRef fn_SpaceDelimited As String, ByRef Do_Percent_Report As Boolean)
		Dim fn_Temp As String
		Dim out_FileType As Short 'IGNORED.
		Call GridFunc_GetTempFilename(fn_Temp)
		Call GridFunc_Convert_SpacesToTabs(fn_SpaceDelimited, fn_Temp, Do_Percent_Report)
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.Sheet = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.NumSheets = 1
		If (Do_Percent_Report) Then
			'frmExcelCurvesProgress.lblProgress(2).Caption = _
			''    "99% Complete"
			'DoEvents
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object foTo.Read. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		foTo.Read(fn_Temp, out_FileType)
		Kill(fn_Temp)
		If (Do_Percent_Report) Then
			'frmExcelCurvesProgress.lblProgress(2).Caption = ""
			'DoEvents
		End If
	End Sub
End Module