Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmVarConcentrations
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer


	Dim Y1, X1, Shifting, X2, Y2 As Short
	Dim TempStr, Filename_Concentration As String
	Dim saveas As Short
	Dim Temp_Array() As String
	
	Dim UserWantsCancel As Short
	
	
	
	Const frmVarConcentrations_declarations_end As Boolean = True
	
	
	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdOK.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdOK.Enabled = False
		End If
	End Sub
	
	
	Private Sub ClearGrid()
		Dim F1ClearAll As Object
		Dim i As Short
		Dim sserror As Short
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		For i = 1 To Sheet1.MaxCol
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.ClearRange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		Sheet1.ClearRange(1, i, 400, i, F1ClearAll)
		'sserror = SSDeleteRange(Sheet1.ss, 1, i, 500, i, 3)
		'		Next i
		'sserror = SSDeleteTable(sheet1.SS)
		'sserror = SSDeleteRange(sheet1.SS, 1, 1, 500, 500, 1)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	
	Private Function CountConc(ByRef i As Short, ByRef npoints As Short) As Short
		Dim currentCol = i - 1
		Dim currentRow = 0
		On Error GoTo Error_In_CountConc
		npoints = 0

		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		Sheet1DataGrid.RowCount = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Do Until Sheet1DataGrid.Rows(currentRow).Cells(currentCol).Value = "" Or currentRow = Number_Max_Influent_Points
			npoints = npoints + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			currentRow += 1
		Loop
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Sheet1DataGrid.Rows(currentRow).Cells(currentCol).value <> "" Then npoints = npoints + 1
		CountConc = True
		Exit Function
Error_In_CountConc: 
		CountConc = False
		Call Show_Error("Invalid data.")
		Resume Exit_CountConc
Exit_CountConc: 
	End Function
	
	
	Private Function Load_Concentrations(ByRef OverrideFilename As String) As Boolean

		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNFileMustExist As Object
		Dim i, f, npoints, J As Short
		'UPGRADE_WARNING: Lower bound of array T was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim T(Number_Max_Influent_Points) As Double
		'UPGRADE_WARNING: Lower bound of array C was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim C(Number_Compo_Max, Number_Max_Influent_Points) As Double
		Load_Concentrations = False
		On Error GoTo Error_In_Reading
		If (OverrideFilename = "") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.CancelError = True
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.DialogTitle = "Load Concentrations"
			OpenFileDialog1.Title = "Load Concentrations"
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel 4.0 (*.xls)|*.xls"
			OpenFileDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel File (*.csv)|*.csv"
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.FilterIndex = 2
			OpenFileDialog1.FilterIndex = 3
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNFileMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.flags = cdlOFNFileMustExist + cdlOFNPathMustExist
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.Action = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'If CMDialog1.Filename = "" Then
			'Exit Function
			'End If
			OpenFileDialog1.ShowDialog()
			If OpenFileDialog1.FileName = "" Then
				Exit Function
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'Filename_Concentration = CMDialog1.Filename
			Filename_Concentration = OpenFileDialog1.FileName

		Else
			Filename_Concentration = OverrideFilename
		End If
		''''mnuFileItem(2).Enabled = True
		If VB.Right(Filename_Concentration, 3) = "XLS" Then
			'OPEN EXCEL FORMAT.
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.ReadFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'			Sheet1.ReadFile = Filename_Concentration
		Else
			'OPEN TEXT FORMAT.
			f = FreeFile
			FileOpen(f, Filename_Concentration, OpenMode.Input)
			Input(f, npoints)
			For i = 1 To npoints
				Select Case Number_Component
					Case 1
						Input(f, T(i))
						Input(f, C(1, i))
					Case 2
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
					Case 3
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
						Input(f, C(3, i))
					Case 4
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
						Input(f, C(3, i))
						Input(f, C(4, i))
					Case 5
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
						Input(f, C(3, i))
						Input(f, C(4, i))
						Input(f, C(5, i))
					Case 6
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
						Input(f, C(3, i))
						Input(f, C(4, i))
						Input(f, C(5, i))
						Input(f, C(6, i))
					Case 7
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
						Input(f, C(3, i))
						Input(f, C(4, i))
						Input(f, C(5, i))
						Input(f, C(6, i))
						Input(f, C(7, i))
					Case 8
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
						Input(f, C(3, i))
						Input(f, C(4, i))
						Input(f, C(5, i))
						Input(f, C(6, i))
						Input(f, C(7, i))
						Input(f, C(8, i))
					Case 9
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
						Input(f, C(3, i))
						Input(f, C(4, i))
						Input(f, C(5, i))
						Input(f, C(6, i))
						Input(f, C(7, i))
						Input(f, C(8, i))
						Input(f, C(9, i))
					Case 10
						Input(f, T(i))
						Input(f, C(1, i))
						Input(f, C(2, i))
						Input(f, C(3, i))
						Input(f, C(4, i))
						Input(f, C(5, i))
						Input(f, C(6, i))
						Input(f, C(7, i))
						Input(f, C(8, i))
						Input(f, C(9, i))
						Input(f, C(10, i))
				End Select
			Next i
			FileClose((f))
			'		Sheet1DataGrid.RowCount = 1
			'		Sheet1DataGrid.ColumnCount = 1
			For i = 1 To npoints
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Sheet1DataGrid.Rows(i - 1).Cells(0).Value = T(i).ToString
				'	Sheet1DataGrid.Text = T(i)
				'	Sheet1DataGrid.RowCount = Sheet1DataGrid.RowCount + 1
			Next i
			For J = 1 To Number_Component
				For i = 1 To npoints
					'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Sheet1DataGrid.Rows(i - 1).Cells(J).Value = C(J, i).ToString
				Next i
			Next J
		End If
		Load_Concentrations = True
		Exit Function
Error_In_Reading:
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = 53) Then
			'DO NOTHING.
		Else
			Call Show_Trapped_Error("Load_Concentrations")
		End If
		FileClose(f)
		Resume Exit_Load_Points
Exit_Load_Points: 
	End Function
	Private Function SaveConcentrations() As Short
		Dim cdlCancel As Object
		Dim F1FileExcel4 As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNOverwritePrompt As Object
		Dim i, f, npoints, J As Short
		Dim Stemp, temp As String
		Dim Error_Code As Short
		Dim PreviousFilename_Concentration, temporaryname As String
		On Error GoTo Error_In_SaveConcentrations
		If Not (CountConc(1, npoints)) Then
			SaveConcentrations = False
			Call Show_Error("Invalid data.  No data has been saved.")
			Exit Function
		End If
		If (Trim(Filename_Concentration) <> "") And Not (saveas) Then GoTo Save_File
		PreviousFilename_Concentration = Filename_Concentration
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.CancelError = True
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.DialogTitle = "Save Concentrations"
		SaveFileDialog1.Title = "Save Concentrations"
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel 4.0 (*.xls)|*.xls"
		SaveFileDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel File (*.csv)|*.csv"
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FilterIndex = 2
		SaveFileDialog1.FilterIndex = 3
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNOverwritePrompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Action = 2
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'temporaryname = CMDialog1.Filename

		SaveFileDialog1.ShowDialog()
		temporaryname = SaveFileDialog1.FileName
		Filename_Concentration = temporaryname

		'If IsValidPath(temporaryname, "C:") And CMDialog1.Filename <> "" Then
		'  temporaryname = Mid$(temporaryname, 1, Len(temporaryname) - 1)
		'  Filename_Concentration = temporaryname
		'Else
		'  Filename_Concentration = PreviousFilename_Concentration
		'  CMDialog1.Filename = ""
		'  MsgBox "No data has been saved.", 64, AppName_For_Display_long
		'  Exit Function
		'End If
Save_File: 
		If VB.Right(Filename_Concentration, 4) = ".XLS" Then
			'EXCEL FORMAT.
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Write. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		Sheet1.Write(Filename_Concentration, F1FileExcel4)
			'Sheet1.WriteFile = Filename_Concentration
		Else
			'TEXT FORMAT.
			mnuFileItem(2).Enabled = True
			f = FreeFile
			'Sheet1.Col = 1
			'Sheet1.Row = 1
			FileOpen(f, Filename_Concentration, OpenMode.Output)
			PrintLine(f, VB6.Format(npoints, "0"))
			For i = 1 To npoints
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Stemp = VB6.Format(CDbl(Sheet1DataGrid.Rows(i - 1).Cells(0).Value), "0.0000E+00")
				For J = 1 To Number_Component
					'Sheet1.Col = Sheet1.Col + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Stemp = Stemp & "," & VB6.Format(CDbl(Sheet1DataGrid.Rows(i - 1).Cells(J).Value), "0.0000E+00")
				Next J
				PrintLine(f, Stemp)
				'Sheet1.Row = Sheet1.Row + 1
				'Sheet1.Col = 1
			Next i
			FileClose((f))
			SaveConcentrations = True
		End If
		Exit Function
Error_In_SaveConcentrations: 
		SaveConcentrations = False
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = 75) Then
			'DO NOTHING.
		Else
			Call Show_Trapped_Error("SaveConcentrations")
		End If
		If Err.Number = 13 Then
			Call Show_Error("The data entered are not valid data.")
		End If
		FileClose(f)
		Resume Exit_Save_Points
Exit_Save_Points: 
	End Function
	
	
	Private Sub cmdCancel_Click()
		frmConcentrations_cancelled = True
		Me.Close()
	End Sub
	Private Sub cmdOK_Click()
		Dim i, response As Short
		Dim currentRow As Short
		Dim currentCol As Short
		'UPGRADE_WARNING: Lower bound of array ndata was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim ndata(24) As Short
		Dim DFlag, f As Short
		Dim J, No_Var_Influent As Short
		No_Var_Influent = False
		If Not (CountConc(1, frmConcentrations_NumPoints)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		Sheet1.SetFocus()
			Exit Sub
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		Sheet1DataGrid.RowCount = 1
		For i = 1 To frmConcentrations_NumConcs
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'			Sheet1DataGrid.ColumnCount = i
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Sheet1DataGrid.Rows(0).Cells(i).Value = "" Then No_Var_Influent = True
		Next i
		If (No_Var_Influent) Then
			response = MsgBox("There is no data for the first row." & vbCrLf & "It will be assumed that there is no concentration data.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKCancel, AppName_For_Display_Long)
			Select Case response
				Case MsgBoxResult.OK
					GoTo NoInfluent_Conc
				Case MsgBoxResult.Cancel
					'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'			Sheet1.SetFocus()
					Exit Sub
			End Select
		End If

		currentCol = 0
		currentRow = 0
		For J = 1 To frmConcentrations_NumConcs + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'			Sheet1DataGrid.ColumnCount = J
			ndata(J) = 0
			currentCol = J - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'			Sheet1DataGrid.RowCount = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			currentRow = 0
			Do Until ((Sheet1DataGrid.Rows(currentRow).Cells(currentCol).Value = "") Or (currentRow >= Number_Max_Influent_Points))
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				Sheet1DataGrid.RowCount = Sheet1DataGrid.RowCount + 1

				ndata(J) = ndata(J) + 1
				currentRow += 1
			Loop
		Next J
		DFlag = False
		For i = 1 To frmConcentrations_NumConcs + 1
			For J = i + 1 To frmConcentrations_NumConcs + 1
				If (ndata(i) <> ndata(J)) Then DFlag = True
			Next J
		Next i
		If (DFlag) Then
			response = MsgBox("There is not the same number of data in each column." & vbCrLf & "It will be assumed that there is no concentration data.", MsgBoxStyle.Exclamation + MsgBoxStyle.OKCancel, AppName_For_Display_Short)
			Select Case response
				Case MsgBoxResult.OK
				Case MsgBoxResult.Cancel
					'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'	Sheet1.SetFocus()
					Exit Sub
			End Select
		End If
		'Store times
		On Error GoTo Time_Error
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Sheet1DataGrid.ColumnCount = 1
		For i = 1 To frmConcentrations_NumPoints
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	Sheet1DataGrid.RowCount = i
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmConcentrations_Times(i) = CDbl(Sheet1DataGrid.Rows(i - 1).Cells(0).Value) * 24.0# * 60.0# 'To convert from days to minutes
			If (i > 1) Then
				If (frmConcentrations_TimeOrderImportant) Then
					If (frmConcentrations_Times(i) <= frmConcentrations_Times(i - 1)) Then GoTo Time_Error2
				End If
			End If
		Next i
		'Store concentrations
		On Error GoTo Conc_Error
		For J = 2 To frmConcentrations_NumConcs + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For i = 1 To frmConcentrations_NumPoints
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmConcentrations_Concs(J - 1, i) = CDbl(Sheet1DataGrid.Rows(i - 1).Cells(J - 1).Value)
			Next i
		Next J
		Me.Close()
Exit_This_OK: 
		frmConcentrations_cancelled = False
		Exit Sub
Time_Error:
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call Show_Error("At least one value in time input (Row #" & VB6.Format(i, "0") & ") is not a real number." & vbCrLf & "Change this cell (currently `" & Sheet1DataGrid.Rows(i).Cells(1).Value & "`) to a number.")
		Resume Exit_This_OK
Time_Error2: 
		Call Show_Error("Time in row #" & VB6.Format(i, "0") & " is less than time in row #" & VB6.Format(i - 1, "0") & "." & vbCrLf & "Change your times to be in chronological order.")
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	Sheet1.SetFocus()
		Exit Sub
Conc_Error:
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call Show_Error("At least one value of concentration (Row# " & VB6.Format(i, "0") & ", Col#" & VB6.Format(J, "0") & ") is not a real number." & vbCrLf & "Change this cell (currently `" & Sheet1DataGrid.Rows(i).Cells(J).Value & "`) to a number.")
		Resume Exit_This_OK
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	Sheet1.SetFocus()
NoInfluent_Conc: 
		Number_Influent_Points = 0
		Me.Close()
	End Sub
	
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub

	Private Sub cmdOK_ClickEvent(sender As Object, e As EventArgs)

	End Sub

	Private Sub cmdCancel_ClickEvent(sender As Object, e As EventArgs)

	End Sub

	Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
		cmdCancel_Click()
	End Sub

	Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
		Call cmdOK_Click()
	End Sub

	Private Sub Sheet1DataGrid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Sheet1DataGrid.CellContentClick

	End Sub

	'UPGRADE_WARNING: Form event frmVarConcentrations.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmVarConcentrations_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Sheet1DataGrid.Refresh()
	End Sub
	Private Sub frmVarConcentrations_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)


		Dim i, J As Short
		Dim TB, CB As String
		Dim temp, LF As String
		Dim C, SetWidth As Short
		'-- Startup watch-for-cancel timer
		frmConcentrations_cancelled = True
		UserWantsCancel = False
		''''Timer1.Enabled = True
		Me.Text = frmConcentrations_caption
		'-- Initialize last-few-files list for this form
		'xaxaxaNC
		''Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADSIM, LASTFEW_ADSIM_FRMCONCE_FRM)
		'Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADXDESIGNS, LASTFEW_ADXDESIGNS_FRMCONCE_FRM)
		'
		'POPULATE LAST-FEW-FILES LIST.
		Call OldFileList_Populate(2, Me._mnuFileItem_190, Me.mnuFileItem(191), Me.mnuFileItem(192), Me.mnuFileItem(193), Me.mnuFileItem(194))
		''Me.HelpContextID = Hlp_Influent_Concentrations
		'mnuEditItem(0) = True
		'mnuEditItem(1) = True
		''mnuEditItem(2) = False       '--- DON'T DO THIS!
		TB = Chr(9)
		CB = Chr(13)
		'set maxcols
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Sheet1DataGrid.ColumnCount = Number_Component + 1
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Sheet1DataGrid.RowCount = 400
		'set col headers
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.ColText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Sheet1DataGrid.Columns(0).HeaderText = "Time (" & frmConcentrations_Tunits & ")"
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.ColWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	Sheet1DataGrid.Columns(0).Width = 12 * 256
		'sserror = SSSetColText(Sheet1.ss, 1, "Time (" & frmConcentrations_Tunits & ")")
		'sserror = SSSetColWidth(Sheet1.ss, 1, 1, (12 * 256), False)
		For i = 1 To Number_Component
			Label1(i - 1).Visible = True
			Label2(i - 1).Visible = True
			Label2(i - 1).Text = Trim(Component(i).Name) & " (" & frmConcentrations_Cunits & ")"
			'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.ColText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Sheet1DataGrid.Columns(i).HeaderText = Chr((Asc("A")) + i - 1)
			'sserror = SSSetColText(Sheet1.ss, 1 + i, Chr$((Asc("A")) + i - 1))
			''sserror = SSSetColWidth(sheet1.SS, 1 + i, 1 + i, ((Len(LTrim(RTrim(temp))) + 3) * 256), False)
		Next i
		' size for amount of chemicals
		'	If (Number_Component < 5) Then
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	Sheet1DataGrid.Top = 975 + (Number_Component * 255)
		'	Else
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	Sheet1DataGrid.Top = 975 + (5 * 255)
		'	End If
		'For I = 0 To Number_Component + 1
		' grid1.FixedAlignment(I) = 2
		'Next I
		SetWidth = VB6.TwipsPerPixelX * 19
		If (frmConcentrations_NumPoints > 0) Then
			For i = 1 To frmConcentrations_NumPoints
				'CONVERT FROM minutes TO days.
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

				'		Sheet1DataGrid.RowCount = i
				'		Sheet1DataGrid.ColumnCount = 1
				Sheet1DataGrid.Rows(i - 1).Cells(1 - 1).Value = (frmConcentrations_Times(i) / 60.0# / 24.0#).ToString

				'Sheet1.number = frmConcentrations_Times(i) / 60# / 24#       'Convert form min. to days
				For J = 2 To Number_Component + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'			Sheet1DataGrid.ColumnCount = J
					Sheet1DataGrid.Rows(i - 1).Cells(J - 1).Value = frmConcentrations_Concs(J - 1, i).ToString
					'Sheet1.Col = J
					'Sheet1.number = frmConcentrations_Concs(J - 1, i)
				Next J
			Next i
		End If
		Sheet1DataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill  'shang
		Sheet1DataGrid.CurrentCell = Sheet1DataGrid.Rows(0).Cells(0)

		''''Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmConcentrations.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmConcentrations.Height / 2)
		Call CenterOnForm(Me, frmMain)
		'
		' DEMO SETTINGS.
		'
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	'UPGRADE_WARNING: Event frmVarConcentrations.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmVarConcentrations_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize

		rs.ResizeAllControls(Me)
		Dim XXX As Integer
		Dim USE_MARGIN As Integer
		If (Me.WindowState = 1) Then
			'CANNOT RESIZE WHEN MINIMIZED; EXIT OUTTA HERE.
			Exit Sub
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		USE_MARGIN = Sheet1DataGrid.Left
		XXX = VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - 2 * USE_MARGIN
		If (XXX < 1000) Then XXX = 1000
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		Sheet1DataGrid.Width = XXX
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XXX = VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - Sheet1DataGrid.Top + USE_MARGIN
		If (XXX < 1000) Then XXX = 1000
		'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		Sheet1DataGrid.Height = XXX
		'If WindowState <> 0 Then Exit Sub
		'If (Grid1.Left + Grid1.Width) > (cdmEdit(3).Left + cdmEdit(3).Width) Then
		'  Width = Grid1.Left + Grid1.Width + 20 * Screen.TwipsPerPixelX
		'Else
		'  Width = cdmEdit(3).Left + cdmEdit(3).Width + 20 * Screen.TwipsPerPixelX
		'End If
		'If Height > (Grid1.Top + 90 * Screen.TwipsPerPixelY + cmdCancel.Height) Then
		' Grid1.Height = Height - Grid1.Top - 90 * Screen.TwipsPerPixelY - cmdCancel.Height
		' cmdCancel.Top = Grid1.Top + Grid1.Height + 15 * Screen.TwipsPerPixelY
		' cmdOK.Top = Grid1.Top + Grid1.Height + 15 * Screen.TwipsPerPixelY
		'End If
		'Top = Screen.Height / 2 - Height / 2
		'Left = Screen.Width / 2 - Width / 2
		Sheet1DataGrid.Focus()
	End Sub
	
	
	Public Sub mnuEditItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditItem.Click
		Dim Index As Short = mnuEditItem.GetIndex(eventSender)
		Dim selectedCells As DataGridViewSelectedCellCollection
		Select Case Index
			Case 0 'CUT.
				Dim oCell As DataGridViewCell
				Dim oRow As DataGridViewRow

				Dim dataObj As DataObject = Me.Sheet1DataGrid.GetClipboardContent()

				If (dataObj IsNot Nothing) Then
					Clipboard.SetDataObject(dataObj)
				End If

				selectedCells = Sheet1DataGrid.SelectedCells
				For Each oCell In selectedCells
					oCell.Value = ""
				Next
				For Each oRow In Sheet1DataGrid.SelectedRows
					Sheet1DataGrid.Rows.Remove(oRow)
				Next

				On Error Resume Next
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EditCut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				Sheet1.EditCut()

			Case 1 'COPY.
				On Error Resume Next
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				  
				Dim dataObj As DataObject = Me.Sheet1DataGrid.GetClipboardContent()

				If (dataObj IsNot Nothing) Then
					Clipboard.SetDataObject(dataObj)
				End If

			Case 2 'PASTE.
				Dim datastring As String = Clipboard.GetText()
				Dim lines() As String = datastring.Split(New String() {Environment.NewLine, "\n", "\r", "\r\n"}, StringSplitOptions.RemoveEmptyEntries)
				Dim iRow As Integer = Sheet1DataGrid.CurrentCell.RowIndex
				Dim iCol As Integer = Sheet1DataGrid.CurrentCell.ColumnIndex
				Dim startCol As Integer = iCol
				Dim oCell As DataGridViewCell
				Dim line As String
				Dim cellvalue As String
				Dim ind As Integer
				For Each line In lines
					Dim cells() As String = line.Split(New String() {vbTab}, StringSplitOptions.RemoveEmptyEntries)
					If iRow > Sheet1DataGrid.RowCount - 1 Then
						Exit For
					End If
					iCol = startCol
					For Each cellvalue In cells
						If iCol > Sheet1DataGrid.ColumnCount - 1 Then
							Exit For
						End If
						oCell = Sheet1DataGrid.Rows(iRow).Cells(iCol)
						oCell.Value = cellvalue
						iCol += 1
					Next
					iRow += 1
				Next

				On Error Resume Next
				'UPGRADE_WARNING: Couldn't resolve default property of object Sheet1.EditPaste. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				Sheet1.EditPaste()
		End Select
		'
		'    Case 0  'cut
		'      sserror = SSEditCut(Sheet1.ss)
		'           mnuEditItem(2) = True
		'           Sheet1.SetFocus
		'           If sserror <> 0 Then
		'              'oops
		'           End If
		'     Case 1  'copy
		'           sserror = SSEditCopy(Sheet1.ss)
		'           mnuEditItem(2) = True
		'           If sserror <> 0 Then
		'              'oops
		'           End If
		'
		'     Case 2 'paste
		'           If (CutString()) Then
		'           Else
		'             MsgBox "Impossible to paste data from the clipboard.", 64, AppName_For_Display_long
		'           End If
		'           'sserror = SSEditpastevalues(sheet1.SS)
		'           'sheet1.SetFocus
		'           'If sserror <> 0 Then
		'           '   'oops
		'           'End If
	End Sub
	
	
	Public Sub mnuFileItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileItem.Click
		Dim Index As Short = mnuFileItem.GetIndex(eventSender)
		Dim J, i, f As Short
		Dim response As Short
		Dim fn_new As String
		Select Case Index
			Case 0 'new
				'save changes?
				response = MsgBox("Do you wish to save the changes?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNoCancel, AppName_For_Display_Long)
				If response = MsgBoxResult.Cancel Then Exit Sub
				If response = MsgBoxResult.Yes Then
					saveas = False
					f = SaveConcentrations()
				End If
				'-- clear grid
				Call ClearGrid()
				
				'   screen.MousePointer = 11
				'   For i = 1 To sheet1.maxcol
				'         sheet1.Col = i
				'      For j = 1 To sheet1.MaxRow
				'          sheet1.Row = j
				'          sheet1.Text = ""
				'      Next
				'   Next
				'
				'   screen.MousePointer = 0
			Case 1 'open
				'save changes?
				response = MsgBox("Do you wish to save the changes?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNoCancel, AppName_For_Display_Long)
				If response = MsgBoxResult.Cancel Then Exit Sub
				If response = MsgBoxResult.Yes Then
					saveas = False
					f = SaveConcentrations()
				End If
				Call Load_Concentrations("")
				''''cd_HomeDir
				'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
				Call OldFileList_Promote(Filename_Concentration, 2, Me._mnuFileItem_190, Me.mnuFileItem(191), Me.mnuFileItem(192), Me.mnuFileItem(193), Me.mnuFileItem(194))
				''''Call LastFewFiles_MoveFilenameToTop(Filename_Concentration)
			Case 2 'save
				saveas = False
				f = SaveConcentrations()
			Case 3 'saveas
				saveas = True
				f = SaveConcentrations()
				'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
				Call OldFileList_Promote(Filename_Concentration, 2, Me._mnuFileItem_190, Me.mnuFileItem(191), Me.mnuFileItem(192), Me.mnuFileItem(193), Me.mnuFileItem(194))
				''''Call LastFewFiles_MoveFilenameToTop(Filename_Concentration)
		End Select
		
		'---- Handle last-few-files stuff
		If ((Index >= 191) And (Index <= 194)) Then
			'---- save first?
			response = MsgBox("Do you wish to save the changes?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNoCancel, AppName_For_Display_Long)
			If response = MsgBoxResult.Cancel Then Exit Sub
			If response = MsgBoxResult.Yes Then
				saveas = False
				f = SaveConcentrations()
			End If
			'---- clear grid
			Call ClearGrid()
			'---- start open
			fn_new = mnuFileItem(Index).Text
			fn_new = VB.Right(fn_new, Len(fn_new) - 5)
			If (Load_Concentrations(fn_new) = False) Then
				'DO NOTHING -- FILE NOT LOADED.
			Else
				mnuFileItem(2).Enabled = True
				'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
				Call OldFileList_Promote(Filename_Concentration, 2, Me._mnuFileItem_190, Me.mnuFileItem(191), Me.mnuFileItem(192), Me.mnuFileItem(193), Me.mnuFileItem(194))
				''''Call LastFewFiles_MoveFilenameToTop(Filename_Concentration)
			End If
		End If
	End Sub
	
	
	'Private Function PasteString(StringToTransfer As String, Row As Integer, Col As Integer) As Integer
	'On Error GoTo Error_In_PasteString
	'  Sheet1.Row = Row
	'  Sheet1.Col = Col
	'  Sheet1.Text = StringToTransfer
	'  PasteString = True
	'  Exit Function
	'Error_In_PasteString:
	'  PasteString = False
	'  Resume Exit_PasteString
	'Exit_PasteString:
	'End Function
	'Private Function CutString() As Integer
	'Dim ClipString As String, Length As Integer
	'Dim CurrentPosition As Integer, PreviousPosition As Integer, Character As String * 1
	'Dim StringToTransfer As String, Row As Integer, Col As Integer
	'On Error GoTo Error_In_CutString
	'  ClipString = Clipboard.GetText()
	'  Length = Len(ClipString)
	'  If Length > 0 Then
	'    PreviousPosition = 1
	'    CurrentPosition = 1
	'    Col = 1
	'    Row = 1
	'    While PreviousPosition <= Length
	'      Character = Mid$(ClipString, CurrentPosition, 1)
	'      Select Case Asc(Character)
	'        Case 10
	'          CurrentPosition = CurrentPosition + 1
	'          PreviousPosition = CurrentPosition
	'        Case 13, 9
	'          StringToTransfer = Mid$(ClipString, PreviousPosition, CurrentPosition - PreviousPosition)
	'          If Not (PasteString(StringToTransfer, Row, Col)) Then
	'            MsgBox "Error while pasting data.", 64, AppName_For_Display_long
	'          End If
	'          Col = Col Mod (Number_Component + 1) + 1
	'          If Col = 1 Then
	'           Row = Row + 1
	'           If Row > Number_Max_Influent_Points Then GoTo Too_Many_Points
	'          End If
	'          CurrentPosition = CurrentPosition + 1
	'          PreviousPosition = CurrentPosition
	'        Case Else
	'          CurrentPosition = CurrentPosition + 1
	'          Character = Mid$(ClipString, CurrentPosition, 1)
	'      End Select
	'    Wend
	'  Else
	'  End If
	'  CutString = True
	'  Exit Function
	'Too_Many_Points:
	'  CutString = True
	'  MsgBox "Too much data was selected. Only the first " & Format$(Number_Max_Influent_Points, "0") & " points were pasted.", 64, AppName_For_Display_long
	'  GoTo Exit_CutString
	'Error_In_CutString:
	'  CutString = False
	'  Resume Exit_CutString
	'Exit_CutString:
	'End Function
End Class