Option Strict Off
Option Explicit On
Friend Class frmExcelCurves
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer

	Dim frmExcelCurves_loading_now As Boolean
	
	
	
	
	
	
	Const frmExcelCurves_declarations_end As Short = 0
	
	
	Sub Do_The_Print_To_F1Book()


		'Dim f1 As System.Windows.Forms.Control
		'	Dim f1 As VCIF1Lib.F1Book
		'		f1 = Me.f1book

		Dim f1 As DataGridView
		f1 = Me.f1bookDataGrid

		'Dim sserror As Integer
		Dim i As Short
		Dim j As Short
		Dim last_row As Short
		Dim last_col As Short
		Dim dblConvertedCP As Double
		Dim dbl_CPConversionFactor() As Double
		Dim OUT_strYAxisTitle As String
		If (CPHSDM_Excel) Then
			'
			'    ----    CPHSDM RESULTS.    ----
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.MinCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.MinCol = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.MinRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.MinRow = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(1, 1) = "CPHSDM Results -- Filename = " & Chr(34) & Trim(Filename) & Chr(34)
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(3, 1) = "Time"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(3, 2) = "BVT"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(3, 3) = "Usage Rate"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(3, 4) = CPM_Results.Component.Name
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(4, 1) = "Minutes"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(4, 2) = "-"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(4, 3) = "m³/kg GAC"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'	f1.EntryRC(4, 4) = "-"
			f1.RowCount = 105
			f1.ColumnCount = 5
			f1.Rows(1).Cells(1).Value = "CPHSDM Results -- Filename = " & Chr(34) & Trim(Filename) & Chr(34)
			f1.Rows(3).Cells(1).Value = "Time"
			f1.Rows(3).Cells(2).Value = "BVT"
			f1.Rows(3).Cells(3).Value = "Usage Rate"
			f1.Rows(3).Cells(4).Value = CPM_Results.Component.Name
			f1.Rows(4).Cells(1).Value = "Minutes"
			f1.Rows(4).Cells(2).Value = "-"
			f1.Rows(4).Cells(3).Value = "m³/kg GAC"
			f1.Rows(4).Cells(4).Value = "-"

			For i = 1 To 100
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(4 + i).Cells(1).Value = Trim(Str(CPM_Results.T(i)))
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(4 + i).Cells(2).Value = Trim(Str(CPM_Results.T(i) * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2.0#) ^ 2))
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(4 + i).Cells(3).Value = Trim(Str(CPM_Results.T(i) * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight))
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(4 + i).Cells(4).Value = Trim(Str(CPM_Results.C_Over_C0(i)))
				'
				' FORMAT NUMBERS IN THIS ROW AS "GENERAL" FORMAT.
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'			f1.SelStartRow = 4 + i
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'		f1.SelEndRow = 4 + i
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'		f1.SelStartCol = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'			f1.SelEndCol = 4
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.FormatGeneral. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'			f1.FormatGeneral()

			Next i
			last_row = 4 + 100
			last_col = 4
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		f1.MaxCol = last_col
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		f1.MaxRow = last_row
		Else
			'
			'    ----    PSDM RESULTS.    ----
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.MinCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		f1.MinCol = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.MinRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		f1.MinRow = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f1.ColumnCount = 5 + Results.NComponent
			f1.RowCount = 5 + Results.npoints
			f1.Rows(1).Cells(1).Value = "PSDM Results -- Filename = " & Chr(34) & Trim(Filename) & Chr(34)
			'
			' GET THE UNIT CONVERSION FACTOR FOR EACH CHEMICAL.
			'
			'UPGRADE_WARNING: Lower bound of array dbl_CPConversionFactor was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim dbl_CPConversionFactor(Results.NComponent)
			For i = 1 To Results.NComponent
				dbl_CPConversionFactor(i) = CBOYAXISTYPE_GetUnitConversion(CShort(VB6.GetItemData(frmModelPSDMResults.cboYAxisType, frmModelPSDMResults.cboYAxisType.SelectedIndex)), Results.is_psdm_in_room_model, Results.AnyCrCloseToZero, i, Results.Bed.Phase, OUT_strYAxisTitle)
			Next i
			'
			' SET UP THE COLUMN HEADINGS.
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f1.Rows(3).Cells(1).Value = "Time"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f1.Rows(3).Cells(2).Value = "BVT"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f1.Rows(3).Cells(3).Value = "Usage Rate"
			For i = 1 To Results.NComponent
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(3).Cells(3 + i).Value = Trim(Results.Component(i).Name)
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(4).Cells(3 + i).Value = OUT_strYAxisTitle
				''''f1.EntryRC(4, 3 + i) = "-"
			Next i
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f1.Rows(4).Cells(1).Value = "Minutes"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f1.Rows(4).Cells(2).Value = "-"
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f1.Rows(4).Cells(3).Value = "m³/kg GAC"
			'
			' OUTPUT THE VALUES.
			'
			For i = 1 To Results.npoints
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(4 + i).Cells(1).Value = Trim(Str(Results.T(i)))
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(4 + i).Cells(2).Value = Trim(Str(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2))
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f1.Rows(4 + i).Cells(3).Value = Trim(Str(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.Weight))
				For j = 1 To Results.NComponent
					dblConvertedCP = Results.CP(j, i) * dbl_CPConversionFactor(j)
					'UPGRADE_WARNING: Couldn't resolve default property of object f1.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					f1.Rows(4 + i).Cells(3 + j).Value = Trim(Str(dblConvertedCP))
					''''f1.EntryRC(4 + i, 3 + j) = Trim$(Str$(Results.CP(j, i)))
				Next j
				'
				' FORMAT NUMBERS IN THIS ROW AS "GENERAL" FORMAT.
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				f1.SelStartRow = 4 + i
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				f1.SelEndRow = 4 + i
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				f1.SelStartCol = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				f1.SelEndCol = 3 + j
				'UPGRADE_WARNING: Couldn't resolve default property of object f1.FormatGeneral. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'				f1.FormatGeneral()
				''''f1.NumberFormat = "0.0"
			Next i
			last_row = 4 + Results.npoints
			last_col = 3 + Results.NComponent
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'			f1.MaxCol = last_col
			'UPGRADE_WARNING: Couldn't resolve default property of object f1.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'			f1.MaxRow = last_row
		End If
		'temp = ""
		'For i = 1 To Results.NPoints
		' temp = Format$(Results.T(i), "0.00")
		' temp = temp & "       " & Format$(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.Length / Pi / (Results.Bed.Diameter / 2) ^ 2, "0.00")
		' temp = temp & "       " & Format$(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.Weight, "0.00")
		' For j = 1 To Results.NComponent
		'   temp = temp & "          " & Format$(Results.CP(j, i), "0.000")
		' Next j
		' Print #f, temp
		' temp = ""
		'Next i
		'Close f
	End Sub
	
	
	Function f1file_saveas(ByRef fn_force As String) As Short
		Dim F1FileExcel4 As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNOverwritePrompt As Object
		Dim fn_saveas As String
		If (fn_force <> "") Then
			fn_saveas = fn_force
		Else
			'INPUT NEW FILENAME.
			On Error GoTo err_filesaveas
			'UPGRADE_WARNING: Couldn't resolve default property of object frmExcelCurves.CommonDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'Me.CommonDialog1.Filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
			Me.SaveFileDialog1.Filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"

			'UPGRADE_WARNING: Couldn't resolve default property of object frmExcelCurves.CommonDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'Me.CommonDialog1.FilterIndex = 2
			Me.SaveFileDialog1.FilterIndex = 2

			'UPGRADE_WARNING: Couldn't resolve default property of object frmExcelCurves.CommonDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'Me.CommonDialog1.CancelError = True

			'UPGRADE_WARNING: Couldn't resolve default property of object frmExcelCurves.CommonDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNOverwritePrompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'Me.CommonDialog1.flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist

			'UPGRADE_WARNING: Couldn't resolve default property of object frmExcelCurves.CommonDialog1.ShowSave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'Me.CommonDialog1.ShowSave()
			Me.SaveFileDialog1.ShowDialog()

			'UPGRADE_WARNING: Couldn't resolve default property of object frmExcelCurves.CommonDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fn_saveas = Trim(Me.SaveFileDialog1.FileName)
			If (fn_saveas = "") Then
				GoTo exit_save_did_not_go_okay
			End If
		End If
		'
		' SAVE THIS FILE.
		'
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		'UPGRADE_WARNING: Couldn't resolve default property of object frmExcelCurves.CommonDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object f1book.Write. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	f1book.Write(Me.CommonDialog1.FileName, F1FileExcel4)
		'
		' TODO: DETERMINE WHAT VERSIONS OF EXCEL THIS IS COMPATIBLE WITH!
		'
		Me.Cursor = System.Windows.Forms.Cursors.Default
exit_save_went_okay: 
		'SAVE WENT OKAY.
		f1file_saveas = True
		Exit Function
exit_save_did_not_go_okay: 
		'SAVE DID NOT GO OKAY.
		f1file_saveas = False
		Exit Function
err_filesaveas: 
		f1file_saveas = False
		Resume exit_save_did_not_go_okay
	End Function
	
	
	'UPGRADE_WARNING: Form event frmExcelCurves.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmExcelCurves_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If (frmExcelCurves_loading_now) Then
			frmExcelCurves_loading_now = False
			'Call frmExcelCurves_Resize(Me, New System.EventArgs())
		End If
	End Sub
	Private Sub frmExcelCurves_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		Dim msg As String
		If (CPHSDM_Excel) Then
			msg = "Pre-Print of CPHSDM Results"
		Else
			msg = "Pre-Print of PSDM Results"
		End If
		'If (Trim$(NowProj.Filename) = "") Then
		'  msg = msg & "(untitled)"
		'Else
		'  msg = msg & NowProj.Filename
		'End If
		Me.Text = msg
		'Me.Width = VB6.TwipsToPixelsX(600)
		'Me.Height = VB6.TwipsToPixelsY(450)
		Call CenterOnForm(Me, frmMain)
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized 'maximized
		frmExcelCurves_loading_now = True
		Call Do_The_Print_To_F1Book()
	End Sub
	'UPGRADE_WARNING: Event frmExcelCurves.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmExcelCurves_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		'rs.ResizeAllControls(Me)

		Dim XXX As Integer
		If (frmExcelCurves_loading_now) Then Exit Sub
		If (Me.WindowState = 1) Then
			'CANNOT RESIZE WHEN MINIMIZED; EXIT OUTTA HERE.
			Exit Sub
		End If
		'If (Me.Width < 500) Then Me.Width = 500
		'If (Me.Height < 200) Then Me.Height = 200
		'UPGRADE_WARNING: Couldn't resolve default property of object f1book.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'XXX = VB6.PixelsToTwipsX(Me.Width) - (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(Me.ClientRectangle.Width)) - 2 * f1bookDataGrid.Left
		'If (XXX < 1000) Then
		'XXX = 1000
		'End If
		'UPGRADE_WARNING: Couldn't resolve default property of object f1book.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'f1bookDataGrid.Width = 450
		'UPGRADE_WARNING: Couldn't resolve default property of object f1book.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'XXX = VB6.PixelsToTwipsY(Me.Height) - (VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(Me.ClientRectangle.Height)) - 2 * f1bookDataGrid.Top
		'		If (XXX > 1000) Then
		'UPGRADE_WARNING: Couldn't resolve default property of object f1book.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'f1bookDataGrid.Height = 300
		'End If
	End Sub
	
	
	Public Sub mnuEditItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditItem.Click
		Dim Index As Short = mnuEditItem.GetIndex(eventSender)
		Dim f1 As DataGridView
		f1 = Me.f1bookDataGrid
		f1.ClipboardCopyMode =
		DataGridViewClipboardCopyMode.EnableWithoutHeaderText
		Select Case Index
			Case 10 'copy
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'	f1book.EditCopy()
				Clipboard.SetDataObject(f1.GetClipboardContent())

			Case 20 'EDIT--COPY ENTIRE TABLE.

				f1.SelectAll()
				Clipboard.SetDataObject(f1.GetClipboardContent())
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'	f1book.SelStartCol = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'	f1book.SelEndCol = f1book.MaxCol
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'	f1book.SelStartRow = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'	f1book.SelEndRow = f1book.MaxRow
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'	f1book.EditCopy()
		End Select
	End Sub
	Public Sub mnuFileItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileItem.Click
		Dim Index As Short = mnuFileItem.GetIndex(eventSender)
		On Error GoTo err_mnuFileItem_Click
		Select Case Index
			Case 40 'save as ...
				Call f1file_saveas("")
			Case 50 'page setup ...
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.FilePageSetupDlg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			f1book.FilePageSetupDlg()
			Case 55 'print setup ...
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.FilePrintSetupDlg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			f1book.FilePrintSetupDlg()
			Case 60 'print ...
				'UPGRADE_WARNING: Couldn't resolve default property of object f1book.FilePrint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			f1book.FilePrint(True)
			Case 199 'close
				Me.Close()
		End Select
exit_err_mnuFileItem_Click: 
		Exit Sub
err_mnuFileItem_Click: 
		'PROBABLY A CANCEL FROM f1book.FilePrint.
		Resume exit_err_mnuFileItem_Click
	End Sub

End Class