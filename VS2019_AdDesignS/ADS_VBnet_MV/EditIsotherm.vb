Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmEditIsotherm
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Dim FORM_MODE As Short
	Const FORM_MODE_EDIT_DATABASE As Short = 2

	'UPGRADE_ISSUE: Database object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Dim DB_Isotherm As dao.Database
	Dim Find_String As String
	
	Dim HALT_LSTCOMPO As Boolean
	Dim HALT_LSTRANGE As Boolean
	
	
	
	
	
	
	Const frmEditIsotherm_declarations_end As Boolean = True
	
	
	Sub frmEditIsotherm_EditDatabase()
		On Error GoTo err_frmEditIsotherm_EditDatabase
		'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
		'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
		'UPGRADE_WARNING: Couldn't resolve default property of object Ws1.OpenDatabase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'DB_Isotherm = DAOEngine.OpenDatabase(fn_DB_Isotherm)   'From Ws1 to DaoEngine ?? Shang
		DB_Isotherm = Ws1.OpenDatabase(fn_DB_Isotherm, True, False, ";pwd=" & decrypt_string(Encrypted_User_Password))
		'Set DB_Isotherm = ws1.OpenDatabase(fn_DB_Isotherm)
		FORM_MODE = FORM_MODE_EDIT_DATABASE
		Me.ShowDialog()
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DB_Isotherm.Close()
		Exit Sub
exit_err_frmEditIsotherm_EditDatabase: 
		Exit Sub
err_frmEditIsotherm_EditDatabase: 
		Call Show_Trapped_Error("frmEditIsotherm_EditDatabase")
		Resume exit_err_frmEditIsotherm_EditDatabase
	End Sub
	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			mnuChemicalItem(1).Enabled = False
			mnuChemicalItem(2).Enabled = False
			mnuChemicalItem(3).Enabled = False
			mnuIsothermItem(1).Enabled = False
			mnuIsothermItem(2).Enabled = False
			mnuIsothermItem(3).Enabled = False
			mnuIsothermItem(4).Enabled = False
		End If
	End Sub
	
	
	Sub populate_lstCompo()
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim Current_Criteria As String
		Dim SAVE_CURRENT_POSITION As Integer
		Dim NEW_LISTINDEX As Short
		Dim This_ID As Integer
		Dim NumRecords As Integer
		Dim SortCode As String
		Dim TempStr As New VB6.FixedLengthString(15)
		Dim Output_Line As String
		Dim ThisChemicalName As String
		Dim ThisChemicalCAS As String
		On Error GoTo err_populate_lstCompo
		'SAVE CURRENT POSITION.
		If (lstCompo.Items.Count > 0) And (lstCompo.SelectedIndex >= 0) Then
			SAVE_CURRENT_POSITION = VB6.GetItemData(lstCompo, lstCompo.SelectedIndex)
		Else
			SAVE_CURRENT_POSITION = -1
		End If
		'SET UP SEARCH CRITERIA.
		'UPGRADE_WARNING: Couldn't resolve default property of object optSort(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optSort_0.Checked) Then SortCode = "Name, CAS"
		'UPGRADE_WARNING: Couldn't resolve default property of object optSort(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optSort_1.Checked) Then SortCode = "CAS, Name"
		Current_Criteria = "select * from [Chemicals] " & "order by " & SortCode
		'START SEARCH.
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveLast()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NumRecords = Rs1.RecordCount
		On Error GoTo err_populate_lstCompo
		'POPULATE LISTBOX.
		lstCompo.Items.Clear()
		If (NumRecords = 0) Then
			'NO RECORDS AVAILABLE.
			lstCompo.Visible = False
			lblEmpty_lstCompo.SetBounds(lstCompo.Left, lstCompo.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			lblEmpty_lstCompo.Visible = True
		Else
			'DISPLAY RECORDS.
			lstCompo.Visible = True
			lblEmpty_lstCompo.Visible = False
			NEW_LISTINDEX = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Do Until Rs1.EOF

				'UPGRADE_WARNING: Untranslated statement in populate_lstCompo. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstCompo. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstCompo. Please check source code.
				This_ID = Database_Get_Long(Rs1, "Compo ID")
				ThisChemicalName = Database_Get_String(Rs1, "Name")
				ThisChemicalCAS = Database_Get_String(Rs1, "CAS")





				TempStr.Value = ThisChemicalCAS
				'THIS STRING IS ENSURED TO BE 15 CHARACTERS LONG.
				Output_Line = TempStr.Value & " " & ThisChemicalName
				NEW_LISTINDEX = lstCompo.Items.Add(New VB6.ListBoxItem(Output_Line, This_ID))
				If (SAVE_CURRENT_POSITION <> -1) Then
					If (SAVE_CURRENT_POSITION = This_ID) Then
						'UPGRADE_ISSUE: ListBox property lstCompo.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
						'NEW_LISTINDEX = lstCompo.NewIndex
					End If
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveNext()
			Loop
			If (lstCompo.Items.Count > 0) Then
				HALT_LSTCOMPO = True
				lstCompo.SelectedIndex = NEW_LISTINDEX
				HALT_LSTCOMPO = False
			End If
		End If
		'CLOSE DATABASE AND EXIT.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		Exit Sub
exit_err_populate_lstCompo: 
		Exit Sub
err_populate_lstCompo: 
		Call Show_Trapped_Error("populate_lstCompo")
		Resume exit_err_populate_lstCompo
	End Sub
	Sub populate_lstRange(ByRef ThisCAS As String, ByRef ThisChemical As String)
		Dim PHASE_CODE As Short
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As Dao.Recordset
		Dim Current_Criteria As String
		Dim SAVE_CURRENT_POSITION As Integer
		Dim This_ID As Integer
		Dim NEW_LISTINDEX As Integer
		Dim NumRecords As Integer
		Dim PhaseCode As String
		Dim ThisCMin As Double
		Dim ThisCMax As Double
		Dim ThisPHMin As Double
		Dim ThisPHMax As Double
		Dim ThisDbl As Double
		Dim ThisOutput As String
		On Error GoTo err_populate_lstRange
		''GET PHASE CODE.
		'Select Case Bed.Phase
		'  Case 0: PhaseCode = "Liquid"
		'  Case 1: PhaseCode = "Gas"
		'End Select
		'SAVE CURRENT POSITION.
		If (lstRange(0).Items.Count > 0) And (lstRange(0).SelectedIndex >= 0) Then
			SAVE_CURRENT_POSITION = VB6.GetItemData(lstRange(0), lstRange(0).SelectedIndex)
		Else
			SAVE_CURRENT_POSITION = -1
		End If
		'SET UP SEARCH CRITERIA.
		If (Trim(ThisCAS) = "0") Then ThisCAS = ""
		If (Trim(ThisCAS) <> "") Then
			Current_Criteria = "select * from Isotherms" & " where [Component Number] = " & Trim(ThisCAS) & " order by CarbonName, ID"
		Else
			Current_Criteria = "select * from Isotherms" & " where Name = '" & Trim(ThisChemical) & "'" & " order by CarbonName, ID"
		End If
		'START SEARCH.
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveLast()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NumRecords = Rs1.RecordCount
		On Error GoTo err_populate_lstRange
		'POPULATE LISTBOX.
		lstRange(0).Items.Clear()
		lstRange(1).Items.Clear()
		If (NumRecords = 0) Then
			'NO RECORDS AVAILABLE.
			'UPGRADE_WARNING: Couldn't resolve default property of object fraTwo.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GroupBox2.Text = "No Isotherms Available."
		Else
			'DISPLAY RECORDS.
			'UPGRADE_WARNING: Couldn't resolve default property of object fraTwo.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GroupBox2.Text = Trim(Str(NumRecords)) & " " & "Isotherm" & IIf(NumRecords = 1, "", "s") & " Available"
			NEW_LISTINDEX = -1
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Do Until Rs1.EOF

				'UPGRADE_WARNING: Untranslated statement in populate_lstRange. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstRange. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstRange. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstRange. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstRange. Please check source code.
				This_ID = Database_Get_Long(Rs1, "ID")
				ThisCMin = Database_Get_Double(Rs1, "C min")
				ThisCMax = Database_Get_Double(Rs1, "C max")
				ThisPHMin = Database_Get_Double(Rs1, "pH min")
				ThisPHMax = Database_Get_Double(Rs1, "pH max")






				If (ThisPHMin = 0#) And (ThisPHMax = 0#) Then
					ThisOutput = "No pH Range"
				Else
					If (ThisPHMin = 0#) Or (ThisPHMax = 0#) Then
						If (ThisPHMin <> 0#) Then ThisDbl = ThisPHMin
						If (ThisPHMax <> 0#) Then ThisDbl = ThisPHMax
						ThisOutput = VB6.Format(ThisDbl, "0.000")
					Else
						ThisOutput = VB6.Format(ThisPHMin, "0.000") & " - " & VB6.Format(ThisPHMax, "0.000")
					End If
				End If
				NEW_LISTINDEX = lstRange(0).Items.Add(New VB6.ListBoxItem(ThisOutput, This_ID))
				If (ThisCMin = 0#) And (ThisCMax = 0#) Then
					ThisOutput = "No Conc. Range"
				Else
					If (ThisCMin = 0#) Or (ThisCMax = 0#) Then
						If (ThisCMin <> 0#) Then ThisDbl = ThisCMin
						If (ThisCMax <> 0#) Then ThisDbl = ThisCMax
						ThisOutput = VB6.Format(ThisDbl, "0.000")
					Else
						ThisOutput = VB6.Format(ThisCMin, "0.000") & " - " & VB6.Format(ThisCMax, "0.000")
					End If
				End If
				lstRange(1).Items.Add(New VB6.ListBoxItem(ThisOutput, This_ID))
				If (SAVE_CURRENT_POSITION <> -1) Then
					If (SAVE_CURRENT_POSITION = This_ID) Then
						'UPGRADE_ISSUE: ListBox property lstRange.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
						'NEW_LISTINDEX = lstRange(0).NewIndex
					End If
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveNext()
			Loop
			If (lstRange(0).Items.Count > 0) And (NEW_LISTINDEX > -1) Then
				lstRange(0).SelectedIndex = NEW_LISTINDEX
				lstRange(1).SelectedIndex = NEW_LISTINDEX
			End If
		End If
		'CLOSE DATABASE AND EXIT.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		Exit Sub
exit_err_populate_lstRange: 
		Exit Sub
err_populate_lstRange: 
		Call Show_Trapped_Error("populate_lstRange")
		Resume exit_err_populate_lstRange
	End Sub
	Sub populate_lblValue(ByRef This_ID As Integer)
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim Current_Criteria As String
		Dim NumRecords As Integer
		On Error GoTo err_populate_lblValue
		'SET UP SEARCH CRITERIA.
		Current_Criteria = "select * from Isotherms" & " where ID = " & Trim(Str(This_ID)) & " order by CarbonName"
		'START SEARCH.
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveLast()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NumRecords = Rs1.RecordCount
		On Error GoTo err_populate_lblValue
		'POPULATE LISTBOX.
		If (NumRecords = 0) Then
			'COULD NOT FIND THAT ISOTHERM (WEIRD PROBLEM).
		Else
			'DISPLAY RECORD.
			'UPGRADE_WARNING: Untranslated statement in populate_lblValue. Please check source code.
			'UPGRADE_WARNING: Untranslated statement in populate_lblValue. Please check source code.
			'UPGRADE_WARNING: Untranslated statement in populate_lblValue. Please check source code.
			'UPGRADE_WARNING: Untranslated statement in populate_lblValue. Please check source code.
			'UPGRADE_WARNING: Untranslated statement in populate_lblValue. Please check source code.
			'UPGRADE_WARNING: Untranslated statement in populate_lblValue. Please check source code.
			'UPGRADE_WARNING: Untranslated statement in populate_lblValue. Please check source code.

			'DISPLAY RECORD.
			Call AssignCaptionAndTag(lblValue(0), Database_Get_Double(Rs1, "K"))
			Call AssignCaptionAndTag(lblValue(1), Database_Get_Double(Rs1, "1/n"))
			Call AssignCaptionAndTag(lblValue(2), Database_Get_String(Rs1, "CarbonName"))
			Call AssignCaptionAndTag(lblTemperature, Database_Get_Double(Rs1, "Tmin"))
			Call AssignCaptionAndTag(lblPhase, Database_Get_String(Rs1, "Phase"))
			Call AssignCaptionAndTag(lblValue(3), Database_Get_String(Rs1, "Source"))
			Call AssignCaptionAndTag(lblComments, Database_Get_String(Rs1, "Comments"))



		End If
		'CLOSE DATABASE AND EXIT.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		Exit Sub
exit_err_populate_lblValue: 
		Exit Sub
err_populate_lblValue: 
		Call Show_Trapped_Error("populate_lblValue")
		Resume exit_err_populate_lblValue
	End Sub
	
	
	Sub Clear_lblValue()
		lblValue(0).Text = "" 'ISODB : K.
		lblValue(1).Text = "" 'ISODB : 1/N.
		lblValue(2).Text = "" 'ISODB : ADSORBENT TYPE.
		lblTemperature.Text = "" 'ISODB : TEMP.
		lblPhase.Text = "" 'ISODB : PHASE.
		lblValue(3).Text = "" 'ISODB : SOURCE.
		lblComments.Text = "" 'ISODB : COMMENTS.
	End Sub
	
	
	'Returns:
	'- TRUE = Succeeded
	'- FALSE = Failed
	Function Search_String(ByRef J As Short, ByRef ShowErrorMessages As Short) As Boolean
		Dim i As Short
		Dim Res As Short
		'If (fraIsothermDB.Visible) Then
		'  lstCompo.SetFocus
		'End If
		'For I = J + 1 To lstCompo.ListCount
		For i = J + 1 To lstCompo.Items.Count - 1
			Res = InStr(1, VB6.GetItemString(lstCompo, i), Find_String, 1)
			If (Res > 0) Then
				'NOTE: BY HALTING lstCompo_Click(), THIS ALLOWS THE
				'COMPONENT TO BE SELECTED WITHOUT CLEARING THE ISOTHERM DB
				'VALUES OF K AND 1/N.
				'lstCompo.ListIndex = I
				HALT_LSTCOMPO = True
				Call Do_Select_Component(i)
				HALT_LSTCOMPO = False
				'If (fraIsothermDB.Visible) Then lstCompo.SetFocus
				Search_String = True
				Exit Function
			End If
		Next i
		For i = 0 To J
			Res = InStr(1, VB6.GetItemString(lstCompo, i), Find_String, 1)
			If (Res > 0) Then
				'NOTE: BY HALTING lstCompo_Click(), THIS ALLOWS THE
				'COMPONENT TO BE SELECTED WITHOUT CLEARING THE ISOTHERM DB
				'VALUES OF K AND 1/N.
				'lstCompo.ListIndex = I
				HALT_LSTCOMPO = True
				Call Do_Select_Component(i)
				HALT_LSTCOMPO = False
				'If (fraIsothermDB.Visible) Then lstCompo.SetFocus
				Search_String = True
				Exit Function
			End If
		Next i
		'----- If not found, show error message: -----
		If (ShowErrorMessages) Then
			Call Show_Error("String Not Found: " & Chr(34) & Trim(Find_String) & Chr(34))
		End If
		Search_String = False
	End Function
	'Returns:
	'- TRUE = Succeeded
	'- FALSE = Failed
	Function Do_Search_For_Text(ByRef ShowErrorMessages As Short) As Short
		Dim LIST_INDEX As Short
		LIST_INDEX = lstCompo.SelectedIndex
		Do_Search_For_Text = Search_String(LIST_INDEX, ShowErrorMessages)
	End Function
	
	
	Private Sub cmdFind_Click(ByRef Index As Short)
		Dim NewName As String
		Dim USER_HIT_CANCEL As Boolean
		Select Case Index
			Case 0 'FIND.
				NewName = Find_String
				Do While (1 = 1)
					NewName = frmNewName.frmNewName_GetName("Search for String", "Enter the string to find:", NewName, USER_HIT_CANCEL)
					If (USER_HIT_CANCEL) Then Exit Sub
					NewName = Trim(NewName)
					If (NewName <> "") Then Exit Do
					Call Show_Error("You may only enter a non-blank search string.")
				Loop 
				Find_String = NewName
				Call Do_Search_For_Text(True)
			Case 1 'FIND AGAIN.
				Call Do_Search_For_Text(True)
		End Select
	End Sub
	
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub
	
	Private Sub frmEditIsotherm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		'MISC INITS.
		'	Me.Height = VB6.TwipsToPixelsY(7395)
		'	Me.Width = VB6.TwipsToPixelsX(9500)
		Call CenterOnForm(Me, frmMain)
		'UPGRADE_WARNING: Couldn't resolve default property of object optSort().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_optSort_0.Enabled = True
		'UPGRADE_WARNING: Couldn't resolve default property of object optSort().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_optSort_1.Enabled = True
		Find_String = ""
		'RE-POPULATE CHEMICAL LIST.
		Call populate_lstCompo()
		' DEMO SETTINGS.
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	
	
	Sub Do_Select_Component(ByRef WhichComp As Short)
		''''Dim THIS_ITEMDATA As Long
		Dim ThisText As String
		Dim ThisCAS As String
		Dim ThisChemical As String
		If (WhichComp < 0) Or (lstCompo.Items.Count <= 0) Then
			Exit Sub
		End If
		lstCompo.SelectedIndex = WhichComp
		''''THIS_ITEMDATA = lstCompo.ItemData(lstCompo.ListIndex)
		'EXTRACT CAS NUMBER AND COMPONENT NAME.
		ThisText = VB6.GetItemString(lstCompo, WhichComp)
		ThisCAS = Trim(VB.Left(ThisText, 15))
		ThisChemical = Trim(Mid(ThisText, 16, Len(ThisText) - 15))
		Call populate_lstRange(ThisCAS, ThisChemical)
	End Sub
	
	'UPGRADE_WARNING: Event lstCompo.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstCompo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCompo.SelectedIndexChanged
		If (HALT_LSTCOMPO) Then Exit Sub
		HALT_LSTCOMPO = True
		Call Do_Select_Component(lstCompo.SelectedIndex)
		HALT_LSTCOMPO = False
		'INVALIDATE EXISTING ISOTHERM RECORD LINK (IF ANY).
		'Component(0).IsothermDB_OneOverN = -1#
		'Component(0).IsothermDB_K = -1#
		HALT_LSTRANGE = True
		lstRange(0).SelectedIndex = -1
		lstRange(1).SelectedIndex = -1
		HALT_LSTRANGE = False
		'CLEAR EXISTING RECORD DATA.
		Call Clear_lblValue()
		'Call frmFreundlich_Refresh
	End Sub
	Private Sub lstCompo_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstCompo.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If ((Button And 2) = 2) Then
			'UPGRADE_ISSUE: Form method frmEditIsotherm.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'Me.PopupMenu(mnuChemical)   //out shang
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event lstRange.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstRange_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstRange.SelectedIndexChanged
		Dim Index As Short = lstRange.GetIndex(eventSender)
		If (HALT_LSTRANGE) Then Exit Sub
		HALT_LSTRANGE = True
		'KEEP THE RANGE LISTBOXES IN SYNCH.
		Select Case Index
			Case 0 : lstRange(1).SelectedIndex = lstRange(0).SelectedIndex
			Case 1 : lstRange(0).SelectedIndex = lstRange(1).SelectedIndex
		End Select
		''TRANSFER LINK TO COMPONENT(0) STRUCTURE.
		'Component(0).IsothermDB_Component_Name = Trim$(lstCompo.List(lstCompo.ListIndex))
		'Component(0).IsothermDB_Range_Num = lstRange(0).ListIndex
		'DISPLAY ISOTHERM RECORD.
		Call populate_lblValue(VB6.GetItemData(lstRange(0), lstRange(0).SelectedIndex))
		''THROW DIRTY FLAG.
		'Call frmFreundlich_DirtyStatus_Throw
		''REFRESH WINDOW.
		'Call frmFreundlich_Refresh
		HALT_LSTRANGE = False
	End Sub
	Private Sub lstRange_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstRange.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lstRange.GetIndex(eventSender)
		'If ((Button And 2) = 2) Then
		'UPGRADE_ISSUE: Form method frmEditIsotherm.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'Me.PopupMenu(mnuIsotherm)  out shang
		'End If
	End Sub
	
	
	Public Sub mnuChemicalItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuChemicalItem.Click
		Dim Index As Short = mnuChemicalItem.GetIndex(eventSender)
		Dim THIS_CHEMICAL_ID As Integer
		Dim DummyStr1, DummyStr2 As String
		Dim DummyBool1, DummyBool2 As Boolean
		Dim USER_HIT_CANCEL As Boolean
		Dim NewName As String
		Dim NewCAS As String
		Dim OldName As String
		Dim OldCAS As String
		'Dim USER_HIT_CANCEL As Boolean
		Dim Current_Criteria As String
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim msg As String
		Dim RetVal As Short
		Dim i As Short
		Dim Select_Index As Short
		Dim NumRecords As Short
		Dim RecordCount_Chemicals As Short
		Dim RecordCount_Isotherms_CAS As Short
		Dim RecordCount_Isotherms_Name As Short
		On Error GoTo err_mnuChemicalItem_Click
		If (Index = 2) Or (Index = 3) Then
			If (lstCompo.SelectedIndex < 0) Or (lstCompo.Items.Count = 0) Then
				Call Show_Error("You must first select a chemical.")
				Exit Sub
			End If
			THIS_CHEMICAL_ID = VB6.GetItemData(lstCompo, lstCompo.SelectedIndex)
		End If
		Select Case Index
			Case 1 'new /////////////////////////////////////////////////////////////////////////////////////////////////
				NewName = "New Chemical"
				Do While (1 = 1)
					Call frmEditIsothermCAS.frmEditIsothermCAS_Run("Creating New Chemical", "New CAS Number", "New Chemical Name", "*", "*", "*", "*", "&Save", USER_HIT_CANCEL, NewCAS, NewName, DummyStr1, DummyStr2, DummyBool1, DummyBool2)
					If (USER_HIT_CANCEL) Then Exit Sub
					NewName = Trim(NewName)
					NewCAS = Trim(NewCAS)
					If (NewName <> "") Then Exit Do
					Call Show_Error("The chemical name must be a non-blank string.")
				Loop 
				'ADD THE NEW CHEMICAL RECORD.
				Current_Criteria = "select * from [Chemicals]"
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'THE FIELD [Compo ID] IS AUTOMATICALLY UPDATED.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Name").Value = NewName
				If (NewCAS <> "") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("CAS").Value = NewCAS
				Else
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("CAS").Value = System.DBNull.Value
				End If

				'UPGRADE_WARNING: Untranslated statement in mnuChemicalItem_Click. Please check source code.
				THIS_CHEMICAL_ID = Database_Get_Long(Rs1, "Compo ID")


				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'REDISPLAY WINDOW.
				Call populate_lstCompo()
				'SELECT THE NEW MANUFACTURER.
				Select_Index = 0
				For i = 0 To lstCompo.Items.Count - 1
					If (VB6.GetItemData(lstCompo, i) = THIS_CHEMICAL_ID) Then
						Select_Index = i
						Exit For
					End If
				Next i
				If (lstCompo.Items.Count > 0) Then
					lstCompo.SelectedIndex = Select_Index
				End If
			Case 2 'edit current /////////////////////////////////////////////////////////////////////////////////////////////////
				'START THE EDIT PROCESS.
				Current_Criteria = "select * from [Chemicals] where " & "[Compo ID] = " & Trim(Str(THIS_CHEMICAL_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Edit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Edit()
				'DO USER INPUT.

				'UPGRADE_WARNING: Untranslated statement in mnuChemicalItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuChemicalItem_Click. Please check source code.
				OldName = Database_Get_String(Rs1, "Name")
				OldCAS = Database_Get_String(Rs1, "CAS")



				NewName = OldName
				NewCAS = OldCAS
				DummyBool1 = True
				DummyBool2 = True
				Do While (1 = 1)
					Call frmEditIsothermCAS.frmEditIsothermCAS_Run("Editing a Chemical", "^Current CAS Number", "^Current Chemical Name", "New CAS Number", "New Chemical Name", "Modify all isotherms with the same CAS number", "Modify all isotherms with the same chemical name", "&Save", USER_HIT_CANCEL, OldCAS, OldName, NewCAS, NewName, DummyBool1, DummyBool2)
					If (USER_HIT_CANCEL) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.CancelUpdate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Rs1.CancelUpdate()
						Exit Sub
					End If
					NewName = Trim(NewName)
					NewCAS = Trim(NewCAS)
					If (NewName = "") Then
						Call Show_Error("The chemical name must be a non-blank string.")
					Else
						If (OldCAS = "") And (NewCAS <> "") And (DummyBool1) Then
							Call Show_Error("You cannot automatically assign CAS numbers " & "to all isotherm records currently without CAS numbers.")
						Else
							Exit Do
						End If
					End If
				Loop 
				If (NewCAS <> "") Then
					NewCAS = VB6.Format(CInt(Val(NewCAS)), "0")
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Name").Value = NewName
				If (NewCAS <> "") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("CAS").Value = NewCAS
				Else
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("CAS").Value = System.DBNull.Value
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				RecordCount_Chemicals = 1
				RecordCount_Isotherms_CAS = 0
				RecordCount_Isotherms_Name = 0
				'MODIFY ISOTHERMS WITH SAME CAS#/CHEMICAL NAME.
				If (DummyBool1) Then
					'MODIFY CAS NUMBER IN ISOTHERM RECORDS.
					If (OldCAS <> "") Then
						Current_Criteria = "select * from [Isotherms] where " & "[Component Number] = " & OldCAS
					Else
						Current_Criteria = "select * from [Isotherms] where " & "[Component Number] = Null"
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
					On Error Resume Next
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveFirst()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveLast()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveFirst()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					NumRecords = Rs1.RecordCount
					On Error GoTo err_mnuChemicalItem_Click
					If (NumRecords = 0) Then
						'DO NOTHING.
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Do Until Rs1.EOF
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Edit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.Edit()
							If (NewCAS <> "") Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Rs1("Component Number").Value = NewCAS
							Else
								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Rs1("Component Number").Value = System.DBNull.Value
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.Update()
							RecordCount_Isotherms_CAS = RecordCount_Isotherms_CAS + 1
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.MoveNext()
						Loop 
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Close()
				End If
				If (DummyBool2) Then
					'MODIFY CHEMICAL NAME IN ISOTHERM RECORDS.
					Current_Criteria = "select * from [Isotherms] where " & "[Name] = " & Chr(34) & OldName & Chr(34)
					'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
					On Error Resume Next
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveFirst()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveLast()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveFirst()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					NumRecords = Rs1.RecordCount
					On Error GoTo err_mnuChemicalItem_Click
					If (NumRecords = 0) Then
						'DO NOTHING.
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Do Until Rs1.EOF
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Edit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.Edit()
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1("Name").Value = NewName
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.Update()
							RecordCount_Isotherms_Name = RecordCount_Isotherms_Name + 1
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.MoveNext()
						Loop 
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Close()
				End If
				'REDISPLAY WINDOW.
				Call populate_lstCompo()
				Call lstCompo_SelectedIndexChanged(lstCompo, New System.EventArgs())
				'DISPLAY SUMMARY.
				Call Show_Message("Modification Summary:" & vbCrLf & vbCrLf & "Total Chemical Records Changed: " & Trim(Str(RecordCount_Chemicals)) & vbCrLf & "CAS Number Modified For: " & Trim(Str(RecordCount_Isotherms_CAS)) & " Isotherm Record" & IIf(RecordCount_Isotherms_CAS = 1, "", "s") & vbCrLf & "Chemical Name Modified For: " & Trim(Str(RecordCount_Isotherms_Name)) & " Isotherm Record" & IIf(RecordCount_Isotherms_Name = 1, "", "s"))
			Case 3 'delete current /////////////////////////////////////////////////////////////////////////////////////////////////
				'START THE EDIT PROCESS.
				Current_Criteria = "select * from [Chemicals] where " & "[Compo ID] = " & Trim(Str(THIS_CHEMICAL_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveFirst()


				'DO USER INPUT.
				'UPGRADE_WARNING: Untranslated statement in mnuChemicalItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuChemicalItem_Click. Please check source code.

				'DO USER INPUT.
				NewName = Database_Get_String(Rs1, "Name")
				NewCAS = Database_Get_String(Rs1, "CAS")

				DummyBool1 = True
				DummyBool2 = True
				Do While (1 = 1)
					Call frmEditIsothermCAS.frmEditIsothermCAS_Run("Deleting a Chemical", "Delete CAS Number", "Delete Chemical Name", "*", "*", "Delete all isotherms with the same CAS number", "Delete all isotherms with the same chemical name", "&Delete", USER_HIT_CANCEL, NewCAS, NewName, DummyStr1, DummyStr2, DummyBool1, DummyBool2)
					If (USER_HIT_CANCEL) Then
						Exit Sub
					End If
					NewName = Trim(NewName)
					NewCAS = Trim(NewCAS)
					If (NewName = "") Then
						Call Show_Error("The chemical name must be a non-blank string.")
					Else
						If (NewCAS = "") And (DummyBool1) Then
							Call Show_Error("You cannot automatically delete " & "all isotherm records currently without CAS numbers.")
						Else
							Exit Do
						End If
					End If
				Loop 
				If (NewCAS <> "") Then
					NewCAS = VB6.Format(CInt(Val(NewCAS)), "0")
				End If
				If (NewCAS <> "") Then
					Current_Criteria = "select * from [Chemicals] where " & "[Name] = " & Chr(34) & NewName & Chr(34) & " and [CAS] = " & NewCAS
				Else
					Current_Criteria = "select * from [Chemicals] where " & "[Name] = " & Chr(34) & NewName & Chr(34)
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				On Error Resume Next
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveFirst()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveLast()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveFirst()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NumRecords = Rs1.RecordCount
				On Error GoTo err_mnuChemicalItem_Click
				RecordCount_Chemicals = 0
				If (NumRecords = 0) Then
					'DO NOTHING.
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Do Until Rs1.EOF
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Rs1.Delete()
						RecordCount_Chemicals = RecordCount_Chemicals + 1
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Rs1.MoveNext()
					Loop 
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				RecordCount_Isotherms_CAS = 0
				RecordCount_Isotherms_Name = 0
				'DELETE ISOTHERMS WITH SAME CAS#/CHEMICAL NAME.
				If (DummyBool1) Then
					'DELETE ISOTHERM RECORDS WITH THIS CAS NUMBER.
					If (NewCAS <> "") Then
						Current_Criteria = "select * from [Isotherms] where " & "[Component Number] = " & NewCAS
					Else
						Current_Criteria = "select * from [Isotherms] where " & "[Component Number] = Null"
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
					On Error Resume Next
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveFirst()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveLast()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveFirst()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					NumRecords = Rs1.RecordCount
					On Error GoTo err_mnuChemicalItem_Click
					If (NumRecords = 0) Then
						'DO NOTHING.
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Do Until Rs1.EOF
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.Delete()
							RecordCount_Isotherms_CAS = RecordCount_Isotherms_CAS + 1
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.MoveNext()
						Loop 
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Close()
				End If
				If (DummyBool2) Then
					'DELETE ISOTHERM RECORDS WITH THIS CHEMICAL NAME.
					Current_Criteria = "select * from [Isotherms] where " & "[Name] = " & Chr(34) & NewName & Chr(34)
					'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
					On Error Resume Next
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveFirst()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveLast()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveFirst()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					NumRecords = Rs1.RecordCount
					On Error GoTo err_mnuChemicalItem_Click
					If (NumRecords = 0) Then
						'DO NOTHING.
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Do Until Rs1.EOF
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.Delete()
							RecordCount_Isotherms_Name = RecordCount_Isotherms_Name + 1
							'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Rs1.MoveNext()
						Loop 
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Close()
				End If
				'REDISPLAY WINDOW.
				Call populate_lstCompo()
				'DISPLAY SUMMARY.
				Call Show_Message("Modification Summary:" & vbCrLf & vbCrLf & "Total Chemical Records Deleted: " & Trim(Str(RecordCount_Chemicals)) & vbCrLf & "Isotherm Records Deleted with Matching CAS Number: " & Trim(Str(RecordCount_Isotherms_CAS)) & " Isotherm Record" & IIf(RecordCount_Isotherms_CAS = 1, "", "s") & vbCrLf & "Isotherm Records Deleted with Matching Chemical Name: " & Trim(Str(RecordCount_Isotherms_Name)) & " Isotherm Record" & IIf(RecordCount_Isotherms_Name = 1, "", "s"))
		End Select
		Exit Sub
exit_err_mnuChemicalItem_Click: 
		Exit Sub
err_mnuChemicalItem_Click: 
		Call Show_Trapped_Error("mnuChemicalItem_Click")
		Resume exit_err_mnuChemicalItem_Click
		
	End Sub
	Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
		'Me.Close()
		Me.Dispose()
		Exit Sub
	End Sub
	Public Sub mnuIsothermItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuIsothermItem.Click
		Dim Index As Short = mnuIsothermItem.GetIndex(eventSender)
		Dim USER_HIT_CANCEL As Boolean
		Dim USER_HIT_SAVE As Boolean
		Dim USER_HIT_SAVEASNEW As Boolean
		Dim THIS_CHEM_ID As Integer
		Dim THIS_ISOTHERM_ID As Integer
		Dim Current_Criteria As String
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim NumRecords As Integer
		Dim DEFAULT_PHASE_IS_LIQUID As Boolean
		Const INDEX_NEW As Short = 1
		Const INDEX_EDIT As Short = 2
		Const INDEX_DELETE As Short = 3
		Const INDEX_DELETE_ALL As Short = 4
		Dim NewName As String
		Dim msg As String
		Dim RetVal As Short
		Dim Select_Index As Short
		Dim i As Short
		Dim ThisName As String
		Dim ThisCAS As String
		Dim DEFAULT_CHEMICALNAME As String
		Dim DEFAULT_CHEMICALCAS As String
		Dim IsothermCount As Short
		Dim RecordCount_Deleted As Short
		On Error GoTo err_mnuIsothermItem_Click
		If (lstCompo.SelectedIndex < 0) Or (lstCompo.Items.Count = 0) Then
			Call Show_Error("You must first select a chemical.")
			Exit Sub
		End If
		THIS_CHEM_ID = VB6.GetItemData(lstCompo, lstCompo.SelectedIndex)
		'GET CHEMICAL NAME AND CAS FOR THIS CHEMICAL.
		Current_Criteria = "select * from [Chemicals] " & "where [Compo ID]=" & Trim(Str(THIS_CHEM_ID)) & " " & "order by [Name]"
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveLast()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NumRecords = Rs1.RecordCount
		On Error GoTo err_mnuIsothermItem_Click
		If (NumRecords = 0) Then
			'NO RECORD(S) AVAILABLE; WEIRD PROBLEM.
			'EXIT SUBROUTINE.
			Exit Sub
		End If

		'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
		'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
		ThisName = Database_Get_String(Rs1, "Name")
		ThisCAS = Database_Get_String(Rs1, "CAS")


		If (Index = INDEX_EDIT) Or (Index = INDEX_DELETE) Then
			If (lstRange(0).SelectedIndex < 0) Or (lstRange(0).Items.Count = 0) Then
				Call Show_Error("You must first select an isotherm.")
				Exit Sub
			End If
			THIS_ISOTHERM_ID = VB6.GetItemData(lstRange(0), lstRange(0).SelectedIndex)
			'POSITION TO CURRENT ISOTHERM RECORD.
			Current_Criteria = "select * from [Isotherms] " & "where [ID]=" & Trim(Str(THIS_ISOTHERM_ID)) & " " & "order by [Name]"
			'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
			On Error Resume Next
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1.MoveFirst()
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1.MoveLast()
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1.MoveFirst()
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NumRecords = Rs1.RecordCount
			On Error GoTo err_mnuIsothermItem_Click
			If (NumRecords = 0) Then
				'NO RECORD(S) AVAILABLE; WEIRD PROBLEM.
				'EXIT SUBROUTINE.
				Exit Sub
			End If
		End If
		Select Case Index
			Case INDEX_NEW '//////// NEW ISOTHERM. /////////////////////////////////////////////////////////////////////////////////////
				'DETERMINE DEFAULT PHASE.
				If (Bed.Phase = 0) Then
					DEFAULT_PHASE_IS_LIQUID = True
				Else
					DEFAULT_PHASE_IS_LIQUID = False
				End If
				'ALLOW USER TO ADD NEW RECORD.
				DEFAULT_CHEMICALNAME = ThisName
				DEFAULT_CHEMICALCAS = ThisCAS
				Call frmEditIsothermData.frmEditIsothermData_AddNew(DEFAULT_PHASE_IS_LIQUID, DEFAULT_CHEMICALNAME, DEFAULT_CHEMICALCAS, USER_HIT_CANCEL, USER_HIT_SAVE)
				If (USER_HIT_CANCEL) Then Exit Sub
				'ADD THE NEW ISOTHERM RECORD.
				Current_Criteria = "select * from [Isotherms]"
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'THE FIELD [ID] IS AUTOMATICALLY UPDATED.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Name").Value = frmEditIsothermData_Record.Name
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("K").Value = frmEditIsothermData_Record.k
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("1/n").Value = frmEditIsothermData_Record.OneOverN
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("C min").Value = frmEditIsothermData_Record.Cmin
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("C max").Value = frmEditIsothermData_Record.Cmax
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("pH min").Value = frmEditIsothermData_Record.pHmin
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("pH max").Value = frmEditIsothermData_Record.pHmax
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Source").Value = frmEditIsothermData_Record.Source
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("CarbonName").Value = frmEditIsothermData_Record.CarbonName
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Tmin").Value = frmEditIsothermData_Record.Tmin
				If (Trim(frmEditIsothermData_Record.CAS) <> "") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(Component Number). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Component Number").Value = CDbl(Val(frmEditIsothermData_Record.CAS))
				Else
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Component Number").Value = System.DBNull.Value
				End If
				If (frmEditIsothermData_Record.PhaseIsLiquid) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Phase").Value = "Liquid"
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Phase").Value = "Gas"
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Comments").Value = frmEditIsothermData_Record.Comments

				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				THIS_ISOTHERM_ID = Database_Get_Long(Rs1, "ID")


				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'CLOSE THE DATABASE.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'UPDATE WINDOW.
				Call lstCompo_SelectedIndexChanged(lstCompo, New System.EventArgs())
				'SELECT THE NEW ISOTHERM.
				Select_Index = 0
				For i = 0 To lstRange(0).Items.Count - 1
					If (VB6.GetItemData(lstRange(0), i) = THIS_ISOTHERM_ID) Then
						Select_Index = i
						Exit For
					End If
				Next i
				If (lstRange(0).Items.Count > 0) Then
					lstRange(0).SelectedIndex = Select_Index
				End If
			Case INDEX_EDIT '//////// EDIT ISOTHERM. //////////////////////////////////////////////////////////////////////////////////////////////
				'TRANSFER DATABASE RECORD FIELDS TO LOCAL MEMORY.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.


				'TRANSFER DATABASE RECORD FIELDS TO LOCAL MEMORY.
				If (Database_Get_String(Rs1, "Phase") = "Liquid") Then
					frmEditIsothermData_Record.PhaseIsLiquid = True
				Else
					frmEditIsothermData_Record.PhaseIsLiquid = False
				End If
				frmEditIsothermData_Record.Name = Database_Get_String(Rs1, "Name")
				frmEditIsothermData_Record.k = Database_Get_Double(Rs1, "K")
				frmEditIsothermData_Record.OneOverN = Database_Get_Double(Rs1, "1/n")
				frmEditIsothermData_Record.Cmin = Database_Get_Double(Rs1, "C min")
				frmEditIsothermData_Record.Cmax = Database_Get_Double(Rs1, "C max")
				frmEditIsothermData_Record.pHmin = Database_Get_Double(Rs1, "pH min")
				frmEditIsothermData_Record.pHmax = Database_Get_Double(Rs1, "pH max")
				frmEditIsothermData_Record.Source = Database_Get_String(Rs1, "Source")
				frmEditIsothermData_Record.CarbonName = Database_Get_String(Rs1, "CarbonName")
				frmEditIsothermData_Record.Tmin = Database_Get_Double(Rs1, "Tmin")
				frmEditIsothermData_Record.CAS = Database_Get_String(Rs1, "Component Number")
				frmEditIsothermData_Record.Comments = Database_Get_String(Rs1, "Comments")








				'ALLOW USER TO EDIT THIS RECORD.
				Call frmEditIsothermData.frmEditIsothermData_Edit(USER_HIT_CANCEL, USER_HIT_SAVE, USER_HIT_SAVEASNEW)
				If (USER_HIT_CANCEL) Then Exit Sub
				'SAVE THE EDITED ISOTHERM RECORD.
				Current_Criteria = "select * from [Isotherms] " & "where [ID]=" & Trim(Str(THIS_ISOTHERM_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				If (USER_HIT_SAVE) Then
					'MODIFY EXISTING RECORD.
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Edit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Edit()
					'KEEP ORIGINAL [ID] FIELD INTACT.
				End If
				If (USER_HIT_SAVEASNEW) Then
					'GENERATE NEW RECORD.
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.AddNew()
					'THE FIELD [ID] IS AUTOMATICALLY CREATED DURING THE .Update COMMAND.
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Name").Value = frmEditIsothermData_Record.Name
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("K").Value = frmEditIsothermData_Record.k
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("1/n").Value = frmEditIsothermData_Record.OneOverN
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("C min").Value = frmEditIsothermData_Record.Cmin
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("C max").Value = frmEditIsothermData_Record.Cmax
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("pH min").Value = frmEditIsothermData_Record.pHmin
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("pH max").Value = frmEditIsothermData_Record.pHmax
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Source").Value = frmEditIsothermData_Record.Source
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("CarbonName").Value = frmEditIsothermData_Record.CarbonName
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Tmin").Value = frmEditIsothermData_Record.Tmin
				If (Trim(frmEditIsothermData_Record.CAS) <> "") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(Component Number). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Component Number").Value = CDbl(Val(frmEditIsothermData_Record.CAS))
				Else
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Component Number").Value = System.DBNull.Value
				End If
				If (frmEditIsothermData_Record.PhaseIsLiquid) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Phase").Value = "Liquid"
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("Phase").Value = "Gas"
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("Comments").Value = frmEditIsothermData_Record.Comments
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'CLOSE THE DATABASE.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'UPDATE WINDOW.
				Call lstCompo_SelectedIndexChanged(lstCompo, New System.EventArgs())
			Case INDEX_DELETE '//////// DELETE ISOTHERM. /////////////////////////////////////////////////////////////////////////

				'UPGRADE_WARNING: Untranslated statement in mnuIsothermItem_Click. Please check source code.
				msg = "Isotherm Selected for Deletion:" & vbCrLf &
		  vbCrLf &
		  "    K = " & NumberToMFBString(Database_Get_Double(Rs1, "K")) & vbCrLf &
		  "    1/n = " & NumberToMFBString(Database_Get_Double(Rs1, "1/n")) & vbCrLf &
		  "    Carbon Name = " & Database_Get_String(Rs1, "CarbonName") & vbCrLf &
		  vbCrLf &
		  "Do you really want to delete this isotherm ?"


				RetVal = MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.YesNo, AppName_For_Display_Long)
				If RetVal = MsgBoxResult.No Then Exit Sub
				'PERFORM DELETION.
				Current_Criteria = "select * from [Isotherms] where " & "[ID] = " & Trim(Str(THIS_ISOTHERM_ID))
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Delete()
				'CLOSE THE DATABASE.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'UPDATE WINDOW.
				Call lstCompo_SelectedIndexChanged(lstCompo, New System.EventArgs())
			Case INDEX_DELETE_ALL '//////// DELETE ALL ISOTHERMS. /////////////////////////////////////////////////////////////////////////
				'DETERMINE ISOTHERM COUNT.
				If (ThisCAS <> "") Then
					Current_Criteria = "select * from [Isotherms] " & "where [Name]=" & Chr(34) & Trim(ThisName) & Chr(34) & " and [Component Number]=" & Trim(ThisCAS)
				Else
					Current_Criteria = "select * from [Isotherms] " & "where [Name]=" & Chr(34) & Trim(ThisName) & Chr(34) & " and [Component Number]=Null"
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				On Error Resume Next
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveFirst()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveLast()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveFirst()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NumRecords = Rs1.RecordCount
				On Error GoTo err_mnuIsothermItem_Click
				If (NumRecords = 0) Then
					'NO RECORD(S) AVAILABLE.
					Call Show_Error("There are no isotherms to delete for this chemical and CAS number.")
					Exit Sub
				Else
					IsothermCount = NumRecords
				End If
				'CHECK WITH USER: "ARE YOU SURE?"
				msg = "Isotherms Selected for Deletion:" & vbCrLf & vbCrLf & "    Chemical Name = " & ThisName & vbCrLf & "    CAS = " & ThisCAS & vbCrLf & vbCrLf & "    Total = " & Trim(Str(IsothermCount)) & " Isotherm Record" & IIf(IsothermCount = 1, "", "s") & vbCrLf & vbCrLf & "Do you really want to delete these isotherms ?"
				RetVal = MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.YesNo, AppName_For_Display_Long)
				If RetVal = MsgBoxResult.No Then Exit Sub
				'PERFORM DELETION.
				'USE CRITERIA "Current_Criteria" FROM ABOVE.
				'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1 = DB_Isotherm.OpenRecordset(Current_Criteria)
				RecordCount_Deleted = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Do Until Rs1.EOF
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Delete()
					RecordCount_Deleted = RecordCount_Deleted + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.MoveNext()
				Loop 
				'CLOSE THE DATABASE.
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Close()
				'UPDATE WINDOW.
				Call lstCompo_SelectedIndexChanged(lstCompo, New System.EventArgs())
				'DISPLAY SUMMARY.
				Call Show_Message("Modification Summary:" & vbCrLf & vbCrLf & "Total Isotherm Records Deleted: " & Trim(Str(RecordCount_Deleted)))
		End Select
		Exit Sub
exit_err_mnuIsothermItem_Click: 
		Exit Sub
err_mnuIsothermItem_Click: 
		Call Show_Trapped_Error("mnuIsothermItem_Click")
		Resume exit_err_mnuIsothermItem_Click
	End Sub


	Private Sub optSort_Click(ByRef Index As Short, ByRef Value As Short)
		If (Me.lstCompo.SelectedIndex < 0) Or (Me.lstCompo.Items.Count <= 0) Then
			Exit Sub
		End If
		Call populate_lstCompo()
	End Sub

	Private Sub _optSort_0_CheckedChanged(sender As Object, e As EventArgs) Handles _optSort_0.CheckedChanged
		Call optSort_Click(0, 0)
	End Sub

	Private Sub _optSort_1_CheckedChanged(sender As Object, e As EventArgs) Handles _optSort_1.CheckedChanged
		Call optSort_Click(1, 0)
	End Sub

	Private Sub _cmdFind_0_ClickEvent(sender As Object, e As EventArgs)
		Call cmdFind_Click(0)
	End Sub

	Private Sub _cmdFind_1_ClickEvent(sender As Object, e As EventArgs)
		Call cmdFind_Click(1)
	End Sub

	Private Sub cmdSelect_ClickEvent(sender As Object, e As EventArgs)
		Call Do_Select_Component(lstCompo.SelectedIndex)
	End Sub

	Private Sub Find_Click(sender As Object, e As EventArgs) Handles _cmdFind_0.Click
		Call cmdFind_Click(0)
	End Sub

	Private Sub FindAgain_Click(sender As Object, e As EventArgs) Handles _cmdFind_1.Click
		Call cmdFind_Click(1)
	End Sub

	Private Sub Select_Click(sender As Object, e As EventArgs) Handles cmdSelect.Click
		Call Do_Select_Component(lstCompo.SelectedIndex)
	End Sub

	Private Sub frmEditIsotherm_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)
	End Sub
End Class