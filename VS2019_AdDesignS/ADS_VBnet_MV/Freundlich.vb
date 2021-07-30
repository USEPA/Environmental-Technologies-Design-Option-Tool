Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports VB = Microsoft.VisualBasic
Friend Class frmFreundlich
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	'Dim FORM_MODE As Integer
	'Const FORM_MODE_ADDNEW = 1
	'Const FORM_MODE_EDIT = 2
	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_OK As Boolean
	Dim frmFreundlich_Is_Dirty As Boolean
	Dim ActivatedYet As Boolean

	Dim HALT_OPTFREUNDLICHSOURCE As Boolean
	Dim HALT_LSTCOMPO As Boolean
	Dim HALT_LSTRANGE As Boolean
	Dim HALT_CBOMETHOD As Boolean
	Const CBOSORTMETHOD_NAME As Short = 1
	Const CBOSORTMETHOD_CAS As Short = 2

	'UPGRADE_WARNING: Arrays in structure SaveOldComponent may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim SaveOldComponent As ComponentPropertyType

	'UPGRADE_ISSUE: Database object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Dim DB_Isotherm As dao.Database
	Dim Find_String As String




	Const frmFreundlich_declarations_end As Boolean = True


	Sub frmFreundlich_Run(ByRef OUTPUT_Raise_Dirty_Flag As Boolean)
		On Error GoTo err_frmFreundlich_Run
		'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
		'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
		'UPGRADE_WARNING: Couldn't resolve default property of object Ws1.OpenDatabase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DB_Isotherm = Ws1.OpenDatabase(fn_DB_Isotherm, True, False, ";pwd=" & decrypt_string(Encrypted_User_Password))
		'		DB_Isotherm = DAOEngine.OpenDatabase(fn_DB_Isotherm)  'From Ws1 to DaoEngine ??? Shang
		'Set DB_Isotherm = ws1.OpenDatabase(fn_DB_Isotherm)
		Me.ShowDialog()
		If (USER_HIT_OK) Then
			OUTPUT_Raise_Dirty_Flag = True
		Else
			OUTPUT_Raise_Dirty_Flag = False
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object DB_Isotherm.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DB_Isotherm.Close()
		Exit Sub
exit_err_frmFreundlich_Run:
		Exit Sub
err_frmFreundlich_Run:
		Call Show_Trapped_Error("frmFreundlich_Run")
		OUTPUT_Raise_Dirty_Flag = False
		Resume exit_err_frmFreundlich_Run
	End Sub
	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdCancelOK_1.Enabled = False
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCalculate.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdCalculate.Enabled = False
		End If
	End Sub


	Sub frmFreundlich_GenericStatus_Set(ByRef fn_Text As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ToolStripStatus.Text = fn_Text
	End Sub
	Sub frmFreundlich_DirtyStatus_Set(ByRef newVal As Boolean)
		If (newVal) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ToolStripDirty.Text = "Data Changed"
			'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ToolStripDirty.ForeColor = Color.FromArgb(QBColor(12))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ToolStripDirty.Text = "Unchanged"
			'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ToolStripDirty.ForeColor = Color.FromArgb(QBColor(0))
		End If
	End Sub
	Sub frmFreundlich_DirtyStatus_Set_Current()
		Call frmFreundlich_DirtyStatus_Set(frmFreundlich_Is_Dirty)
	End Sub
	Sub frmFreundlich_DirtyStatus_Throw()
		frmFreundlich_Is_Dirty = True
		Call frmFreundlich_DirtyStatus_Set_Current()
	End Sub
	Sub frmFreundlich_DirtyStatus_Clear()
		frmFreundlich_Is_Dirty = False
		Call frmFreundlich_DirtyStatus_Set_Current()
	End Sub


	Sub populate_lstCompo()
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim Current_Criteria As String
		Dim SAVE_CURRENT_POSITION As Integer
		Dim NEW_LISTINDEX As Short
		Dim This_ID As Integer
		Dim NumRecords As Integer
		Dim TempStr As New VB6.FixedLengthString(15)
		Dim Output_Line As String
		Dim PhaseCode As String
		Dim SortCode As String
		Dim ThisChemicalName As String
		Dim ThisChemicalCAS As String
		Dim LastChemicalName As String
		Dim LastChemicalCAS As String
		HALT_LSTCOMPO = True
		On Error GoTo err_populate_lstCompo
		'
		' SAVE CURRENT POSITION.
		'
		If (lstCompo.Items.Count > 0) And (lstCompo.SelectedIndex >= 0) Then
			SAVE_CURRENT_POSITION = VB6.GetItemData(lstCompo, lstCompo.SelectedIndex)
		Else
			SAVE_CURRENT_POSITION = -1
		End If
		'
		' SET UP SEARCH CRITERIA.
		'
		Select Case Bed.Phase
			Case 0 : PhaseCode = "Liquid"
			Case 1 : PhaseCode = "Gas"
		End Select
		Select Case VB6.GetItemData(cboSortMethod, cboSortMethod.SelectedIndex)
			Case CBOSORTMETHOD_NAME : SortCode = "Name, [Component Number]"
			Case CBOSORTMETHOD_CAS : SortCode = "[Component Number], Name"
		End Select
		'Current_Criteria = "select * from [Chemicals] " & _
		''    "order by [Name]"
		Current_Criteria = "select * from Isotherms" & " where Phase = '" & PhaseCode & "'" & " order by " & SortCode
		'
		' START SEARCH.
		'
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
		'
		' POPULATE LISTBOX.
		'
		lstCompo.Items.Clear()
		If (NumRecords = 0) Then
			' NO RECORDS AVAILABLE.
			lstCompo.Visible = False
			lblEmpty_lstCompo.SetBounds(lstCompo.Left, lstCompo.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			lblEmpty_lstCompo.Visible = True
		Else
			'
			' DISPLAY RECORDS.
			'
			lstCompo.Visible = True
			lblEmpty_lstCompo.Visible = False
			NEW_LISTINDEX = -1
			LastChemicalName = "n/a yet"
			LastChemicalCAS = "n/a yet"
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Do Until Rs1.EOF

				'UPGRADE_WARNING: Untranslated statement in populate_lstCompo. Please check source code.
				'UPGRADE_WARNING: Untranslated statement in populate_lstCompo. Please check source code.
				ThisChemicalName = Database_Get_String(Rs1, "Name")
				ThisChemicalCAS = Database_Get_String(Rs1, "Component Number")





				If ((ThisChemicalName <> LastChemicalName) Or (ThisChemicalCAS <> LastChemicalCAS)) Then
					'
					' THIS "IF" STATEMENT EXISTS IN ORDER TO ELIMINATE DUPLICATE
					' CHEMICALS FROM THE LIST.  THIS CODE DEPENDS COMPLETELY
					' ON THE FACT THAT THE LIST IS SORTED !
					'
					LastChemicalCAS = ThisChemicalCAS
					LastChemicalName = ThisChemicalName
					TempStr.Value = ThisChemicalCAS
					'THIS STRING IS ENSURED TO BE 15 CHARACTERS LONG.
					Output_Line = TempStr.Value & " " & ThisChemicalName
					NEW_LISTINDEX = lstCompo.Items.Add(Output_Line)

					'UPGRADE_WARNING: Untranslated statement in populate_lstCompo. Please check source code.
					This_ID = Database_Get_Long(Rs1, "ID")



					'UPGRADE_ISSUE: ListBox property lstCompo.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					VB6.SetItemData(lstCompo, NEW_LISTINDEX, This_ID)
					If (SAVE_CURRENT_POSITION <> -1) Then
						If (SAVE_CURRENT_POSITION = This_ID) Then
							'UPGRADE_ISSUE: ListBox property lstCompo.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
							NEW_LISTINDEX = lstCompo.SelectedIndex
						End If
					End If
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.MoveNext()
			Loop
			If (lstCompo.Items.Count > 0) And (NEW_LISTINDEX > -1) Then
				HALT_LSTCOMPO = True
				lstCompo.SelectedIndex = NEW_LISTINDEX
				HALT_LSTCOMPO = False
			End If
		End If
		'
		' CLOSE DATABASE AND EXIT.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		HALT_LSTCOMPO = False
		Exit Sub
exit_err_populate_lstCompo:
		HALT_LSTCOMPO = False
		Exit Sub
err_populate_lstCompo:
		Call Show_Trapped_Error("populate_lstCompo")
		Resume exit_err_populate_lstCompo
	End Sub
	Sub populate_lstRange(ByRef ThisCAS As String, ByRef ThisChemical As String)
		Dim PHASE_CODE As Short
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
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
		'GET PHASE CODE.
		Select Case Bed.Phase
			Case 0 : PhaseCode = "Liquid"
			Case 1 : PhaseCode = "Gas"
		End Select
		'SAVE CURRENT POSITION.
		If (lstRange(0).Items.Count > 0) And (lstRange(0).SelectedIndex >= 0) Then
			SAVE_CURRENT_POSITION = VB6.GetItemData(lstRange(0), lstRange(0).SelectedIndex)
		Else
			SAVE_CURRENT_POSITION = -1
		End If
		'SET UP SEARCH CRITERIA.
		If (Trim(ThisCAS) = "0") Then ThisCAS = ""
		If (Trim(ThisCAS) <> "") Then
			Current_Criteria = "select * from Isotherms" & " where Phase = '" & PhaseCode & "'" & " and [Component Number] = " & Trim(ThisCAS) & " order by CarbonName"
		Else
			Current_Criteria = "select * from Isotherms" & " where Phase = '" & PhaseCode & "'" & " and Name = '" & Trim(ThisChemical) & "'" & " order by CarbonName"
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
			grpTwo.Text = "No Isotherms Available."
		Else
			'DISPLAY RECORDS.
			'UPGRADE_WARNING: Couldn't resolve default property of object fraTwo.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			grpTwo.Text = Trim(Str(NumRecords)) & " " & PhaseCode & " Phase Isotherm" & IIf(NumRecords = 1, "", "s") & " Available"
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
				lstRange(0).Items.Add(New VB6.ListBoxItem(ThisOutput, This_ID))
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
						NEW_LISTINDEX = lstRange(0).SelectedIndex
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
			Call AssignCaptionAndTag(lblValue(2), Database_Get_String(Rs1, "CarbonName"))
			Call AssignCaptionAndTag(lblTemperature, Database_Get_Double(Rs1, "Tmin"))
			Call AssignCaptionAndTag(lblPhase, Database_Get_String(Rs1, "Phase"))
			Call AssignCaptionAndTag(lblValue(3), Database_Get_String(Rs1, "Source"))
			Call AssignCaptionAndTag(lblComments, Database_Get_String(Rs1, "Comments"))
			Component(0).IsothermDB_OneOverN = Database_Get_Double(Rs1, "1/n")
			Component(0).IsothermDB_K = Database_Get_Double(Rs1, "K")

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


	Sub Populate_cboSortMethod()
		cboSortMethod.Items.Clear()
		cboSortMethod.Items.Add(New VB6.ListBoxItem("By Name", CBOSORTMETHOD_NAME))
		cboSortMethod.Items.Add(New VB6.ListBoxItem("By CAS", CBOSORTMETHOD_CAS))
		cboSortMethod.SelectedIndex = 0
	End Sub
	Sub Populate_cboMethod()
		Dim NewTag As Short
		Dim i As Short
		HALT_CBOMETHOD = True
		'xaxaxa (12/10/97)
		Select Case Bed.Phase
			Case 0
				' ***** Phase = Water *****
				'lblInput(3).Visible = False
				'txtInput(3).Visible = False
				cboMethod.Items.Clear()
				cboMethod.Items.Add(New VB6.ListBoxItem("3 - Parameter Polanyi Isotherm Correlation", IPESMETHOD_LIQ_3PARAM))
				cboMethod.Items.Add(New VB6.ListBoxItem("D-R Uniform Adsorbate", IPESMETHOD_LIQ_DRUNIFORM))
				'cboMethod.AddItem "D-R Non-Uniform Adsorbate"
			Case 1
				' ***** Phase = Air *****
				'lblInput(3).Visible = True
				'txtInput(3).Visible = True
				cboMethod.Items.Clear()
				cboMethod.Items.Add(New VB6.ListBoxItem("D-R based on Spreading Pressure Eval.", IPESMETHOD_GAS_DRSPREADINGP))
				'EJO 12/16/97 -- This correlation killed off today.
				'cboMethod.AddItem "D-R Isotherm Correlation for RH < 50%"
				'If (check_internal_to_mtu()) Then
				'  cboMethod.AddItem "D-R based on Spreading Pressure Eval. (MTU only)"
				'End If
				'cboMethod.AddItem "Calgon BPL"
				'cboMethod.AddItem "D-R based on Spreading Pressure Evaluation"
				'EJO 8/5/97 -- These two correlations were killed off!
		End Select
		'PERFORM LOOKUP FOR CURRENT METHOD.
		'[NEW AS OF 8/28/98.]
		NewTag = 0
		For i = 0 To cboMethod.Items.Count - 1
			If (Component(0).IPES_EstimationMethod = VB6.GetItemData(cboMethod, i)) Then
				NewTag = i
				Exit For
			End If
		Next i
		cboMethod.SelectedIndex = NewTag
		HALT_CBOMETHOD = False
	End Sub
	Sub frmFreundlich_PopulateUnits()
		Call unitsys_register(Me, lblInput(5), txtInput(11), Nothing, "", "", "", "", "", 100.0#, False)
		Call unitsys_register(Me, lblInput(6), txtInput(12), Nothing, "", "", "", "", "", 100.0#, False)
		Call unitsys_register(Me, lblText(5), UserOneOverN, Nothing, "", "", "", "", "", 100.0#, False)
		Call unitsys_register(Me, lblText(4), UserK, Nothing, "", "", "", "", "", 100.0#, False)
	End Sub


	'UPGRADE_WARNING: Event cboMethod.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboMethod_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboMethod.SelectedIndexChanged
		Dim OldValue As Short
		Dim NewValue As Short
		If (HALT_CBOMETHOD) Then Exit Sub
		OldValue = Component(0).IPES_EstimationMethod
		NewValue = VB6.GetItemData(cboMethod, cboMethod.SelectedIndex)
		If (OldValue <> NewValue) Then
			Component(0).IPES_EstimationMethod = NewValue
			'THROW DIRTY FLAG.
			Call frmFreundlich_DirtyStatus_Throw()
		End If
	End Sub


	'UPGRADE_WARNING: Event cboSortMethod.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboSortMethod_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
		Call populate_lstCompo()
	End Sub


	Private Sub cmdCalculate_Click()
		Dim WhichModule As Short
		Dim INPUT_NL As Short
		Dim INPUT_OMAG As Double
		Dim Raise_Dirty_Flag As Boolean
		'CHECK FOR ZERO VALUES OF POLANYI PARAMETERS.
		If (Carbon.W0 = 0#) Or (Carbon.BB = 0#) Or (Carbon.PolanyiExponent = 0#) Then
			Call Show_Error("The Polanyi parameters have not been properly " & "specified.  To properly specify the Polanyi parameters, you must " & "enter a non-zero value for each of the following parameters: " & "W0, BB, and GM (Polanyi Exponent).  " & "Click on the button marked Edit Parameters to " & "make these changes.")
			Exit Sub
		End If
		Select Case VB6.GetItemData(cboMethod, cboMethod.SelectedIndex)
			Case IPESMETHOD_LIQ_3PARAM
				WhichModule = 1
			Case IPESMETHOD_GAS_DRSPREADINGP
				WhichModule = 4
			Case IPESMETHOD_LIQ_DRUNIFORM
				WhichModule = 5
			Case Else
				Call Show_Error("IPE calculation code #" & Trim(Str(VB6.GetItemData(cboMethod, cboMethod.SelectedIndex))) & " is invalid.  Select another method.")
				Exit Sub
		End Select
		INPUT_NL = CShort(Component(0).IPES_NumRegressionPts)
		INPUT_OMAG = CDbl(Component(0).IPES_OrderOfMagnitude)
		Call ModelIPE_Go(WhichModule, INPUT_NL, INPUT_OMAG, Raise_Dirty_Flag)
		If (Raise_Dirty_Flag) Then
			'THROW DIRTY FLAG.
			Call frmFreundlich_DirtyStatus_Throw()
		End If
		'REFRESH WINDOW.
		Call frmFreundlich_Refresh()
	End Sub


	Private Sub cmdCancelOK_Click(ByRef Index As Short)
		Dim WhichSelected As Short
		Select Case Index
			Case 0 'CANCEL.
				'ROLLBACK TO ORIGINAL COMPONENT DATA.
				'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Component(0) = SaveOldComponent
				'EXIT OUT OF HERE.
				USER_HIT_CANCEL = True
				USER_HIT_OK = False
				Me.Close()
				Exit Sub
			Case 1 'OK.
				'VERIFY THAT NEW K AND 1/N SOURCE IS VALID.
				'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CBool(RadioButton1.Checked) Then WhichSelected = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CBool(RadioButton2.Checked) Then WhichSelected = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CBool(RadioButton3.Checked) Then WhichSelected = 2
				Select Case WhichSelected
					Case 0
						'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource().Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (VB.Left(RadioButton1.Text, 1) = "(") Then
							'FORCE SOURCE TO USER-INPUT.
							Call Show_Error("Unable to validate isotherm " & "database as source of K and 1/n: reverting " & "to user-input as source of K and 1/n.")
							Component(0).Source_KandOneOverN = KNSOURCE_USERINPUT
						End If
					Case 1
						'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource().Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (VB.Left(RadioButton2.Text, 1) = "(") Then
							'FORCE SOURCE TO USER-INPUT.
							Call Show_Error("Unable to validate IPES as " & "source of K and 1/n: reverting to user-input " & "as source of K and 1/n.")
							Component(0).Source_KandOneOverN = KNSOURCE_USERINPUT
						End If
					Case 2
						'DO NOTHING!
				End Select
				'TRANSFER K AND 1/N DATA TO "USED" VARIABLES IN COMPONENT STRUCTURE.
				Select Case Component(0).Source_KandOneOverN
					Case KNSOURCE_ISOTHERMDB
						'ISOTHERM DATABASE.
						Component(0).Use_OneOverN = Component(0).IsothermDB_OneOverN
						Component(0).Use_K = Component(0).IsothermDB_K
					Case KNSOURCE_IPES
						'IPE CALCULATION.
						Component(0).Use_OneOverN = Component(0).IPESResult_OneOverN
						Component(0).Use_K = Component(0).IPESResult_K
					Case KNSOURCE_USERINPUT
						'USER INPUT.
						Component(0).Use_OneOverN = Component(0).UserEntered_OneOverN
						Component(0).Use_K = Component(0).UserEntered_K
				End Select
				'EXIT OUT OF HERE.
				USER_HIT_CANCEL = False
				USER_HIT_OK = True
				Me.Close()
				Exit Sub
		End Select
	End Sub


	Private Sub cmdEditPolanyi_Click()
		Dim Raise_Dirty_Flag As Boolean
		Call frmPolanyi.frmPolanyi_Edit(Me, Raise_Dirty_Flag)
		If (Raise_Dirty_Flag) Then
			'DO NOT THROW FREUNDLICH DIRTY FLAG; THROW MAIN DIRTY FLAG.
			'REASON: THE POLANYI PARAMETERS ARE _NOT_ SPECIFIC
			'TO THIS COMPONENT; THEY ARE SPECIFIC TO THE CARBON.
			'---------------------------------
			''THROW DIRTY FLAG.
			'Call frmFreundlich_DirtyStatus_Throw
			'THROW (MAIN) DIRTY FLAG.
			Call DirtyStatus_Throw()
		End If
	End Sub


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




	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub


	'UPGRADE_WARNING: Form event frmFreundlich.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmFreundlich_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Dim No_IsothermDBData As Boolean
		If (Not ActivatedYet) Then
			ActivatedYet = True
			'RE-POPULATE COMPONENT LIST FROM ISOTHERM DATABASE.
			Call populate_lstCompo()
			'SETUP SOURCE OF K AND 1/N.
			''''Debug.Print "xxxx1 " & Now
			HALT_OPTFREUNDLICHSOURCE = True
			Select Case Component(0).Source_KandOneOverN
				Case KNSOURCE_ISOTHERMDB
					'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'_optFreundlichSource_0.Value = True
					RadioButton1.Checked = True
				Case KNSOURCE_IPES
					'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'_optFreundlichSource_1.Value = True
					RadioButton2.Checked = True
				Case KNSOURCE_USERINPUT
					'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'_optFreundlichSource_2.Value = True
					RadioButton3.Checked = True
			End Select
			HALT_OPTFREUNDLICHSOURCE = False
			'
			'OUTLINE OF SEARCHES:
			'====================
			'1.) If user has already selected an isotherm from the DB,
			'    select that one on-screen, along with their selected
			'    pH/conc. range.
			'2.) Otherwise, if they have a CAS# for this component,
			'    search for the CAS#.  If found, select it on-screen.
			'3.) If the above two criteria fail, search for an exact
			'    match on the component name.  If found, select
			'    it on-screen.
			'
			'SEARCH #1:
			'==========
			No_IsothermDBData = True
			If (Trim(Component(0).IsothermDB_Component_Name) <> "") Then
				'----- Find the selected component: -----
				Find_String = Trim(Component(0).IsothermDB_Component_Name)
				If (Do_Search_For_Text(False) = True) Then
					'SEARCH SUCEEDED; SELECT THIS ISOTHERM.
					'NOTE: BY HALTING lstCompo_Click(), THE COMPOUND CAN
					'BE SELECTED WITHOUT CLEARING THE ISOTHERM DB VALUES
					'OF K AND 1/N.
					'Call cmdSelect_Click
					HALT_LSTCOMPO = True
					Call Do_Select_Component(lstCompo.SelectedIndex)
					HALT_LSTCOMPO = False
					If (Component(0).IsothermDB_K <> -1.0#) And (Component(0).IsothermDB_OneOverN <> -1.0#) Then
						'----- Find the selected pH/conc. range:
						If (Component(0).IsothermDB_Range_Num <> -1) Then
							'JUST BEING PARANOID ABOUT DATABASE CHANGES.
							If (Component(0).IsothermDB_Range_Num <= lstRange(0).Items.Count - 1) Then
								HALT_LSTRANGE = True
								lstRange(0).SelectedIndex = Component(0).IsothermDB_Range_Num
								lstRange(1).SelectedIndex = Component(0).IsothermDB_Range_Num
								HALT_LSTRANGE = False
								'Call lstRange_Click(0)
								Call populate_lblValue(VB6.GetItemData(lstRange(0), lstRange(0).SelectedIndex))
								No_IsothermDBData = False
							Else
								HALT_LSTRANGE = True
								lstRange(0).SelectedIndex = 0
								lstRange(1).SelectedIndex = 0
								HALT_LSTRANGE = False
								'Call lstRange_Click(0)
								Call populate_lblValue(VB6.GetItemData(lstRange(0), lstRange(0).SelectedIndex))
								No_IsothermDBData = False
							End If
						End If
					End If
				End If
			End If
			'
			'SEARCH #2:
			'==========
			If (No_IsothermDBData) Then
				If (Component(0).CAS <> 0) Then
					Find_String = Trim(Str(Component(0).CAS)) & "   "
					If (Do_Search_For_Text(False) = True) Then
						'SEARCH SUCEEDED; SELECT THIS ISOTHERM.
						Call lstCompo_SelectedIndexChanged(lstCompo, New System.EventArgs())
						No_IsothermDBData = False
					End If
				End If
			End If
			'
			'SEARCH #3:
			'==========
			If (No_IsothermDBData) Then
				Find_String = "   " & Trim(Component(0).Name)
				If (Do_Search_For_Text(False) = True) Then
					'SEARCH SUCEEDED; SELECT THIS ISOTHERM.
					Call lstCompo_SelectedIndexChanged(lstCompo, New System.EventArgs())
					No_IsothermDBData = False
				End If
			End If
			''
			''NONE OF THE ABOVE SUCCEEDED; K AND 1/N UNSPECIFIED SO FAR.
			''
			'If (No_IsothermDBData) Then
			'  Component(0).IsothermDB_K = -1#
			'  Component(0).IsothermDB_OneOverN = -1#
			'End If
			'
			'SOME OF THE SELECTIONS ABOVE MAY HAVE SET THE DIRTY FLAG.
			'THIS CODE CLEARS IT.
			'
			Call frmFreundlich_DirtyStatus_Clear()
			'REFRESH DISPLAY.
			Call frmFreundlich_Refresh()
			'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'_optFreundlichSource_0.Enabled = True
			'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'_optFreundlichSource_0.Enabled = True
			'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'_optFreundlichSource_0.Enabled = True
			RadioButton1.Enabled = True
			'
			'CLEAR HOURGLASS MOUSE POINTER.
			'
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		End If
	End Sub
	Private Sub frmFreundlich_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		'SAVE OLD COMPONENT FOR CANCEL ROLLBACK.
		'UPGRADE_WARNING: Couldn't resolve default property of object SaveOldComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SaveOldComponent = Component(0)
		'MISC INITS.
		ActivatedYet = False
		'	Me.Height = VB6.TwipsToPixelsY(7200)
		'	Me.Width = VB6.TwipsToPixelsX(9585)
		'Me.Height = 480
		'Me.Width = 640
		Call CenterOnForm(Me, frmMain)
		Find_String = ""
		lblWarning.Text = ""
		Call Populate_cboMethod()
		Call Populate_cboSortMethod()
		Me.Text = "Freundlich Isotherm Parameters for " & Trim(Component(0).Name)
		If (Bed.Phase = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fraIsothermDB.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			grpIsothermDB.Text = "Liquid Phase Isotherm Database"
			lblEstimationMethod.Text = "Liquid Phase Estimation Method:"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object fraIsothermDB.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			grpIsothermDB.Text = "Gas Phase Isotherm Database"
			lblEstimationMethod.Text = "Gas Phase Estimation Method:"
		End If
		Call Clear_lblValue()
		Call frmFreundlich_DirtyStatus_Clear()
		Call frmFreundlich_GenericStatus_Set("")
		'POPULATE UNIT CONTROLS.
		Call frmFreundlich_PopulateUnits()
		'DEMO SETTINGS.
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	Private Sub frmFreundlich_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
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
	Private Sub lstCompo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
		If (HALT_LSTCOMPO) Then Exit Sub
		HALT_LSTCOMPO = True
		Call Do_Select_Component(lstCompo.SelectedIndex)
		HALT_LSTCOMPO = False
		'CLEAR EXISTING RECORD DATA.
		Call Clear_lblValue()
		'INVALIDATE EXISTING ISOTHERM RECORD LINK (IF ANY).
		Component(0).IsothermDB_OneOverN = -1.0#
		Component(0).IsothermDB_K = -1.0#
		HALT_LSTRANGE = True
		lstRange(0).SelectedIndex = -1
		lstRange(1).SelectedIndex = -1
		HALT_LSTRANGE = False
		Call frmFreundlich_Refresh()
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
		'TRANSFER LINK TO COMPONENT(0) STRUCTURE.
		Component(0).IsothermDB_Component_Name = Trim(VB6.GetItemString(lstCompo, lstCompo.SelectedIndex))
		Component(0).IsothermDB_Range_Num = lstRange(0).SelectedIndex
		'DISPLAY ISOTHERM RECORD.
		Call populate_lblValue(VB6.GetItemData(lstRange(0), lstRange(0).SelectedIndex))
		'THROW DIRTY FLAG.
		Call frmFreundlich_DirtyStatus_Throw()
		'REFRESH WINDOW.
		Call frmFreundlich_Refresh()
		HALT_LSTRANGE = False
	End Sub


	Private Sub optFreundlichSource_Click(ByRef Index As Short, ByRef Value As Short)
		Dim KandOneOverN_Enabled As Short
		Dim X As Short
		Dim temp As String
		Dim WhichSelected As Short
		If (HALT_OPTFREUNDLICHSOURCE) Then Exit Sub
		'DETERMINE WHICH OPTION WAS SELECTED.
		'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool(RadioButton1.Checked) Then WhichSelected = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool(RadioButton2.Checked) Then WhichSelected = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object optFreundlichSource(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool(RadioButton3.Checked) Then WhichSelected = 2
		'Debug.Print "optFreundlichSource_Click; WhichSelected = " & _
		'Trim$(Str$(WhichSelected))
		'TRANSFER K AND 1/N TO "USED" VARIABLES IN COMPONENT STRUCTURE.
		Select Case WhichSelected
			Case 0
				Component(0).Source_KandOneOverN = KNSOURCE_ISOTHERMDB
				Component(0).Use_K = Component(0).IsothermDB_K
				Component(0).Use_OneOverN = Component(0).IsothermDB_OneOverN
				KandOneOverN_Enabled = False
			Case 1
				Component(0).Source_KandOneOverN = KNSOURCE_IPES
				Component(0).Use_K = Component(0).IPESResult_K
				Component(0).Use_OneOverN = Component(0).IPESResult_OneOverN
				KandOneOverN_Enabled = False
			Case 2
				Component(0).Source_KandOneOverN = KNSOURCE_USERINPUT
				Component(0).Use_K = Component(0).UserEntered_K
				Component(0).Use_OneOverN = Component(0).UserEntered_OneOverN
				KandOneOverN_Enabled = True
		End Select
		'REFRESH WINDOW.
		Call frmFreundlich_Refresh()
	End Sub



	Private Sub UCtl_GotFocus(ByRef Ctl As TextBox)
		Dim StatusMessagePanel As String
		Dim CtlIndex As Short
		Call unitsys_control_txtx_gotfocus(Ctl)
		CtlIndex = 0
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CtlIndex = Ctl.TabIndex
		On Error GoTo 0
		If (Trim(UCase(Ctl.Name)) = Trim(UCase("txtInput"))) And (CtlIndex = 11) Then
			StatusMessagePanel = "Type in the order of magnitude"
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtInput"))) And (CtlIndex = 12) Then
			StatusMessagePanel = "Type in the number of regression points"
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("UserOneOverN"))) Then
			StatusMessagePanel = "Type in the user-input Freundlich K value"
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("UserK"))) Then
			StatusMessagePanel = "Type in the user-input Freundlich 1/n value"
		Else
			'NOT RECOGNIZED -- DO NOTHING.
		End If
		Call frmFreundlich_GenericStatus_Set(StatusMessagePanel)
	End Sub
	Sub UCtl_LostFocus(ByRef Ctl As TextBox)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim CtlIndex As Short
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
		CtlIndex = 0
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CtlIndex = Ctl.TabIndex   ' from index
		On Error GoTo 0
		If (Trim(UCase(Ctl.Name)) = Trim(UCase("txtInput"))) And (CtlIndex = 11) Then
			Val_Low = 1.0# : Val_High = 10.0#
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtInput"))) And (CtlIndex = 12) Then
			Val_Low = 1.0# : Val_High = 10000.0#
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("UserOneOverN"))) Then
			Val_Low = 1.0E-40 : Val_High = 1.0E+40
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("UserK"))) Then
			Val_Low = 1.0E-40 : Val_High = 1.0E+40
		Else
			'NOT RECOGNIZED -- DO NOTHING.
		End If
		NewValue_Okay = False
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		Call frmFreundlich_GenericStatus_Set("")
		If (NewValue_Okay) Then
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				If (Trim(UCase(Ctl.Name)) = Trim(UCase("txtInput"))) And (CtlIndex = 11) Then
					Component(0).IPES_OrderOfMagnitude = NewValue
				ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtInput"))) And (CtlIndex = 12) Then
					Component(0).IPES_NumRegressionPts = CShort(NewValue)
				ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("UserOneOverN"))) Then
					Component(0).UserEntered_OneOverN = NewValue
				ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("UserK"))) Then
					Component(0).UserEntered_K = NewValue
				Else
					'NOT RECOGNIZED -- DO NOTHING.
				End If
				'RAISE DIRTY FLAG IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					'THROW DIRTY FLAG.
					Call frmFreundlich_DirtyStatus_Throw()
				End If
				'REFRESH WINDOW.
				Call frmFreundlich_Refresh()
			End If
		End If
	End Sub


	Private Sub txtInput_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInput.Enter
		Dim Index As Short = txtInput.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtInput(Index) : Call UCtl_GotFocus(Ctl)
	End Sub
	Private Sub txtInput_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInput.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtInput.GetIndex(eventSender)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtInput_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInput.Leave
		Dim Index As Short = txtInput.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtInput(Index) : Call UCtl_LostFocus(Ctl)
	End Sub

	Private Sub UserK_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UserK.Enter
		Dim Ctl As System.Windows.Forms.Control : Ctl = UserK : Call UCtl_GotFocus(Ctl)
	End Sub
	Private Sub UserK_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles UserK.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub UserK_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UserK.Leave
		Dim Ctl As System.Windows.Forms.Control : Ctl = UserK : Call UCtl_LostFocus(Ctl)
	End Sub

	Private Sub UserOneOverN_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UserOneOverN.Enter
		Dim Ctl As System.Windows.Forms.Control : Ctl = UserOneOverN : Call UCtl_GotFocus(Ctl)
	End Sub
	Private Sub UserOneOverN_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles UserOneOverN.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub UserOneOverN_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UserOneOverN.Leave
		Dim Ctl As System.Windows.Forms.Control : Ctl = UserOneOverN : Call UCtl_LostFocus(Ctl)
	End Sub


	Private Sub _cmdCancelOK_0_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_0.Click
		Call cmdCancelOK_Click(0)

	End Sub

	Private Sub _cmdCancelOK_1_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_1.Click
		Call cmdCancelOK_Click(1)

	End Sub

	Private Sub cmdEditPolanyi_Click(sender As Object, e As EventArgs) Handles cmdEditPolanyi.Click
		Call cmdEditPolanyi_Click()
	End Sub

	Private Sub _cmdFind_0_Click(sender As Object, e As EventArgs) Handles _cmdFind_0.Click
		Call cmdFind_Click(0)
	End Sub

	Private Sub _cmdFind_1_Click(sender As Object, e As EventArgs) Handles _cmdFind_1.Click
		Call cmdFind_Click(1)
	End Sub

	Private Sub cmdSelect_Click(sender As Object, e As EventArgs) Handles cmdSelect.Click
		Call lstCompo_SelectedIndexChanged(lstCompo, New System.EventArgs())
	End Sub

	Private Sub cmdCalculate_Click(sender As Object, e As EventArgs) Handles cmdCalculate.Click
		cmdCalculate_Click()
	End Sub

	Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
		RadioButton1.Checked = True
		RadioButton2.Checked = False
		RadioButton3.Checked = False
		Call optFreundlichSource_Click(0, 0)
	End Sub

	Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
		RadioButton1.Checked = False
		RadioButton2.Checked = True
		RadioButton3.Checked = False
		Call optFreundlichSource_Click(1, 0)
	End Sub

	Private Sub RadioButton3_Click(sender As Object, e As EventArgs) Handles RadioButton3.Click
		RadioButton1.Checked = False
		RadioButton2.Checked = False
		RadioButton3.Checked = True
		Call optFreundlichSource_Click(2, 0)
	End Sub

	Private Sub frmFreundlich_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class