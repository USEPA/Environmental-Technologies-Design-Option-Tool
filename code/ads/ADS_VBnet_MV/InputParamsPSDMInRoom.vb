Option Strict Off
Option Explicit On
Friend Class frmInputParamsPSDMInRoom
	Inherits System.Windows.Forms.Form
	
	'UPGRADE_WARNING: Arrays in structure TempData may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim TempData As RoomParam_Type
	Dim NOW_CONTAMINANT As Short
	
	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_OK As Boolean
	'UPGRADE_WARNING: Arrays in structure Temp_RP may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim Temp_RP As RoomParam_Type
	
	Dim frmInputParamsPSDMInRoom_Is_Dirty As Boolean
	Dim HALT_cboChemical As Boolean
	Public HALT_cbo_RXN_PRODUCT As Boolean
	Public HALT_ALL_CONTROLS As Boolean
	
	Const IN_cmdTimeVar_WhichButton___CO As Short = 1
	Const IN_cmdTimeVar_WhichButton___WA As Short = 2
	Const IN_cmdTimeVar_WhichButton___K As Short = 3
	
	
	
	
	Const frmInputParamsPSDMInRoom_declarations_end As Boolean = True
	
	
	Sub frmInputParamsPSDMInRoom_Edit(ByRef OUTPUT_Raise_Dirty_Flag As Boolean)
		'UPGRADE_WARNING: Couldn't resolve default property of object Temp_RP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Temp_RP = RoomParams
		If (Temp_RP.COUNT_CONTAMINANT <> Number_Component) Then
			Temp_RP.COUNT_CONTAMINANT = Number_Component
		End If
		Me.ShowDialog()
		If (USER_HIT_OK) Then
			OUTPUT_Raise_Dirty_Flag = True
			'UPGRADE_WARNING: Couldn't resolve default property of object RoomParams. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RoomParams = Temp_RP
		Else
			OUTPUT_Raise_Dirty_Flag = False
		End If
	End Sub
	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CancelButton.Enabled = False
		End If
	End Sub
	
	
	Sub frmInputParamsPSDMInRoom_PopulateUnits()
		Dim Frm As System.Windows.Forms.Form
		Frm = Me
		'MAIN BLOCK OF UNITS.
		Call unitsys_register(Frm, lblData(0), txtData(0), cboUnits(0), "volume", Temp_RP.ROOM_VOL_Units, "m", "", "", 100#, True)
		Call unitsys_register(Frm, lblData(1), txtData(1), cboUnits(1), "flow_volumetric", Temp_RP.ROOM_FLOWRATE_Units, "m³/s", "", "", 100#, True)
		Call unitsys_register(Frm, lblData(2), txtData(2), cboUnits(2), "concentration", Temp_RP.ROOM_C0_Units, "mg/L", "", "", 100#, True)
		Call unitsys_register(Frm, lblData(3), txtData(3), cboUnits(3), "mass_emission_rate", Temp_RP.ROOM_EMIT_Units, "µg/s", "", "", 100#, True)
		Call unitsys_register(Frm, lblData(4), txtData(4), cboUnits(4), "concentration", Temp_RP.INITIAL_ROOM_CONC_Units, "mg/L", "", "", 100#, True)
		Call unitsys_register(Frm, lblData(5), txtData(5), cboUnits(5), "inverse_time", "1/s", "1/s", "", "", 100#, True)
		Call unitsys_register(Frm, lblData(6), txtData(6), Nothing, "", "", "", "", "", 100#, False)
		'Call unitsys_register(Frm, lblData(7), txtData(7), cboUnits(7), "freundlich_k", "(mg/g)*(L/mg)^(1/n)", "(mg/g)*(L/mg)^(1/n)", "", "", 100#, True)
		Call unitsys_register(Frm, Nothing, txtData(7), cboUnits(7), "freundlich_k", "(mg/g)*(L/mg)^(1/n)", "(mg/g)*(L/mg)^(1/n)", "", "", 100.0#, True)

	End Sub
	Sub Store_Unit_Settings()
		Temp_RP.ROOM_VOL_Units = unitsys_get_units(cboUnits(0))
		Temp_RP.ROOM_FLOWRATE_Units = unitsys_get_units(cboUnits(1))
		Temp_RP.ROOM_C0_Units = unitsys_get_units(cboUnits(2))
		Temp_RP.ROOM_EMIT_Units = unitsys_get_units(cboUnits(3))
		Temp_RP.INITIAL_ROOM_CONC_Units = unitsys_get_units(cboUnits(4))
		'Temp_RP.u_ROOM_KINI = unitsys_get_units(cboUnits(7))
	End Sub
	
	
	Sub Do_Refresh()
		Call frmInputParamsPSDMInRoom_Refresh(Temp_RP, NOW_CONTAMINANT)
	End Sub
	
	
	Private Sub populate_cboChemical()
		Dim i As Short
		Dim Ctl As ComboBox
		Ctl = cboChemical
		HALT_cboChemical = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ctl.Items.Clear()
		If (frmMain.cboSelectCompo.Items.Count > 0) Then
			For i = 1 To frmMain.cboSelectCompo.Items.Count
				'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Ctl.Items.Add(Trim(VB6.GetItemString(frmMain.cboSelectCompo, i - 1)))
			Next
			'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Ctl.SelectedIndex = 0
		End If
		HALT_cboChemical = False
	End Sub
	Private Sub populate_cbo_RXN_PRODUCT()
		Dim i As Short
		Dim Ctl As ComboBox
		Ctl = cbo_RXN_PRODUCT
		HALT_cbo_RXN_PRODUCT = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ctl.Items.Clear()
		If (frmMain.cboSelectCompo.Items.Count > 0) Then
			For i = 1 To frmMain.cboSelectCompo.Items.Count
				Ctl.Items.Add(New VB6.ListBoxItem(Trim(VB6.GetItemString(frmMain.cboSelectCompo, i - 1)), i))
			Next
			'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Ctl.SelectedIndex = 0
		End If
		HALT_cbo_RXN_PRODUCT = False
	End Sub
	
	
	Sub frmInputParamsPSDMInRoom_GenericStatus_Set(ByRef fn_Text As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Me.sspanel_Status.Caption = fn_Text
	End Sub
	Sub frmInputParamsPSDMInRoom_DirtyStatus_Set(ByRef newVal As Boolean)
		Dim Frm As frmInputParamsPSDMInRoom
		Frm = Me
		If (newVal) Then
			'UPGRADE_ISSUE: Control sspanel_Dirty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm.sspanel_Dirty.Caption = "Data Changed"
			'UPGRADE_ISSUE: Control sspanel_Dirty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm.sspanel_Dirty.ForeColor = Color.FromArgb(QBColor(12))
		Else
			'UPGRADE_ISSUE: Control sspanel_Dirty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm.sspanel_Dirty.Caption = "Unchanged"
			'UPGRADE_ISSUE: Control sspanel_Dirty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm.sspanel_Dirty.ForeColor = Color.FromArgb(QBColor(0))
		End If
	End Sub
	Sub frmInputParamsPSDMInRoom_DirtyStatus_Set_Current()
		Call frmInputParamsPSDMInRoom_DirtyStatus_Set(frmInputParamsPSDMInRoom_Is_Dirty)
	End Sub
	Sub frmInputParamsPSDMInRoom_DirtyStatus_Throw()
		frmInputParamsPSDMInRoom_Is_Dirty = True
		Call frmInputParamsPSDMInRoom_DirtyStatus_Set_Current()
	End Sub
	Sub frmInputParamsPSDMInRoom_DirtyStatus_Clear()
		frmInputParamsPSDMInRoom_Is_Dirty = False
		Call frmInputParamsPSDMInRoom_DirtyStatus_Set_Current()
	End Sub
	
	
	'UPGRADE_WARNING: Event cbo_RXN_PRODUCT.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbo_RXN_PRODUCT_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbo_RXN_PRODUCT.SelectedIndexChanged
		Dim Ctl As ComboBox
		Ctl = cbo_RXN_PRODUCT
		If (HALT_cbo_RXN_PRODUCT) Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Temp_RP.RXN_PRODUCT(NOW_CONTAMINANT) = Ctl.Items(Ctl.SelectedIndex)
		'
		' THROW DIRTY FLAG AND REFRESH.
		Call frmInputParamsPSDMInRoom_DirtyStatus_Throw()
		''''Call RoomParam_Recalculate(Temp_RP)
		Call Do_Refresh()
	End Sub
	'UPGRADE_WARNING: Event cboChemical.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboChemical_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboChemical.SelectedIndexChanged
		If (HALT_cboChemical) Then Exit Sub
		NOW_CONTAMINANT = cboChemical.SelectedIndex + 1
		Call Do_Refresh()
	End Sub
	
	
	'UPGRADE_WARNING: Event cboUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboUnits.SelectedIndexChanged
		Dim Index As Short = cboUnits.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = cboUnits(Index)
		Call unitsys_control_cbox_click(Ctl)
	End Sub
	Private Sub cboUnits_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboUnits.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = cboUnits.GetIndex(eventSender)
		KeyAscii = Global_TextKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Private Sub cmdCancelOK_Click(ByRef Index As Short)
		Dim i As Short
		Select Case Index
			Case 0 'CANCEL.
				'If (frmCompoProp_Query_Unload() = False) Then
				'  'THE CANCEL WAS CANCELLED.
				'  Exit Sub
				'End If
				USER_HIT_CANCEL = True
				USER_HIT_OK = False
				Me.Close()
				Exit Sub
			Case 1 'OK.
				'STORE ALL UNIT SETTINGS.
				Call Store_Unit_Settings()
				'EXIT OUT OF HERE.
				USER_HIT_CANCEL = False
				USER_HIT_OK = True
				Me.Close()
				Exit Sub
		End Select
	End Sub
	
	Sub Do___cmdTimeVar___ButtonClick(ByRef IN_cmdTimeVar_WhichButton As Short)
		Dim FormCaption As String
		'UPGRADE_WARNING: Lower bound of array UnitType was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim UnitType(2) As String
		'UPGRADE_WARNING: Lower bound of array BaseUnits was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim BaseUnits(2) As String
		'UPGRADE_WARNING: Lower bound of array CurrentUnits was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim CurrentUnits(2) As String
		'UPGRADE_WARNING: Lower bound of array lblUnitType was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim lblUnitType(2) As String
		Dim DataRowCount As Short
		Dim MaxRows As Short
		Dim ColumnCount As Short
		Dim ColumnNames() As String
		Dim foStoreTo As VCIF1Lib.F1Book
		Dim USER_HIT_CANCEL As Boolean
		'
		' EXTRA STEP:
		' TRANSFER dbl_ROOM_COINI(), dbl_ROOM_TCOINI()
		' DATA INTO frmMain.foVarConc.
		'
		Dim i As Short
		Dim J As Short
		Dim Ctl As VCIF1Lib.F1Book
		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.foVarConc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	Ctl = frmMain.foVarConc   'Where is forVarConc
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ctl.Sheet = 1
		Select Case IN_cmdTimeVar_WhichButton
			Case IN_cmdTimeVar_WhichButton___CO
				If (Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT) = 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.MaxRow = 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.MaxRow = Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT)
				End If
				For i = 1 To Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.EntryRC(i, 1) = Temp_RP.dbl_ROOM_TCOINI(NOW_CONTAMINANT, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.EntryRC(i, 2) = Temp_RP.dbl_ROOM_COINI(NOW_CONTAMINANT, i)
				Next i
			Case IN_cmdTimeVar_WhichButton___WA
				If (Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT) = 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.MaxRow = 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.MaxRow = Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT)
				End If
				For i = 1 To Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.EntryRC(i, 1) = Temp_RP.dbl_ROOM_TEMITINI(NOW_CONTAMINANT, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.EntryRC(i, 2) = Temp_RP.dbl_ROOM_EMITINI(NOW_CONTAMINANT, i)
				Next i
			Case IN_cmdTimeVar_WhichButton___K
				If (Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT) = 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.MaxRow = 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.MaxRow = Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT)
				End If
				For i = 1 To Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.EntryRC(i, 1) = Temp_RP.dbl_ROOM_TKINI(NOW_CONTAMINANT, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Ctl.EntryRC(i, 2) = Temp_RP.dbl_ROOM_KINI(NOW_CONTAMINANT, i)
				Next i
		End Select
		'
		' NOW, PROCEED WITH NORMAL CODE.
		'
		ColumnCount = 2
		'UPGRADE_WARNING: Lower bound of array ColumnNames was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim ColumnNames(2)
		ColumnNames(1) = "Time"
		Select Case IN_cmdTimeVar_WhichButton
			Case IN_cmdTimeVar_WhichButton___CO
				FormCaption = VB6.GetItemString(Me.cboChemical, Me.cboChemical.SelectedIndex) & " Influent Concentrations To Room (Time-Variable)"
				UnitType(1) = "time"
				UnitType(2) = "concentration"
				BaseUnits(1) = "min"
				BaseUnits(2) = "µg/L"
				CurrentUnits(1) = Temp_RP.u_ROOM_TCOINI
				CurrentUnits(2) = Temp_RP.u_ROOM_COINI
				lblUnitType(1) = "Time Units:"
				lblUnitType(2) = "Concentration Units:"
				DataRowCount = Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT)
				MaxRows = Max_int_ROOM_NCOINI
				ColumnNames(2) = "Concentration"
			Case IN_cmdTimeVar_WhichButton___WA
				FormCaption = VB6.GetItemString(Me.cboChemical, Me.cboChemical.SelectedIndex) & " Mass Emission Rates (Time-Variable)"
				UnitType(1) = "time"
				UnitType(2) = "mass_emission_rate"
				BaseUnits(1) = "min"
				BaseUnits(2) = "µg/s"
				CurrentUnits(1) = Temp_RP.u_ROOM_TEMITINI
				CurrentUnits(2) = Temp_RP.u_ROOM_EMITINI
				lblUnitType(1) = "Time Units:"
				lblUnitType(2) = "Emission Rate Units:"
				DataRowCount = Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT)
				MaxRows = Max_int_ROOM_NEMITINI
				ColumnNames(2) = "Mass Emission Rate"
			Case IN_cmdTimeVar_WhichButton___K
				FormCaption = VB6.GetItemString(Me.cboChemical, Me.cboChemical.SelectedIndex) & " Freundlich K (Time-Variable)"
				UnitType(1) = "time"
				UnitType(2) = "freundlich_k"
				BaseUnits(1) = "min"
				BaseUnits(2) = "(mg/g)*(L/mg)^(1/n)"
				CurrentUnits(1) = Temp_RP.u_ROOM_TKINI
				CurrentUnits(2) = Temp_RP.u_ROOM_KINI
				lblUnitType(1) = "Time Units:"
				lblUnitType(2) = "Freundlich K Units:"
				DataRowCount = Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT)
				MaxRows = Max_int_ROOM_NKINI
				ColumnNames(2) = "Freundlich K"
		End Select
		'
		' DISPLAY THE USER INPUT WINDOW.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.foVarConc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	foStoreTo = frmMain.foVarConc    Where is foVarConc  Shang
		Call frmTimeVarGrid.frmTimeVarGrid_Run(FormCaption, UnitType, BaseUnits, CurrentUnits, lblUnitType, DataRowCount, MaxRows, ColumnCount, ColumnNames, foStoreTo, USER_HIT_CANCEL)
		If (USER_HIT_CANCEL) Then
			Exit Sub
		End If
		Select Case IN_cmdTimeVar_WhichButton
			Case IN_cmdTimeVar_WhichButton___CO
				Temp_RP.u_ROOM_TCOINI = CurrentUnits(1)
				Temp_RP.u_ROOM_COINI = CurrentUnits(2)
				Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT) = DataRowCount
			Case IN_cmdTimeVar_WhichButton___WA
				Temp_RP.u_ROOM_TEMITINI = CurrentUnits(1)
				Temp_RP.u_ROOM_EMITINI = CurrentUnits(2)
				Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT) = DataRowCount
			Case IN_cmdTimeVar_WhichButton___K
				Temp_RP.u_ROOM_TKINI = CurrentUnits(1)
				Temp_RP.u_ROOM_KINI = CurrentUnits(2)
				Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT) = DataRowCount
		End Select
		'
		' EXTRA STEP:
		' TRANSFER frmMain.foVarConc DATA INTO
		' dbl_ROOM_COINI(), dbl_ROOM_TCOINI().
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.foVarConc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	Ctl = frmMain.foVarConc    'Where is foVarConce??? Shang
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ctl.Sheet = 1
		Select Case IN_cmdTimeVar_WhichButton
			Case IN_cmdTimeVar_WhichButton___CO
				For i = 1 To Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Temp_RP.dbl_ROOM_TCOINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 1)))
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Temp_RP.dbl_ROOM_COINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 2)))
				Next i
			Case IN_cmdTimeVar_WhichButton___WA
				For i = 1 To Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Temp_RP.dbl_ROOM_TEMITINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 1)))
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Temp_RP.dbl_ROOM_EMITINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 2)))
				Next i
			Case IN_cmdTimeVar_WhichButton___K
				For i = 1 To Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT)
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Temp_RP.dbl_ROOM_TKINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 1)))
					'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Temp_RP.dbl_ROOM_KINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 2)))
				Next i
		End Select
		'
		''DO _NOT_ ALLOW USER TO HIT CANCEL FROM INFLUENT FORM.
		''THEY MUST SAVE ALL INFLUENT DATA BECAUSE THEY MODIFIED
		''AN INFLUENT GRID.
		'cmdCancelOK(0).Enabled = False
		'
		'RAISE DIRTY FLAG AND REFRESH WINDOW.
		Call frmInputParamsPSDMInRoom_DirtyStatus_Throw()
		Call Do_Refresh()
	End Sub
	
	
	Private Sub cmdTimeVarConc_Click()
		Call Do___cmdTimeVar___ButtonClick(IN_cmdTimeVar_WhichButton___CO)
	End Sub
	Private Sub cmdTimeVarEmit_Click()
		Call Do___cmdTimeVar___ButtonClick(IN_cmdTimeVar_WhichButton___WA)
	End Sub
	Private Sub cmdTimeVarK_Click()
		Call Do___cmdTimeVar___ButtonClick(IN_cmdTimeVar_WhichButton___K)
	End Sub
	
	
	Private Sub frmInputParamsPSDMInRoom_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'
		' MISC INITS.
		'
		HALT_cboChemical = False
		HALT_ALL_CONTROLS = False
		Call CenterOnForm(Me, frmMain)
		lblSSValueUnits.Text = Chr(181) & "g/L"
		NOW_CONTAMINANT = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Component(0) = Component(NOW_CONTAMINANT)
		Call populate_cboChemical()
		Call populate_cbo_RXN_PRODUCT()
		'
		' POPULATE UNIT CONTROLS.
		'
		Call frmInputParamsPSDMInRoom_PopulateUnits()
		'
		' REFRESH WINDOW.
		'
		Call Do_Refresh()
		'
		' DATA UNCHANGED AS YET.
		'
		Call frmInputParamsPSDMInRoom_DirtyStatus_Clear()
		'
		' DEMO SETTINGS.
		'
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	Private Sub frmInputParamsPSDMInRoom_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub
	
	
	Private Sub optTimeVarConc_Click(ByRef Index As Short, ByRef Value As Short)
		Dim Ctl0 As AxThreed.AxSSOption
		Dim Ctl1 As AxThreed.AxSSOption
		'	Ctl0 = _optTimeVarConc_0
		'	Ctl1 = _optTimeVarConc_1
		Dim NewTag As Short
		Dim NewSetting As Short
		If (HALT_ALL_CONTROLS = True) Then Exit Sub
		NewTag = Index
		If (CShort(Val(Ctl0.Tag)) <> NewTag) Then
			NewSetting = IIf(NewTag = 0, False, True)
			Temp_RP.bool_ROOM_COINI_ISTIMEVAR(NOW_CONTAMINANT) = NewSetting
			'RAISE DIRTY FLAG AND REFRESH WINDOW.
			Call frmInputParamsPSDMInRoom_DirtyStatus_Throw()
			Call Do_Refresh()
		End If
	End Sub
	Private Sub optTimeVarEmit_Click(ByRef Index As Short, ByRef Value As Short)
		Dim Ctl0 As AxThreed.AxSSOption
		Dim Ctl1 As AxThreed.AxSSOption
		'	Ctl0 = _optTimeVarEmit_0
		'	Ctl1 = _optTimeVarEmit_1
		Dim NewTag As Short
		Dim NewSetting As Short
		If (HALT_ALL_CONTROLS = True) Then Exit Sub
		NewTag = Index
		If (CShort(Val(Ctl0.Tag)) <> NewTag) Then
			NewSetting = IIf(NewTag = 0, False, True)
			Temp_RP.bool_ROOM_EMITINI_ISTIMEVAR(NOW_CONTAMINANT) = NewSetting
			'RAISE DIRTY FLAG AND REFRESH WINDOW.
			Call frmInputParamsPSDMInRoom_DirtyStatus_Throw()
			Call Do_Refresh()
		End If
	End Sub
	Private Sub optTimeVarK_Click(ByRef Index As Short, ByRef Value As Short)
		Dim Ctl0 As AxThreed.AxSSOption
		Dim Ctl1 As AxThreed.AxSSOption
		'	Ctl0 = _optTimeVarK_0
		'	Ctl1 = _optTimeVarK_1
		Dim NewTag As Short
		Dim NewSetting As Short
		If (HALT_ALL_CONTROLS = True) Then Exit Sub
		NewTag = Index
		If (CShort(Val(Ctl0.Tag)) <> NewTag) Then
			NewSetting = IIf(NewTag = 0, False, True)
			Temp_RP.bool_ROOM_KINI_ISTIMEVAR(NOW_CONTAMINANT) = NewSetting
			'RAISE DIRTY FLAG AND REFRESH WINDOW.
			Call frmInputParamsPSDMInRoom_DirtyStatus_Throw()
			Call Do_Refresh()
		End If
	End Sub
	
	
	Private Sub txtData_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Enter
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim Ctl As TextBox
		Ctl = txtData(Index)
		Dim StatusMessagePanel As String
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Locked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Ctl.Enabled = False) Then Exit Sub
		'If (Index = 0) Then
		'  Call Global_GotFocus(Ctl)
		'Else
		Call unitsys_control_txtx_gotfocus(Ctl)
		'End If
		Select Case Index
			Case 0
				StatusMessagePanel = "Type in the volume of the room"
			Case 1
				StatusMessagePanel = "Type in the volumetric flow rate of air"
			Case 2
				StatusMessagePanel = "Type in the influent concentration to the room"
			Case 3
				StatusMessagePanel = "Type in the mass emission rate within the room"
			Case 4
				StatusMessagePanel = "Type in the concentration at time = zero"
		End Select
		Call frmInputParamsPSDMInRoom_GenericStatus_Set(StatusMessagePanel)
	End Sub
	Private Sub txtData_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtData.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtData.GetIndex(eventSender)
		'If (Index = 0) Then
		'  KeyAscii = Global_TextKeyPress(KeyAscii)
		'Else
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		'End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtData_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Leave
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As TextBox
		Ctl = txtData(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Locked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Ctl.Enabled = False) Then Exit Sub
		'  'HANDLE THE COMPONENT NAME TEXTBOX.
		'  If (Index = 0) Then
		'    If (Trim$(Ctl.Text) = "") Then
		'      Ctl.Text = Component(0).Name
		'      'Call Show_Error("You must enter a non-blank string for the component name.")
		'      'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
		'      'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
		'    Else
		'      If (Trim$(Component(0).Name) <> Trim$(Ctl.Text)) Then
		'        Component(0).Name = Trim$(Ctl.Text)
		'        'THROW DIRTY FLAG.
		'        Call frmCompoProp_DirtyStatus_Throw
		'      End If
		'    End If
		'    Call Global_LostFocus(Ctl)
		'    Call frmCompoProp_GenericStatus_Set("")
		'    Exit Sub
		'  End If
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
		Select Case Index
			Case 0 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 1 : Val_Low = 0# : Val_High = 1E+20
			Case 2 : Val_Low = 0# : Val_High = 1E+20
			Case 3 : Val_Low = 0# : Val_High = 1E+20
			Case 4 : Val_Low = 0# : Val_High = 1E+20
			Case 5 : Val_Low = 0# : Val_High = 1E+20
			Case 6 : Val_Low = 0# : Val_High = 1E+20
		End Select
		NewValue_Okay = False
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		Call frmInputParamsPSDMInRoom_GenericStatus_Set("")
		If (NewValue_Okay) Then
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				Select Case Index
					Case 0 : Temp_RP.ROOM_VOL = NewValue
					Case 1 : Temp_RP.ROOM_FLOWRATE = NewValue
					Case 2 : Temp_RP.ROOM_C0(NOW_CONTAMINANT) = NewValue
					Case 3 : Temp_RP.ROOM_EMIT(NOW_CONTAMINANT) = NewValue
					Case 4 : Temp_RP.INITIAL_ROOM_CONC(NOW_CONTAMINANT) = NewValue
					Case 5 : Temp_RP.RXN_RATE_CONSTANT(NOW_CONTAMINANT) = NewValue
					Case 6 : Temp_RP.RXN_RATIO(NOW_CONTAMINANT) = NewValue
				End Select
				'RAISE DIRTY FLAG AND RECALCULATE IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					'THROW DIRTY FLAG.
					Call frmInputParamsPSDMInRoom_DirtyStatus_Throw()
					Call RoomParam_Recalculate(Temp_RP)
				End If
				'REFRESH WINDOW.
				Call Do_Refresh()
			End If
		End If
	End Sub

	Private Sub OKButton_Click(sender As Object, e As EventArgs)
		'STORE ALL UNIT SETTINGS.
		Call Store_Unit_Settings()
		'EXIT OUT OF HERE.
		USER_HIT_CANCEL = False
		USER_HIT_OK = True
		Me.Close()
		Exit Sub
	End Sub

	Private Sub CancelButton_Click(sender As Object, e As EventArgs)
		USER_HIT_CANCEL = True
		USER_HIT_OK = False
		Me.Close()
		Exit Sub
	End Sub

End Class