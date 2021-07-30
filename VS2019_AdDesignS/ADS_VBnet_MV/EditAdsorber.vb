Option Strict Off
Option Explicit On
Friend Class frmEditAdsorber
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Dim frmEditAdsorber_Cancelled As Short
	Dim frmEditAdsorber_RunMode As Short
	Const frmEditAdsorber_RunMode_QUERY_DATABASE As Short = 1
	Const frmEditAdsorber_RunMode_EDIT_DATABASE As Short = 2
	
	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_USE_THESE As Boolean
	
	
	
	
	Const frmEditAdsorber_declarations_end As Boolean = True
	
	
	Sub frmEditAdsorber_QueryDatabase(ByRef OUTPUT_User_Transferred_Data As Boolean)
		frmEditAdsorber_RunMode = frmEditAdsorber_RunMode_QUERY_DATABASE
		Me.ShowDialog()

		If (USER_HIT_USE_THESE) Then
			OUTPUT_User_Transferred_Data = True
		Else
			OUTPUT_User_Transferred_Data = False
		End If
	End Sub
	Sub frmEditAdsorber_EditDatabase()
		frmEditAdsorber_RunMode = frmEditAdsorber_RunMode_EDIT_DATABASE
		Me.ShowDialog()

	End Sub
	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdOK.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdOK.Enabled = False
			mnuManufacturerItem(1).Enabled = False
			mnuManufacturerItem(2).Enabled = False
			mnuManufacturerItem(3).Enabled = False
			mnuAdsorberItem(1).Enabled = False
			mnuAdsorberItem(2).Enabled = False
			mnuAdsorberItem(3).Enabled = False
		End If
	End Sub
	
	
	Private Function adsorber_db_AssignUniqueID() As Short
		Dim this_try As Short
		Dim i As Short
		Dim Found As Short
		Do While (1 = 1)
			this_try = Int(Rnd(1) * 32000) + 1
			Found = False
			For i = 1 To adsorber_db_num_manufacturers
				If (CShort(adsorber_db_manufacturers(i).UniqueID) = this_try) Then
					Found = True
					Exit For
				End If
			Next i
			If (Not Found) Then
				adsorber_db_AssignUniqueID = this_try
				Exit Function
			End If
		Loop 
	End Function
	
	
	Private Sub adsorber_db_displayall()
		Dim i As Short
		Dim N As Short
		'POPULATE LISTBOX lstManu:
		lstManu.Items.Clear()
		For i = 1 To adsorber_db_num_manufacturers
			N = lstManu.Items.Add(adsorber_db_manufacturers(i).Name)
			'UPGRADE_ISSUE: ListBox property lstManu.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			'N = lstManu.NewIndex
			VB6.SetItemData(lstManu, N, CShort(adsorber_db_manufacturers(i).UniqueID))
		Next i
		lstManu.SetSelected(0, True)

	End Sub
	
	
	'Load all adsorber DB entries.
	Private Sub adsorber_db_loadall()
		Dim f As Short
		Dim i As Short
		Dim fn As String
		
		On Error GoTo err_adsorber_db_loadall
		
		'LOAD MANUFACTURERS.
		fn = Database_Path & "\beds2.txt"
		f = FreeFile
		FileOpen(f, fn, OpenMode.Input)
		Input(f, adsorber_db_num_manufacturers)
		If (adsorber_db_num_manufacturers <> 0) Then
			'UPGRADE_WARNING: Lower bound of array adsorber_db_manufacturers was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim adsorber_db_manufacturers(adsorber_db_num_manufacturers)
			For i = 1 To adsorber_db_num_manufacturers
				Input(f, adsorber_db_manufacturers(i).UniqueID)
				Input(f, adsorber_db_manufacturers(i).Name)
			Next i
		End If
		FileClose(f)
		
		'LOAD ADSORBERS.
		fn = Database_Path & "\beds1.txt"
		f = FreeFile
		FileOpen(f, fn, OpenMode.Input)
		Input(f, adsorber_db_num_adsorbers)
		If (adsorber_db_num_adsorbers <> 0) Then
			'UPGRADE_WARNING: Lower bound of array adsorber_db_adsorbers was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim adsorber_db_adsorbers(adsorber_db_num_adsorbers)
			For i = 1 To adsorber_db_num_adsorbers
				Input(f, adsorber_db_adsorbers(i).UniqueID_Manufacturer)
				Input(f, adsorber_db_adsorbers(i).Phase)
				Input(f, adsorber_db_adsorbers(i).PartNumber)
				Input(f, adsorber_db_adsorbers(i).InternalArea)
				Input(f, adsorber_db_adsorbers(i).MaxCapacity)
				Input(f, adsorber_db_adsorbers(i).OutsideDiameter)
				Input(f, adsorber_db_adsorbers(i).DesignPressure)
				Input(f, adsorber_db_adsorbers(i).DesignFlowRange)
				Input(f, adsorber_db_adsorbers(i).DefaultFlowRate)
				Input(f, adsorber_db_adsorbers(i).Note)
				'MsgBox CStr(adsorber_db_adsorbers(i).PartNumber)
			Next i
		End If
		FileClose(f)
exit_err_adsorber_db_loadall: 
		Exit Sub
err_adsorber_db_loadall: 
		Call Show_Trapped_Error("Load Adsorber Database")
		Resume exit_err_adsorber_db_loadall
	End Sub
	
	
	'RETURNS:
	'  -1 if not found
	'  index if found
	Private Function adsorber_db_lookup_UniqueID_Manufacturer(ByRef search_for As Short) As Short
		Dim i As Short
		Dim Found As Short
		Found = False
		For i = 1 To adsorber_db_num_manufacturers
			If (CShort(adsorber_db_manufacturers(i).UniqueID) = search_for) Then
				Found = True
				Exit For
			End If
		Next i
		If (Found) Then
			adsorber_db_lookup_UniqueID_Manufacturer = i
		Else
			adsorber_db_lookup_UniqueID_Manufacturer = -1
		End If
	End Function
	
	
	Private Sub adsorber_db_saveall()
		Dim f As Short
		Dim i As Short
		Dim fn As String
		
		On Error GoTo err_adsorber_db_saveall
		
		'SAVE MANUFACTURERS.
		fn = Database_Path & "\beds2.txt"
		f = FreeFile
		FileOpen(f, fn, OpenMode.Output)
		WriteLine(f, adsorber_db_num_manufacturers)
		'ReDim adsorber_db_manufacturers(1 To adsorber_db_num_manufacturers)
		For i = 1 To adsorber_db_num_manufacturers
			WriteLine(f, adsorber_db_manufacturers(i).UniqueID, adsorber_db_manufacturers(i).Name)
		Next i
		FileClose(f)
		
		'SAVE ADSORBERS.
		fn = Database_Path & "\beds1.txt"
		f = FreeFile
		FileOpen(f, fn, OpenMode.Output)
		WriteLine(f, adsorber_db_num_adsorbers)
		'ReDim adsorber_db_adsorbers(1 To adsorber_db_num_adsorbers)
		For i = 1 To adsorber_db_num_adsorbers
			WriteLine(f, adsorber_db_adsorbers(i).UniqueID_Manufacturer, adsorber_db_adsorbers(i).Phase, Trim(adsorber_db_adsorbers(i).PartNumber), Trim(adsorber_db_adsorbers(i).InternalArea), Trim(adsorber_db_adsorbers(i).MaxCapacity), Trim(adsorber_db_adsorbers(i).OutsideDiameter), Trim(adsorber_db_adsorbers(i).DesignPressure), Trim(adsorber_db_adsorbers(i).DesignFlowRange), Trim(adsorber_db_adsorbers(i).DefaultFlowRate), Trim(adsorber_db_adsorbers(i).Note))
			'MsgBox CStr(adsorber_db_adsorbers(i).PartNumber)
		Next i
		FileClose(f)
		
		Call clear_this_record()
		lstName.Items.Clear()
		Call adsorber_db_loadall()
		Call adsorber_db_displayall()
exit_err_adsorber_db_saveall: 
		Exit Sub
err_adsorber_db_saveall: 
		Call Show_Trapped_Error("Save Adsorber Database")
		Resume exit_err_adsorber_db_saveall
	End Sub
	
	
	Private Sub clear_this_record()
		Dim i As Short
		For i = 0 To 6
			lblData(i).Text = ""
		Next i
	End Sub
	
	
	Private Sub cmdCancel_Click()
		'frmEditAdsorber_Cancelled = True
		USER_HIT_CANCEL = True
		USER_HIT_USE_THESE = False
		Me.Dispose()
	End Sub
	Private Sub cmdOK_Click()
		Dim N As Short
		Dim this_A As Double
		Dim this_M As Double
		Dim this_rhoB As Double
		Dim New_D As Double
		Dim new_V As Double
		Dim New_L As Double
		Dim this_Q As Double
		Dim now_phase As Short
		If (lstName.SelectedIndex < 0) Or (lstName.Items.Count = 0) Then
			Call Show_Error("You must first select an adsorber.")
			Exit Sub
		End If
		N = VB6.GetItemData(lstName, lstName.SelectedIndex)
		'CONVERT AREA FROM FT^2 TO M^2
		this_A = Convert.ToDouble(adsorber_db_adsorbers(N).InternalArea)
		this_A = this_A / 10.7639104167
		'CONVERT MASS FROM LBS TO KG
		this_M = Convert.ToDouble(adsorber_db_adsorbers(N).MaxCapacity)
		this_M = this_M / 2.20462262185
		'CONVERT BULK DENSITY FROM LBM/FT^3 TO KG/M^3
		this_rhoB = 28#
		this_rhoB = this_rhoB * 0.45359237 / 0.028316846592
		'CALCULATE INTERNAL DIAMETER IN M
		New_D = (4# * this_A / 3.14159) ^ 0.5
		'CALCULATE VOLUME IN M^3
		new_V = this_M / this_rhoB
		'CALCULATE LENGTH IN M
		New_L = new_V / this_A
		'CONVERT DEFAULT FLOW RATE TO M^3/S
		this_Q = Convert.ToDouble(adsorber_db_adsorbers(N).DefaultFlowRate)
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optphase_1.Checked) Then
			now_phase = 1 'LIQUID PHASE
		Else
			now_phase = 2 'GAS PHASE
		End If
		If (now_phase = 1) Then
			'CONVERT FROM GAL/MIN TO M^3/S
			this_Q = this_Q * 0.003785411784 / 60#
		End If
		If (now_phase = 2) Then
			'CONVERT FROM FT^3/MIN TO M^3/S
			this_Q = this_Q * 0.028316846592 / 60#
		End If
		'TRANSFER PARAMETERS BACK TO MAIN SCREEN
		frmEditAdsorber_ReturnParameters.D = New_D
		frmEditAdsorber_ReturnParameters.L = New_L
		frmEditAdsorber_ReturnParameters.Q = this_Q
		frmEditAdsorber_ReturnParameters.M = this_M
		'frmEditAdsorber_Cancelled = False
		USER_HIT_CANCEL = False
		USER_HIT_USE_THESE = True
		Me.Dispose()
	End Sub
	
	
	Private Sub frmEditAdsorber_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)
		'MISC INITS.
		'		Me.Height = VB6.TwipsToPixelsY(7290)
		'		Me.Width = VB6.TwipsToPixelsX(9255)
		Call CenterOnForm(Me, frmMain)
		If (frmEditAdsorber_RunMode = frmEditAdsorber_RunMode_QUERY_DATABASE) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdOK.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdOK.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancel.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdCancel.Visible = True
			lblData(7).Text = "28.0"
			lblData(7).Visible = True
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdOK.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdOK.Visible = False
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancel.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdCancel.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancel.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdCancel.Text = "E&xit"
			lblData(7).Visible = False
			lblDesc(7).Visible = False
			lblUnits(7).Visible = False
		End If
		If (Bed.Phase = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

			_optphase_1.Checked = True
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_2.Checked = False
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_1.Checked = False
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_2.Checked = True
		End If
		Call adsorber_db_loadall()
		Call clear_this_record()
		Call adsorber_db_displayall()

		' DEMO SETTINGS.
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	
	
	'UPGRADE_WARNING: Event lstManu.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstManu_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstManu.SelectedIndexChanged
		Dim thismanu As Short
		Dim i As Short
		Dim N As Short
		Dim now_phase As Short
		Dim partNumberStr As String
		Call clear_this_record()
		'DISPLAY NAMES FOR THIS MANUFACTURER (IF ANY)
		lstName.Items.Clear()
		If (lstManu.SelectedIndex < 0) Then Exit Sub
		thismanu = VB6.GetItemData(lstManu, lstManu.SelectedIndex)
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optphase_1.Checked) Then
			now_phase = 1 'LIQUID PHASE
		Else
			now_phase = 2 'GAS PHASE
		End If
		For i = 1 To adsorber_db_num_adsorbers
			If (thismanu = adsorber_db_adsorbers(i).UniqueID_Manufacturer) Then
				If (now_phase = adsorber_db_adsorbers(i).Phase) Then
					partNumberStr = New String(adsorber_db_adsorbers(i).PartNumber)
					N = lstName.Items.Add(partNumberStr)
					'UPGRADE_ISSUE: ListBox property lstName.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					'N = lstName.NewIndex
					VB6.SetItemData(lstName, N, i)
					'lstName.Items(N).itemdata = i
				End If
			End If
		Next i
	End Sub
	Private Sub lstManu_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstManu.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If ((Button And 2) = 2) Then
			'UPGRADE_ISSUE: Form method frmEditAdsorber.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'Me.PopupMenu(mnuManufacturer)
		End If
	End Sub
	'UPGRADE_WARNING: Event lstName.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstName_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstName.SelectedIndexChanged
		Dim N As Short
		'DISPLAY ADSORBER PROPERTIES:
		N = VB6.GetItemData(lstName, lstName.SelectedIndex)
		lblData(0).Text = Trim(adsorber_db_adsorbers(N).InternalArea)
		lblData(1).Text = Trim(adsorber_db_adsorbers(N).MaxCapacity)
		lblData(2).Text = Trim(adsorber_db_adsorbers(N).OutsideDiameter)
		lblData(3).Text = Trim(adsorber_db_adsorbers(N).DesignPressure)
		lblData(4).Text = Trim(adsorber_db_adsorbers(N).DesignFlowRange)
		lblData(5).Text = Trim(adsorber_db_adsorbers(N).DefaultFlowRate)
		lblData(6).Text = Trim(adsorber_db_adsorbers(N).Note)
	End Sub
	Private Sub lstName_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstName.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If ((Button And 2) = 2) Then
			'UPGRADE_ISSUE: Form method frmEditAdsorber.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'Me.PopupMenu(mnuAdsorber)
		End If
	End Sub
	
	
	Public Sub mnuAdsorberItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAdsorberItem.Click
		Dim Index As Short = mnuAdsorberItem.GetIndex(eventSender)
		Dim N As Short
		Dim i As Short
		Dim response As Short
		Dim msg As String
		Dim n_manu As Short
		Dim now_phase As Short
		Dim USER_HIT_CANCEL As Boolean
		If (lstManu.SelectedIndex < 0) Or (lstManu.Items.Count = 0) Then
			Call Show_Error("You must first select a manufacturer.")
			Exit Sub
		End If
		n_manu = VB6.GetItemData(lstManu, lstManu.SelectedIndex)
		n_manu = adsorber_db_lookup_UniqueID_Manufacturer(n_manu)
		If (Index = 2) Or (Index = 3) Then
			If (lstName.SelectedIndex < 0) Or (lstName.Items.Count = 0) Then
				Call Show_Error("You must first select an adsorber.")
				Exit Sub
			End If
			N = VB6.GetItemData(lstName, lstName.SelectedIndex)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optphase_1.Checked) Then
			now_phase = 1 'LIQUID PHASE
		Else
			now_phase = 2 'GAS PHASE
		End If
		Select Case Index
			Case 1 'new
				Call frmEditAdsorberData.frmEditAdsorberData_AddNew(now_phase, USER_HIT_CANCEL)
				If (Not USER_HIT_CANCEL) Then
					frmEditAdsorberData_Record.UniqueID_Manufacturer = CShort(adsorber_db_manufacturers(n_manu).UniqueID)
					adsorber_db_num_adsorbers = adsorber_db_num_adsorbers + 1
					N = adsorber_db_num_adsorbers
					'UPGRADE_WARNING: Lower bound of array adsorber_db_adsorbers was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim Preserve adsorber_db_adsorbers(N)
					'UPGRADE_WARNING: Couldn't resolve default property of object adsorber_db_adsorbers(N). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					adsorber_db_adsorbers(N) = frmEditAdsorberData_Record
					Call adsorber_db_saveall()
				End If
			Case 2 'edit current
				'UPGRADE_WARNING: Couldn't resolve default property of object frmEditAdsorberData_Record. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmEditAdsorberData_Record = adsorber_db_adsorbers(N)
				Call frmEditAdsorberData.frmEditAdsorberData_Edit(adsorber_db_adsorbers(N).Phase, USER_HIT_CANCEL)
				If (Not USER_HIT_CANCEL) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object adsorber_db_adsorbers(N). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					adsorber_db_adsorbers(N) = frmEditAdsorberData_Record
					Call adsorber_db_saveall()
				End If
			Case 3 'delete current
				msg = "Do you really want to delete adsorber '" & Trim(adsorber_db_adsorbers(N).PartNumber) & "' ?"
				response = MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.YesNo, AppName_For_Display_Short)
				If response = MsgBoxResult.No Then Exit Sub
				'PERFORM DELETION
				For i = N To adsorber_db_num_adsorbers - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object adsorber_db_adsorbers(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					adsorber_db_adsorbers(i) = adsorber_db_adsorbers(i + 1)
				Next i
				adsorber_db_num_adsorbers = adsorber_db_num_adsorbers - 1
				'SAVE MANUFACTURER FILE
				Call adsorber_db_saveall()
		End Select
	End Sub
	
	
	Public Sub mnuManufacturerItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuManufacturerItem.Click
		Dim Index As Short = mnuManufacturerItem.GetIndex(eventSender)
		Dim N As Short
		Dim i As Short
		Dim response As Short
		Dim msg As String
		Dim new_UniqueID As Short
		Dim NewName As String
		Dim USER_HIT_CANCEL As Boolean
		If (Index = 2) Or (Index = 3) Then
			If (lstManu.SelectedIndex < 0) Or (lstManu.Items.Count = 0) Then
				Call Show_Error("You must first select a manufacturer.")
				Exit Sub
			End If
			N = VB6.GetItemData(lstManu, lstManu.SelectedIndex)
			N = adsorber_db_lookup_UniqueID_Manufacturer(N)
		End If
		Select Case Index
			Case 1 'new
				NewName = "New Manufacturer"
				Do While (1 = 1)
					NewName = frmNewName.frmNewName_GetName("Creating New Manufacturer", "Each manufacturer record should have a unique name.", NewName, USER_HIT_CANCEL)
					If (USER_HIT_CANCEL) Then Exit Sub
					NewName = Trim(NewName)
					If (NewName <> "") Then Exit Do
					Call Show_Error("Manufacturer name must be a non-blank string.")
				Loop 
				'ADD THE NEW MANUFACTURER RECORD.
				new_UniqueID = adsorber_db_AssignUniqueID()
				adsorber_db_num_manufacturers = adsorber_db_num_manufacturers + 1
				N = adsorber_db_num_manufacturers
				'UPGRADE_WARNING: Lower bound of array adsorber_db_manufacturers was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve adsorber_db_manufacturers(N)
				adsorber_db_manufacturers(N).Name = Trim(CStr(NewName))
				adsorber_db_manufacturers(N).UniqueID = Trim(Str(new_UniqueID))
				Call adsorber_db_saveall()
			Case 2 'edit current
				'If Number_Of_Manufacturers = 0 Then
				'  MsgBox "There is no manufacturer name to edit.", MB_ICONEXCLAMATION, AppName_For_Display_long
				'  Exit Sub
				'End If
				NewName = Trim(VB6.GetItemString(lstManu, lstManu.SelectedIndex))
				Do While (1 = 1)
					NewName = frmNewName.frmNewName_GetName("Editing Existing Manufacturer Name", "Each manufacturer record should have a unique name.", NewName, USER_HIT_CANCEL)
					If (USER_HIT_CANCEL) Then Exit Sub
					NewName = Trim(NewName)
					If (NewName <> "") Then Exit Do
					Call Show_Error("Manufacturer name must be a non-blank string.")
				Loop 
				adsorber_db_manufacturers(N).Name = NewName
				Call adsorber_db_saveall()
			Case 3 'delete current
				msg = "Do you really want to delete manufacturer '" & Trim(adsorber_db_manufacturers(N).Name) & "' ?"
				response = MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.YesNo, AppName_For_Display_Long)
				If response = MsgBoxResult.No Then Exit Sub
				'PERFORM DELETION
				For i = N To adsorber_db_num_manufacturers - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object adsorber_db_manufacturers(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					adsorber_db_manufacturers(i) = adsorber_db_manufacturers(i + 1)
				Next i
				adsorber_db_num_manufacturers = adsorber_db_num_manufacturers - 1
				'SAVE MANUFACTURER FILE
				Call adsorber_db_saveall()
		End Select
	End Sub
	
	
	Private Sub optPhase_Click(ByRef Index As Short, ByRef Value As Short)
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optphase_1.Checked = True And Index = 1) Then
			'LIQUID PHASE
			_optphase_2.Checked = False
			lblUnits(4).Text = "(gal/min)"
			lblUnits(5).Text = "(gal/min)"
		ElseIf (_optphase_2.Checked = True And Index = 2) Then
			'GAS PHASE
			_optphase_1.Checked = False
			lblUnits(4).Text = "(ft³/min)"
			lblUnits(5).Text = "(ft³/min)"
		ElseIf (_optphase_1.Checked = False And Index = 1) Then
			'Gas PHASE
			_optphase_2.Checked = True
			lblUnits(4).Text = "(ft³/min)"
			lblUnits(5).Text = "(ft³/min)"
		ElseIf (_optphase_2.Checked = False And Index = 2) Then
			'Liquid Phase
			_optphase_1.Checked = True
			lblUnits(4).Text = "(gal/min)"
			lblUnits(5).Text = "(gal/min)"
		End If
		If (lstManu.Items.Count > 0) Then
			'UPDATE LIST OF NAMES:
			If (lstManu.SelectedIndex < 0) Then
				lstManu.SelectedIndex = 0
			End If
			Call lstManu_SelectedIndexChanged(lstManu, New System.EventArgs())
		End If
	End Sub


	Private Sub _optphase_1_CheckedChanged(sender As Object, e As EventArgs) Handles _optphase_1.CheckedChanged
		Call optPhase_Click(1, 0)
	End Sub

	Private Sub _optphase_2_CheckedChanged(sender As Object, e As EventArgs) Handles _optphase_2.CheckedChanged
		Call optPhase_Click(2, 0)
	End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
		Call cmdCancel_Click()
	End Sub

	Private Sub OK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
		Call cmdOK_Click()
	End Sub

	Private Sub frmEditAdsorber_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)
	End Sub
End Class