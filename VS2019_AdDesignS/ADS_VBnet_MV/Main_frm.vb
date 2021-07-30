Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmMain
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Const frmMain_declarations_end As Boolean = True


	Sub frmMain_Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			mnuFileItem(0).Enabled = False
			''''mnuFileItem(1).Enabled = False
			mnuFileItem(2).Enabled = False
			mnuFileItem(3).Enabled = False
			''''mnuFileItem(191).Enabled = False
			''''mnuFileItem(192).Enabled = False
			''''mnuFileItem(193).Enabled = False
			''''mnuFileItem(194).Enabled = False
			mnuPhaseItem(0).Enabled = False
			mnuPhaseItem(1).Enabled = False
			cmdADEComponent(0).Enabled = False
			cmdADEComponent(1).Enabled = False
		End If
	End Sub


	Sub Avoid_Weird_Focus_Problem()
		Call unitsys_control_MostRecent_Force_lostfocus()
		Me.lstComponents.Focus()
	End Sub


	Sub Populate_frmMain_Units()
		'  'Fixed Bed Properties:
		'  Call Populate_Length_Units(txtBedUnits(0), LENGTH_M)
		'  Call Populate_Length_Units(txtBedUnits(1), LENGTH_M)
		'  Call Populate_Mass_Units(txtBedUnits(2), MASS_KG)
		'  Call Populate_Flowrate_Units(txtBedUnits(3), FLOW_M3_per_S)
		'  Call Populate_Time_Units(txtBedUnits(4), TIME_MIN)
		'  'time properties
		'  Call Populate_Time_Units(txttimeunits(0), TIME_D)
		'  Call Populate_Time_Units(txttimeunits(1), TIME_D)
		'  Call Populate_Time_Units(txttimeunits(2), TIME_D)
		'  'Adsorbent Properties:
		'  Call Populate_Density_Units(txtCarbonUnits(1), APPARENT_DENSITY_G_per_ML)
		'  Call Populate_Length_Units(txtCarbonUnits(2), LENGTH_M)

		'WATER/AIR PROPERTIES.
		Call unitsys_register(Me, lblWater(1), txtWater(1), Nothing, "", "", "", "", "", 100.0#, False)
		Call unitsys_register(Me, lblWater(0), txtWater(0), Nothing, "", "", "", "", "", 100.0#, False)
		'PSDM PARAMETERS.
		Call unitsys_register(Me, lblTime(0), txtTime(0), txtTimeUnits(0), "time", "d", "min", "", "", 100.0#, True)
		Call unitsys_register(Me, lblTime(1), txtTime(1), txtTimeUnits(1), "time", "d", "min", "", "", 100.0#, True)
		Call unitsys_register(Me, lblTime(2), txtTime(2), txtTimeUnits(2), "time", "d", "min", "", "", 100.0#, True)
		'unitsys w/ the spinners changed
		Call unitsys_register(Me, lblAxialElementsDesc, NumericUpDown1, Nothing, "", "", "", "0", "0", 100.0#, False)
		Call unitsys_register(Me, lblText(0), NumericUpDown2, Nothing, "", "", "", "0", "0", 100.0#, False)
		Call unitsys_register(Me, lblText(1), NumericUpDown3, Nothing, "", "", "", "0", "0", 100.0#, False)
		'BED PROPERTIES.
		Call unitsys_register(Me, lblBed(0), txtBedValue(0), txtBedUnits(0), "length", "m", "m", "", "", 100.0#, True)
		Call unitsys_register(Me, lblBed(1), txtBedValue(1), txtBedUnits(1), "length", "m", "m", "", "", 100.0#, True)
		Call unitsys_register(Me, lblBed(2), txtBedValue(2), txtBedUnits(2), "mass", "kg", "kg", "", "", 100.0#, True)
		Call unitsys_register(Me, lblBed(3), txtBedValue(3), txtBedUnits(3), "flow_volumetric", "m³/s", "m³/s", "", "", 100.0#, True)
		Call unitsys_register(Me, lblBed(4), txtBedValue(4), txtBedUnits(4), "time", "s", "s", "", "", 100.0#, True)
		'ADSORBENT PROPERTIES.
		Call unitsys_register(Me, lblCarbon(1), txtCarbon(1), txtCarbonUnits(1), "density", "g/mL", "g/mL", "", "", 100.0#, True)
		Call unitsys_register(Me, lblCarbon(2), txtCarbon(2), txtCarbonUnits(2), "length", "m", "m", "", "", 100.0#, True)
		Call unitsys_register(Me, lblCarbon(3), txtCarbon(3), Nothing, "", "", "", "", "", 100.0#, False)
		Call unitsys_register(Me, lblCarbon(4), txtCarbon(4), Nothing, "", "", "", "", "", 100.0#, False)
	End Sub


	Private Sub cmdADEComponent_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdADEComponent.Click
		Dim Index As Short = cmdADEComponent.GetIndex(eventSender)
		Dim Raise_Dirty_Flag As Boolean
		Dim temp As String
		Dim RetVal As Short
		Dim N As Short
		Dim i As Short
		Dim J As Short
		Select Case Index
			Case 0 'ADD.
				Call frmCompoProp.frmCompoProp_Add(Raise_Dirty_Flag)
				If (Raise_Dirty_Flag) Then
					'RAISE DIRTY FLAG.
					Call DirtyStatus_Throw()
				End If
			Case 1 'DELETE.
				If (cboSelectCompo.SelectedIndex = -1) Or (cboSelectCompo.SelectedIndex > Number_Component - 1) Then
					Call Show_Error("You must first select a component.")
					Exit Sub
				End If
				temp = Trim(VB6.GetItemString(cboSelectCompo, cboSelectCompo.SelectedIndex))
				RetVal = MsgBox("Do you really want to delete component '" & temp & "' ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Delete Component ?")
				If RetVal = MsgBoxResult.No Then Exit Sub
				N = cboSelectCompo.SelectedIndex + 1
				'
				' DELETE COMPONENT FROM MAIN COMPONENT PROPERTIES DATA AREA.
				'
				For i = N To Number_Component - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Component(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Component(i) = Component(i + 1)
					For J = 1 To 400
						C_Influent(i, J) = C_Influent(i + 1, J)
						C_Data_Points(i, J) = C_Data_Points(i + 1, J)
					Next J
				Next i
				Number_Component = Number_Component - 1
				'
				' DELETE COMPONENT FROM ROOM PROPERTIES DATA AREA.
				'
				For i = N To RoomParams.COUNT_CONTAMINANT - 1
					RoomParams.ROOM_C0(i) = RoomParams.ROOM_C0(i + 1)
					RoomParams.ROOM_EMIT(i) = RoomParams.ROOM_EMIT(i + 1)
					RoomParams.ROOM_SS_VALUE(i) = RoomParams.ROOM_SS_VALUE(i + 1)
					RoomParams.INITIAL_ROOM_CONC(i) = RoomParams.INITIAL_ROOM_CONC(i + 1)
					RoomParams.RXN_RATE_CONSTANT(i) = RoomParams.RXN_RATE_CONSTANT(i + 1)
					RoomParams.RXN_PRODUCT(i) = RoomParams.RXN_PRODUCT(i + 1)
					RoomParams.RXN_RATIO(i) = RoomParams.RXN_RATIO(i + 1)
				Next i
				RoomParams.COUNT_CONTAMINANT = RoomParams.COUNT_CONTAMINANT - 1
				'
				' RAISE DIRTY FLAG AND REFRESH MAIN WINDOW.
				'
				Call DirtyStatus_Throw()
				Call frmMain_Refresh()
			Case 2 'EDIT.
				If (cboSelectCompo.SelectedIndex = -1) Or (cboSelectCompo.SelectedIndex > Number_Component - 1) Then
					Call Show_Error("You must first select a component.")
					Exit Sub
				End If
				Call frmCompoProp.frmCompoProp_Edit(Raise_Dirty_Flag, cboSelectCompo.SelectedIndex + 1)
				If (Raise_Dirty_Flag) Then
					'RAISE DIRTY FLAG.
					Call DirtyStatus_Throw()
				End If
		End Select
	End Sub
	Private Sub cmdAdsorberDB_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdsorberDB.Click
		Dim User_Transferred_Data As Boolean
		Dim New_L As Double
		Dim New_D As Double
		Dim New_M As Double
		Dim New_Q As Double
		Call frmEditAdsorber.frmEditAdsorber_QueryDatabase(User_Transferred_Data)
		If (Not User_Transferred_Data) Then Exit Sub
		'TRANSFER DATA TO MAIN WINDOW.
		''SET L,D,M,Q UNITS TO M, M, KG, and M^3/S RESPECTIVELY
		'txtBedUnits(0).ListIndex = 0
		'txtBedUnits(1).ListIndex = 0
		'txtBedUnits(2).ListIndex = 0
		'txtBedUnits(3).ListIndex = 0
		'TRANSFER PARAMETERS BACK TO MAIN SCREEN
		New_L = frmEditAdsorber_ReturnParameters.L
		New_D = frmEditAdsorber_ReturnParameters.D
		New_M = frmEditAdsorber_ReturnParameters.M
		New_Q = frmEditAdsorber_ReturnParameters.Q
		'txtBedValue(0).Text = Format_It(New_L, 3)
		'txtBedValue(1).Text = Format_It(New_D, 3)
		'txtBedValue(2).Text = Format_It(New_M, 2)
		'txtBedValue(3).Text = Format_It(New_Q, 3)
		Bed.length = New_L
		Bed.Diameter = New_D
		Bed.Weight = New_M
		Bed.Flowrate = New_Q
		'RAISE DIRTY FLAG.
		Call DirtyStatus_Throw()
		'UPDATE WINDOW DISPLAY.
		Call frmMain_Refresh()
		''UPDATE SOME BED PROPERTY DISPLAYS:
		'Call Update_Bed_Density_Display
		'Call Update_Several_Bed_Properties(3)
	End Sub
	Private Sub cmdCarbon_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCarbon.Click
		Dim User_Transferred_Data As Boolean
		Call frmEditCarbon.frmEditCarbon_QueryDatabase(User_Transferred_Data)
		If (Not User_Transferred_Data) Then Exit Sub
		'
		'    DATA WAS ALREADY TRANSFERRED IN THE SUB-WINDOW.
		'    NO FURTHER DATA TRANSFER IS REQUIRED.
		'    ONLY THE DIRTY FLAG AND THE WINDOW DISPLAY
		'    UPDATE NEED BE PERFORMED (SEE BELOW).
		'
		'RAISE DIRTY FLAG.
		Call DirtyStatus_Throw()
		'UPDATE WINDOW DISPLAY.
		Call frmMain_Refresh()
	End Sub



	Private Sub cmdNote_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNote.Click
		Dim Index As Short = cmdNote.GetIndex(eventSender)
		Dim Temp_FileNote As String
		Dim RaiseDirtyFlag As Boolean
		Temp_FileNote = FileNote
		Call frmFileNote.frmFileNote_Run(Temp_FileNote, RaiseDirtyFlag)
		If (RaiseDirtyFlag) Then
			FileNote = Temp_FileNote
			'THROW DIRTY FLAG.
			Call DirtyStatus_Throw()
			'REFRESH WINDOW.
			Call frmMain_Refresh()
		End If
	End Sub
	Private Sub cmdParamsPSDMInRoom_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
		Dim Raise_Dirty_Flag As Boolean
		Call frmInputParamsPSDMInRoom.frmInputParamsPSDMInRoom_Edit(Raise_Dirty_Flag)
		If (Raise_Dirty_Flag) Then
			'THROW DIRTY FLAG.
			Call DirtyStatus_Throw()
		End If
	End Sub
	Private Sub cmdpolanyi_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdpolanyi.Click
		Dim Raise_Dirty_Flag As Boolean
		Call frmPolanyi.frmPolanyi_Edit(Me, Raise_Dirty_Flag)
		If (Raise_Dirty_Flag) Then
			'THROW DIRTY FLAG.
			Call DirtyStatus_Throw()
		End If
	End Sub
	Private Sub cmdViewDimensionless_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdViewDimensionless.Click
		frmDimensionless.ShowDialog()
	End Sub
	Private Sub cmdWaterCorrelations_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdWaterCorrelations.Click
		Dim Raise_Dirty_Flag As Boolean
		Call frmFluidProps.frmFluidProps_Edit(Raise_Dirty_Flag)
		If (Raise_Dirty_Flag) Then
			'THROW DIRTY FLAG.
			Call DirtyStatus_Throw()
		End If
	End Sub



	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub

	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim is_internal_mtu As Boolean
		Dim TurnOff_ForPSDMInRoom As Boolean
		'
		' MISC INITS.
		'
		rs.FindAllControls(Me)

		'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_Dirty.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'sspanel_Dirty.Caption = ""
		ToolStripStatusLabelDirty.Text = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_Status.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'sspanel_Status.Caption = ""
		ToolStripStatusLabelStatus.Text = ""
		Me.Text = AppName_For_Display_Short
		Me.Width = VB6.TwipsToPixelsX(9600)
		Me.Height = VB6.TwipsToPixelsY(7600)
		Call CenterOnScreen(Me)
		'UPGRADE_WARNING: Couldn't resolve default property of object CommonDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CommonDialog1.FileName = My.Application.Info.DirectoryPath & "\examples\*.dat"
		'OpenFileDialog1.FileName = My.Application.Info.DirectoryPath & "\examples\*.dat"
		OpenFileDialog1.FileName = My.Application.Info.DirectoryPath


		lblWaterUnit(0).Text = "C"
		lblWaterUnit(1).Text = "atm"
		cmdNote(1).SetBounds(cmdNote(0).Left, cmdNote(0).Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		cmdNote(0).Visible = False
		cmdNote(1).Visible = False
		'
		' CHECK FOR FILE THAT INDICATES THIS IS INTERNAL TO MTU:
		'
		is_internal_mtu = False
		If (check_internal_to_mtu()) Then is_internal_mtu = True
		mnuMTU.Visible = is_internal_mtu

		'///Modefication///Sinan///07/03/06, adding bouth the PSDM and the PSDM in room Models
		'for the Run menu.
		' PSDM IN ROOM INITS.
		'
		'  TurnOff_ForPSDMInRoom = False
		'  If (Activate_PSDMInRoom = True) Then
		'    TurnOff_ForPSDMInRoom = True
		'  End If
		'  If (is_internal_mtu = True) Then
		'    TurnOff_ForPSDMInRoom = False
		'  End If
		'  mnuRunItem(0).Visible = Not TurnOff_ForPSDMInRoom
		'  mnuRunItem(1).Visible = Not TurnOff_ForPSDMInRoom
		'  mnuRunItem(2).Visible = Not TurnOff_ForPSDMInRoom
		'  mnuRunItem(10).Visible = Activate_PSDMInRoom
		'  mnuRunItem(20).Visible = Activate_PSDMInRoom
		'  mnuResultsItem(1).Visible = Not TurnOff_ForPSDMInRoom
		'  mnuResultsItem(2).Visible = Not TurnOff_ForPSDMInRoom
		'  mnuResultsItem(4).Visible = Not TurnOff_ForPSDMInRoom
		'  If (is_internal_mtu = True) Then
		'    mnuRunItem(0).Caption = mnuRunItem(0).Caption & " (*)"
		'    mnuRunItem(1).Caption = mnuRunItem(1).Caption & " (*)"
		'    mnuRunItem(2).Caption = mnuRunItem(2).Caption & " (*)"
		'    mnuResultsItem(1).Caption = mnuResultsItem(1).Caption & " (*)"
		'    mnuResultsItem(2).Caption = mnuResultsItem(2).Caption & " (*)"
		'    mnuResultsItem(4).Caption = mnuResultsItem(4).Caption & " (*)"
		'  End If
		''''mnuResultsItem(10).Visible = Activate_PSDMInRoom
		'cmdParamsPSDMInRoom.Visible = Activate_PSDMInRoom
		'///End of Modefication.///

		' POPULATE UNITS INTO SCROLLBOX CONTROLS.
		'
		Call Populate_frmMain_Units()
		'
		' CREATE A NEW FILE IN MEMORY.
		'
		Call file_new()
		'
		' POPULATE LAST-FEW-FILES LIST.
		'
		Call OldFileList_Populate(1, Me._mnuFileItem_199, Me.mnuFileItem(191), Me.mnuFileItem(192), Me.mnuFileItem(193), Me.mnuFileItem(194))
	End Sub
	Private Sub frmMain_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		If (file_query_unload() = False) Then
			Cancel = True
		End If
		eventArgs.Cancel = Cancel
	End Sub
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call frmMain_Close_All_Windows()
		Call unitsys_unregister_all_on_form(Me)
	End Sub


	'UPGRADE_WARNING: Event lstComponents.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstComponents_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstComponents.SelectedIndexChanged
		Dim i As Short
		'Debug.Print "lstComponents_Click"
		'For i = 0 To lstComponents.ListCount - 1
		'  Debug.Print "lst.selected(" & Trim$(Str$(i)) & ") = " & _
		''      Trim$(Str$(lstComponents.Selected(i)))
		'Next i
		For i = 0 To lstComponents.Items.Count - 1
			Component(i + 1).Is_Selected_On_List = (lstComponents.GetSelected(i))
		Next i
	End Sub


	Public Sub mnuDatabasesItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDatabasesItem.Click
		Dim Index As Short = mnuDatabasesItem.GetIndex(eventSender)
		Select Case Index
			Case 0 'ADSORBENT DB.
				Call frmEditCarbon.frmEditCarbon_EditDatabase()
			Case 1 'ISOTHERM DB.
				Call frmEditIsotherm.frmEditIsotherm_EditDatabase()
			Case 2 'ADSORBER DB.
				Call frmEditAdsorber.frmEditAdsorber_EditDatabase()
		End Select
	End Sub

	'	Public Sub mnuFileItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileItem.Click
	'Dim Index As Short = mnuFileItem.GetIndex(eventSender)
	'Select Case Index
	'Case 0 'New
	'If (file_query_unload()) Then
	'Call Avoid_Weird_Focus_Problem()
	'Call file_new()
	'End If
	'Case 1 'Open ...
	'If (file_query_unload()) Then
	'Call Avoid_Weird_Focus_Problem()
	'Call File_OpenAs("")
	'End If
	'Case 2 'Save
	'If (Filename = "") Then
	'Call Avoid_Weird_Focus_Problem()
	'Call File_SaveAs("")
	'Else
	'Call Avoid_Weird_Focus_Problem()
	'Call File_SaveAs(Filename)
	'End If
	'Case 3 'Save As ...
	'Call Avoid_Weird_Focus_Problem()
	'Call File_SaveAs("")
	'Case 6 'Select Printer ...
	'			'UPGRADE_WARNING: Couldn't resolve default property of object CommonDialog1.ShowPrinter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			CommonDialog1.ShowPrinter()
	'			'Case 85:      'Print ...
	'			'  frmPrint.Show 1
	'Case 191 To 194 'Last-few-files list
	'If (file_query_unload()) Then
	'If (mnuFileItem(Index).Visible) Then
	'Call File_OpenAs(OldFiles(1, Index - 190))
	'End If
	'End If
	'Case 200 'Exit
	'  'NOTE: MDIForm_QueryUnload() TAKES CAKE OF THIS.
	'  'If we do it here, _two_ message boxes will pop up
	'   'when the user has data which needs saving !
	'  'If (file_query_unload()) Then
	'  '  Unload Me
	'  'End If
	'  Me.Close()
	'End Select
	'End Sub

	Public Sub mnuHelpItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpItem.Click
		Dim Index As Short = mnuHelpItem.GetIndex(eventSender)
		Dim fn_This As String
		Select Case Index
			Case 10 'ONLINE HELP.
				'NOTE: We currently do NOT have the resources to
				'create an online help file for the program (1/7/98)
				'therefore no online help is available.
				Call Show_Message("Online help is currently unavailable.  " & "Please refer to the printed manual or the Acrobat-format ADS.PDF file.")
				Exit Sub
				'Call LaunchFile_General("", MAIN_APP_PATH & "\help\ads.hlp")
			Case 20 'ONLINE MANUAL.
				''''fn_This = MAIN_APP_PATH & "\help\ads.pdf"
				fn_This = MAIN_APP_PATH & "\help\ads.doc"
				If (FileExists(fn_This) = False) Then
					Call Show_Message("The file `" & fn_This & "` is missing.")
					Exit Sub
				End If
				Call ShellExecute_LocalFile(fn_This)
				''''Call LaunchFile_General("", fn_This)
			Case 22 'MANUAL PRINTING INSTRUCTIONS.
				fn_This = Global_fpath_dir_CPAS & "\dbase\printing.txt"
				If (FileExists(fn_This) = False) Then
					Call Show_Message("The file `" & fn_This & "` is missing.")
					Exit Sub
				End If
				Call Launch_Notepad(fn_This)
			Case 80
				fn_This = My.Application.Info.DirectoryPath & "\dbase\readme.txt"
				If (FileExists(fn_This) = False) Then
					Call Show_Message("The file `" & fn_This & "` is missing.")
					Exit Sub
				End If
				Call Launch_Notepad(fn_This)
			Case 85 'VIEW DISCLAIMER.
				'SHOW THE DISCLAIMER WINDOW.
				splash_mode = 101
				splash_button_pressed = 0
				frmSplash.ShowDialog()
			Case 90 'TECH ASSISTANCE.
				frmAbout2.ShowDialog()
			Case 99 'ABOUT.
				frmAbout.ShowDialog()
		End Select
	End Sub
	Public Sub mnuMTUItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMTUItem.Click
		Dim Index As Short = mnuMTUItem.GetIndex(eventSender)
		Select Case Index
			Case 40 'KEEP TEMPORARY MODEL FILES.
				mnuMTUItem(40).Checked = Not mnuMTUItem(40).Checked
		End Select
	End Sub
	Public Sub mnuOptionsItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOptionsItem.Click
		Dim Index As Short = mnuOptionsItem.GetIndex(eventSender)
		Dim msg As String
		Dim i As Short
		Dim J As Short
		Dim Raise_Dirty_Flag As Boolean
		Select Case Index
			Case 0 'FOULING.
				''''frmFouling.Show 1
				Call frmFouling.frmFouling_Go(Raise_Dirty_Flag)
				If (Raise_Dirty_Flag) Then
					'RAISE DIRTY FLAG.
					Call DirtyStatus_Throw()
				End If

			Case 1 'INFLUENT CONCENTRATIONS.
				'---- Options--Influent Concentrations
				'-- Setup global variables to make the call
				frmConcentrations_caption = "Influent Concentrations"
				frmConcentrations_Cunits = "mg/L"
				frmConcentrations_Tunits = "days"
				frmConcentrations_TimeOrderImportant = True
				frmConcentrations_NumPoints = Number_Influent_Points
				frmConcentrations_NumConcs = Number_Component
				For i = 1 To frmConcentrations_NumPoints
					frmConcentrations_Times(i) = T_Influent(i)
					For J = 1 To frmConcentrations_NumConcs
						frmConcentrations_Concs(J, i) = C_Influent(J, i)
					Next J
				Next i
				'-- Make the call
				frmVarConcentrations.ShowDialog()
				'-- Reinitialize Last-Few-Files list after frmConcentrations is done
				'xaxaxaNC
				'Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADSIM, LASTFEW_ADSIM_frmPFPSDM)
				''''Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADXDESIGNS, LASTFEW_ADXDESIGNS_frmPFPSDM)
				If (Not frmConcentrations_cancelled) Then
					Number_Influent_Points = frmConcentrations_NumPoints
					Number_Component = frmConcentrations_NumConcs
					For i = 1 To frmConcentrations_NumPoints
						T_Influent(i) = frmConcentrations_Times(i)
						For J = 1 To frmConcentrations_NumConcs
							C_Influent(J, i) = frmConcentrations_Concs(J, i)
						Next J
					Next i
				End If
				frmVarConcentrations.Sheet1DataGrid.Rows.Clear()

			Case 2 'EFFLUENT CONCENTRATIONS.
				'---- Options--Effluent Concentrations
				'-- Setup global variables to make the call
				frmConcentrations_caption = "Effluent Concentrations"
				frmConcentrations_Cunits = "C/C0"
				frmConcentrations_Tunits = "days"
				frmConcentrations_TimeOrderImportant = False
				frmConcentrations_NumPoints = NData_Points
				frmConcentrations_NumConcs = Number_Component
				For i = 1 To frmConcentrations_NumPoints
					frmConcentrations_Times(i) = T_Data_Points(i) * 24.0# * 60.0#
					For J = 1 To frmConcentrations_NumConcs
						frmConcentrations_Concs(J, i) = C_Data_Points(J, i)
					Next J
				Next i
				'-- Make the call
				frmVarConcentrations.ShowDialog()
				'-- Reinitialize Last-Few-Files list after frmConcentrations is done
				'xaxaxaNC
				'Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADSIM, LASTFEW_ADSIM_frmPFPSDM)
				''''Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADXDESIGNS, LASTFEW_ADXDESIGNS_frmPFPSDM)
				If (Not frmConcentrations_cancelled) Then
					NData_Points = frmConcentrations_NumPoints
					If NData_Points > 0 And mnuResultsItem(1).Enabled = True Then
						mnuResultsItem(4).Enabled = True
					End If
					If NData_Points > 0 And mnuResultsItem(0).Enabled = True Then
						mnuResultsItem(3).Enabled = True
					End If
					Number_Component = frmConcentrations_NumConcs
					For i = 1 To frmConcentrations_NumPoints
						T_Data_Points(i) = frmConcentrations_Times(i) / 24.0# / 60.0#
						For J = 1 To frmConcentrations_NumConcs
							C_Data_Points(J, i) = frmConcentrations_Concs(J, i)
						Next J
					Next i
				End If
				frmVarConcentrations.Sheet1DataGrid.Rows.Clear()
		End Select
	End Sub
	Public Sub mnuPhaseItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPhaseItem.Click
		Dim Index As Short = mnuPhaseItem.GetIndex(eventSender)
		Dim OldBedPhase As Short
		OldBedPhase = Bed.Phase
		Select Case Index
			Case 0 'LIQUID PHASE.
				Call chem_phase(0)
			Case 1 'GAS PHASE.
				Call chem_phase(1)
		End Select
		If (Bed.Phase <> OldBedPhase) Then
			'THROW DIRTY FLAG AND REFRESH WINDOW.
			Call DirtyStatus_Throw()
			Call frmMain_Refresh()
		End If
	End Sub
	Public Sub mnuPrintSubItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPrintSubItem.Click
		Dim Index As Short = mnuPrintSubItem.GetIndex(eventSender)
		Select Case Index
			Case 0 'PRINT-TO-PRINTER.
				Print_To_Printer = True
				frmPrintInputs.ShowDialog()
			Case 1 'PRINT-TO-FILE.
				Print_To_Printer = False
				frmPrintInputs.ShowDialog()
		End Select
	End Sub
	Public Sub mnuResultsItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuResultsItem.Click
		Dim Index As Short = mnuResultsItem.GetIndex(eventSender)
		Select Case Index
			Case 0 'PSDM.
				frmModelPSDMResults.ShowDialog()
			Case 1 'CPHSDM.
				frmModelCPHSDMResults.ShowDialog()
			Case 2 'ECM.
				frmModelECMResults.ShowDialog()
			Case 3
				frmCompareData_WhichSet = frmCompareData_WhichSet_PSDM
				frmCompareData_caption = "Comparison of PSDM Results with Effluent Data"
				frmModelDataComparison.ShowDialog()
			Case 4
				frmCompareData_WhichSet = frmCompareData_WhichSet_CPHSDM
				frmCompareData_caption = "Comparison of CPHSDM Results with Effluent Data"
				frmModelDataComparison.ShowDialog()
		End Select
	End Sub
	Public Sub mnuRunItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRunItem.Click
		Dim Index As Short = mnuRunItem.GetIndex(eventSender)
		Dim i As Short
		Dim J As Short
		Dim Num_K_Reduction As Short
		Select Case Index
			'
			'///////////////////////////////////////////////////////////////////////////////////////////////////
			'/////////    PSDM
			'///////////////////////////////////////////////////////////////////////////////////////////////////
			Case 0
				If (Prepare_To_Run_PSDM() = False) Then
					Exit Sub
				End If
				'RUN THE MODEL.
				Call ModelPSDM_Go()
				'
				'///////////////////////////////////////////////////////////////////////////////////////////////////
				'/////////    PSDMR in Room
				'///////////////////////////////////////////////////////////////////////////////////////////////////
			Case 10
				If (Prepare_To_Run_PSDM_In_Room() = False) Then
					Exit Sub
				End If
				'RUN THE MODEL.
				Call ModelPSDMInRoom_Go(PSDMR_MODE_INROOM)
				'
				'///////////////////////////////////////////////////////////////////////////////////////////////////
				'/////////    PSDMR Alone
				'///////////////////////////////////////////////////////////////////////////////////////////////////
			Case 20
				If (Prepare_To_Run_PSDM_In_Room() = False) Then
					Exit Sub
				End If
				'RUN THE MODEL.
				Call ModelPSDMInRoom_Go(PSDMR_MODE_ALONE)
				'
				'///////////////////////////////////////////////////////////////////////////////////////////////////
				'/////////    CPHSDM
				'///////////////////////////////////////////////////////////////////////////////////////////////////
			Case 1
				'---- Make sure # fouling components is = 0.
				Num_K_Reduction = 0
				For i = 0 To lstComponents.Items.Count - 1
					If (lstComponents.GetSelected(i)) Then
						If (Component(i + 1).K_Reduction) Then
							Num_K_Reduction = Num_K_Reduction + 1
						End If
					End If
				Next i
				If (Num_K_Reduction > 0) Then
					Call Show_Message("Warning: There are currently " & Trim(Str(Num_K_Reduction)) & " components specified with fouling correlations.  The CPHSDM model " & "does not use the fouling correlations and will ignore them.")
				End If
				Call AllModels_Verify_Selected_Components(MODELTYPE_CPHSDM)
				If (Number_Component_CPM = 0) Then
					Exit Sub 'ERROR MESSAGE HANDLED IN AllModels_... SUBROUTINE.
				End If
				'RUN THE MODEL.
				Call ModelCPHSDM_Go()
				'
				'///////////////////////////////////////////////////////////////////////////////////////////////////
				'/////////    ECM
				'///////////////////////////////////////////////////////////////////////////////////////////////////
			Case 2
				'---- Make sure # fouling components is = 0.
				Num_K_Reduction = 0
				For i = 0 To lstComponents.Items.Count - 1
					If (lstComponents.GetSelected(i)) Then
						If (Component(i + 1).K_Reduction) Then
							Num_K_Reduction = Num_K_Reduction + 1
						End If
					End If
				Next i
				If (Num_K_Reduction > 0) Then
					Call Show_Message("Warning: There are currently " & Trim(Str(Num_K_Reduction)) & " components specified with fouling correlations.  The ECM model " & "does not use the fouling correlations and will ignore them.")
				End If
				Call AllModels_Verify_Selected_Components(MODELTYPE_ECM)
				If (Number_Component_ECM = 0) Then
					Exit Sub 'ERROR MESSAGE HANDLED IN AllModels_... SUBROUTINE.
				End If
				For i = 1 To Number_Component_ECM
					For J = i + 1 To Number_Component_ECM
						'UPGRADE_WARNING: Couldn't resolve default property of object Component_Index_ECM(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (Trim(Component(Component_Index_ECM(i)).Name) = Trim(Component(Component_Index_ECM(J)).Name)) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Component_Index_ECM(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Call Show_Error("Components " & VB6.Format(Component_Index_ECM(i), "0") & " and " & VB6.Format(Component_Index_ECM(J), "0") & " have the same name." & vbCrLf & "Please change one before running the ECM.")
							Exit Sub
						End If
					Next J
				Next i
				'RUN THE MODEL.
				Call ModelECM_Go()
		End Select
	End Sub


	Private Sub spnNumberOfBeds_SpinDown()
		If (Bed.NumberOfBeds > 1) Then
			Bed.NumberOfBeds = Bed.NumberOfBeds - 1
			'THROW DIRTY FLAG AND REFRESH WINDOW.
			Call DirtyStatus_Throw()
			Call frmMain_Refresh()
		End If
	End Sub
	Private Sub spnNumberOfBeds_SpinUp()
		If (Bed.NumberOfBeds < Maximum_Beds_In_Series) Then
			Bed.NumberOfBeds = Bed.NumberOfBeds + 1
			'THROW DIRTY FLAG AND REFRESH WINDOW.
			Call DirtyStatus_Throw()
			Call frmMain_Refresh()
		End If
	End Sub


	'Private Sub spnPoint_SpinDown(ByRef Index As Short)
	'Select Case Index
	'Case 0 'AXIAL POINTS.
	'If (MC > 1) Then
	'				MC = MC - 1
	''THROW DIRTY FLAG AND REFRESH WINDOW.
	'Call DirtyStatus_Throw()
	'Call frmMain_Refresh()
	'End If
	'Case 1 'RADIAL POINTS.
	'If (NC > 1) Then
	'				NC = NC - 1
	''THROW DIRTY FLAG AND REFRESH WINDOW.
	'Call DirtyStatus_Throw()
	'Call frmMain_Refresh()
	'End If
	'End Select
	'End Sub
	'Private Sub spnPoint_SpinUp(ByRef Index As Short)
	'Select Case Index
	'Case 0 'AXIAL POINTS.
	'If (MC < Max_Axial_Collocation) Then
	'				MC = MC + 1
	''THROW DIRTY FLAG AND REFRESH WINDOW.
	'Call DirtyStatus_Throw()
	'Call frmMain_Refresh()
	'End If
	'Case 1 'RADIAL POINTS.
	'If (NC < Max_Radial_Collocation) Then
	'				NC = NC + 1
	'THROW DIRTY FLAG AND REFRESH WINDOW.
	'Call DirtyStatus_Throw()
	'Call frmMain_Refresh()
	'End If
	'End Select
	'End Sub


	'UPGRADE_WARNING: Event txtBedUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBedUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBedUnits.SelectedIndexChanged
		Dim Index As Short = txtBedUnits.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtBedUnits(Index)
		Call unitsys_control_cbox_click(Ctl)
	End Sub
	Private Sub txtBedUnits_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBedUnits.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtBedUnits.GetIndex(eventSender)
		KeyAscii = Global_TextKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub


	Private Sub txtBedValue_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBedValue.Enter
		Dim Index As Short = txtBedValue.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtBedValue(Index)
		Dim StatusMessagePanel As String
		Call unitsys_control_txtx_gotfocus(Ctl)
		Select Case Index
			Case 0
				StatusMessagePanel = "Type in the bed length"
			Case 1
				StatusMessagePanel = "Type in the bed diameter"
			Case 2
				StatusMessagePanel = "Type in the mass of adsorbent in the bed"
			Case 3
				StatusMessagePanel = "Type in the inlet flowrate"
			Case 4
				StatusMessagePanel = "Type in the Empty Bed Contact Time (EBCT)"
		End Select
		Call GenericStatus_Set(StatusMessagePanel)
	End Sub
	Private Sub txtBedValue_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBedValue.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtBedValue.GetIndex(eventSender)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtBedValue_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBedValue.Leave
		Dim Index As Short = txtBedValue.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtBedValue(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS
		If (Index = 4) Then
			Val_Low = 1.0E-20 * 60.0#
			Val_High = 1.0E+20 * 60.0#
		Else
			Val_Low = 1.0E-20
			Val_High = 1.0E+20
		End If
		NewValue_Okay = False
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		Call GenericStatus_Set("")
		If (NewValue_Okay) Then
			If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
				''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
				Raise_Dirty_Flag = False
			End If
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				Select Case Index
					Case 0 'BED LENGTH.
						Call Check_Length(NewValue, Too_Small)
						Raise_Dirty_Flag = False
						If Not (Too_Small) Then
							Bed.length = NewValue
							Raise_Dirty_Flag = True
						End If
						'Call Update_Display
						Call Update_KP_Values()
						'Call Update_Bed_Density_Display
						'Call Update_Several_Bed_Properties(1)
					Case 1 'BED DIAMETER.
						Call Check_Diameter(NewValue, Too_Small)
						Raise_Dirty_Flag = False
						If Not (Too_Small) Then
							Bed.Diameter = NewValue
							Raise_Dirty_Flag = True
						End If
						'Call Update_Display
						Call Update_KP_Values()
						'Call Update_Bed_Density_Display
						'Call Update_Several_Bed_Properties(3)
					Case 2 'BED MASS.
						Call Check_Weight(NewValue, Too_Small)
						Raise_Dirty_Flag = False
						If Not (Too_Small) Then
							Bed.Weight = NewValue
							Raise_Dirty_Flag = True
						End If
						''Call Update_Display
						Call Update_KP_Values()
						'Call Update_Bed_Density_Display
						'Call Update_Several_Bed_Properties(1)
					Case 3 'BED FLOW RATE.
						Bed.Flowrate = NewValue
						''Call Update_Display       'Updates display of flowrate and EBCT.
						Call Update_KP_Values()
						'Call Update_Several_Bed_Properties(2)
					Case 4 'BED EBCT.
						Bed.Flowrate = Bed.length * PI * Bed.Diameter * Bed.Diameter / 4.0# / NewValue 'EBCT in sec
						''Call Update_Display       'Updates display of flowrate and EBCT.
						Call Update_KP_Values()
						'Call Update_Several_Bed_Properties(2)
				End Select
				If (Raise_Dirty_Flag) Then
					'THROW DIRTY FLAG.
					Call DirtyStatus_Throw()
				End If
			End If
			'REFRESH WINDOW.
			Call frmMain_Refresh()
		End If
	End Sub


	Private Sub txtCarbon_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCarbon.Enter
		Dim Index As Short = txtCarbon.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtCarbon(Index)
		Dim StatusMessagePanel As String
		If (Index = 0) Then
			Call Global_GotFocus(Ctl)
		Else
			Call unitsys_control_txtx_gotfocus(Ctl)
		End If
		Select Case Index
			Case 0
				StatusMessagePanel = "Type in the adsorbent name"
			Case 1
				StatusMessagePanel = "Type in the adsorbent density (that includes pore volume)"
			Case 2
				StatusMessagePanel = "Type in the average particle radius"
			Case 3
				StatusMessagePanel = "Type in the particle porosity"
			Case 4
				StatusMessagePanel = "Type in the particle shape factor (spheres=1.0)"
				'    Case 4
				'      StatusMessagePanel = " Type in the particle tortuosity"
		End Select
		Call GenericStatus_Set(StatusMessagePanel)
	End Sub
	Private Sub txtCarbon_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCarbon.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtCarbon.GetIndex(eventSender)
		If (Index = 0) Then
			KeyAscii = Global_TextKeyPress(KeyAscii)
		Else
			KeyAscii = Global_NumericKeyPress(KeyAscii)
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtCarbon_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCarbon.Leave
		Dim Index As Short = txtCarbon.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtCarbon(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		'HANDLE THE CARBON NAME TEXTBOX.
		If (Index = 0) Then
			If (Trim(Ctl.Text) = "") Then
				Ctl.Text = Carbon.Name
				'Call Show_Error("You must enter a non-blank string for the carbon name.")
				'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
				'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
			Else
				If (Trim(Carbon.Name) <> Trim(Ctl.Text)) Then
					Carbon.Name = Trim(Ctl.Text)
					'THROW DIRTY FLAG.
					Call DirtyStatus_Throw()
				End If
			End If
			Call Global_LostFocus(Ctl)
			Call GenericStatus_Set("")
			Exit Sub
		End If
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS
		Val_Low = 1.0E-20
		Val_High = 1.0E+20
		NewValue_Okay = False
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		Call GenericStatus_Set("")
		If (NewValue_Okay) Then
			If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
				''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
				Raise_Dirty_Flag = False
			End If
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				Select Case Index
					Case 1 'APPARENT DENSITY.
						Call Check_Density(NewValue, Too_Small)
						Raise_Dirty_Flag = False
						If Not (Too_Small) Then
							Carbon.Density = NewValue
							Raise_Dirty_Flag = True
						End If
						'Call Update_Display
						Call Update_KP_Values()
						'Call Update_Bed_Density_Display
						'Call Update_Several_Bed_Properties(1)
					Case 2 'PARTICLE RADIUS.
						Carbon.ParticleRadius = NewValue
						'Call Update_Display
						Call Update_KP_Values()
					Case 3 'POROSITY.
						Carbon.Porosity = NewValue
						'Call Update_Display
						Call Update_KP_Values()
					Case 4 'PARTICLE SHAPE FACTOR.
						Carbon.ShapeFactor = NewValue
						'Call Update_Display
						Call Update_KP_Values()
				End Select
				If (Raise_Dirty_Flag) Then
					'THROW DIRTY FLAG.
					Call DirtyStatus_Throw()
				End If
			End If
			'REFRESH WINDOW.
			Call frmMain_Refresh()
		End If
	End Sub


	'UPGRADE_WARNING: Event txtCarbonUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtCarbonUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCarbonUnits.SelectedIndexChanged
		Dim Index As Short = txtCarbonUnits.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtCarbonUnits(Index)
		Call unitsys_control_cbox_click(Ctl)
	End Sub
	Private Sub txtCarbonUnits_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCarbonUnits.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtCarbonUnits.GetIndex(eventSender)
		KeyAscii = Global_TextKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub


	Private Sub txtNPoint_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNPoint.Enter
		Dim Index As Short = txtNPoint.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtNPoint(Index)
		Dim StatusMessagePanel As String
		Call unitsys_control_txtx_gotfocus(Ctl)
		Select Case Index
			Case 0
				StatusMessagePanel = "Type in the number of collocation points in the axial direction (PSDM only)"
			Case 1
				StatusMessagePanel = "Type in the number of collocation points in the radial direction (PSDM only)"
		End Select
		Call GenericStatus_Set(StatusMessagePanel)
	End Sub
	'Private Sub txtNPoint_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNPoint.KeyPress
	'Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
	'Dim Index As Short = txtNPoint.GetIndex(eventSender)
	'	KeyAscii = Global_NumericKeyPress(KeyAscii)
	'	eventArgs.KeyChar = Chr(KeyAscii)
	'If KeyAscii = 0 Then
	'		eventArgs.Handled = True
	'End If
	'End Sub
	'Private Sub txtNPoint_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNPoint.Leave
	'Dim Index As Short = txtNPoint.GetIndex(eventSender)
	'Dim NewValue_Okay As Short
	'Dim NewValue As Double
	'Dim Ctl As System.Windows.Forms.Control
	'	Ctl = txtNPoint(Index)
	'Dim Val_Low As Double
	'Dim Val_High As Double
	'Dim Raise_Dirty_Flag As Boolean
	'Dim Too_Small As Short
	''NOTE: LOW AND HIGH VALUES IN BASE UNITS
	'Select Case Index
	'Case 0 'AXIAL DIR.
	'			Val_Low = 1.0#
	'			Val_High = CDbl(Max_Axial_Collocation)
	'Case 1 'RADIAL DIR.
	'			Val_Low = 1.0#
	'			Val_High = CDbl(Max_Radial_Collocation)
	'End Select
	'	NewValue_Okay = False
	'If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
	'		NewValue_Okay = True
	'End If
	'Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
	'Call GenericStatus_Set("")
	'If (NewValue_Okay) Then
	'If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
	'			''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
	'			Raise_Dirty_Flag = False
	'End If
	'If (Raise_Dirty_Flag) Then
	'STORE TO MEMORY.
	'Select Case Index
	'Case 0 'AXIAL DIR.
	'					MC = CShort(NewValue)
	'Case 1 'RADIAL DIR.
	'					NC = CShort(NewValue)
	'End Select
	'If (Raise_Dirty_Flag) Then
	''THROW DIRTY FLAG.
	'Call DirtyStatus_Throw()
	'End If
	'End If
	'REFRESH WINDOW.
	'Call frmMain_Refresh()
	'End If
	'End Sub


	'Private Sub txtNumberOfBeds_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
	'Dim Ctl As System.Windows.Forms.Control
	'	Ctl = txtNumberOfBeds
	'Dim StatusMessagePanel As String
	'Call unitsys_control_txtx_gotfocus(Ctl)
	'	StatusMessagePanel = "Type in the number of axial elements (PSDM only)"
	'Call GenericStatus_Set(StatusMessagePanel)
	'End Sub
	'Private Sub txtNumberOfBeds_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
	'Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
	'	KeyAscii = Global_NumericKeyPress(KeyAscii)
	'	eventArgs.KeyChar = Chr(KeyAscii)
	'If KeyAscii = 0 Then
	'		eventArgs.Handled = True
	'End If
	'End Sub
	'Private Sub txtNumberOfBeds_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
	'Dim NewValue_Okay As Short
	'Dim NewValue As Double
	'Dim Ctl As System.Windows.Forms.Control
	'	Ctl = txtNumberOfBeds
	'Dim Val_Low As Double
	'Dim Val_High As Double
	'Dim Raise_Dirty_Flag As Boolean
	'Dim Too_Small As Short
	'	'NOTE: LOW AND HIGH VALUES IN BASE UNITS
	'	Val_Low = 1.0#
	'	Val_High = CDbl(Maximum_Beds_In_Series)
	'	NewValue_Okay = False
	'If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
	'		NewValue_Okay = True
	'End If
	'Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
	'Call GenericStatus_Set("")
	'If (NewValue_Okay) Then
	'If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
	'			''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
	'			Raise_Dirty_Flag = False
	'End If
	'If (Raise_Dirty_Flag) Then
	'			'STORE TO MEMORY.
	'			Bed.NumberOfBeds = CShort(NewValue)
	'If (Raise_Dirty_Flag) Then
	''THROW DIRTY FLAG.
	'Call DirtyStatus_Throw()
	'End If
	'End If
	'REFRESH WINDOW.
	'Call frmMain_Refresh()
	'End If
	'End Sub


	Private Sub txtTime_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTime.Enter
		Dim Index As Short = txtTime.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtTime(Index)
		Dim StatusMessagePanel As String
		Call unitsys_control_txtx_gotfocus(Ctl)
		Select Case Index
			Case 0
				StatusMessagePanel = "Type in the total run time of the fixed bed adsorber (PSDM only)"
			Case 1
				StatusMessagePanel = "Type in the time of the first point to be displayed (PSDM only)"
			Case 2
				StatusMessagePanel = "Type in the time step (PSDM only)"
		End Select
		Call GenericStatus_Set(StatusMessagePanel)
	End Sub
	Private Sub txtTime_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTime.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtTime.GetIndex(eventSender)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtTime_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTime.Leave
		Dim Index As Short = txtTime.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtTime(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		Dim ForceAbort As Short
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS
		Select Case Index
			Case 0, 2 'TOTAL RUN TIME, TIME STEP.
				Val_Low = 1.0E-20
				Val_High = 1.0E+20
			Case 1 'FIRST POINT DISPLAYED.
				Val_Low = 0#
				Val_High = 1.0E+20
		End Select
		NewValue_Okay = False
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		Call GenericStatus_Set("")
		If (NewValue_Okay) Then
			If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
				''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
				Raise_Dirty_Flag = False
			End If
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				Select Case Index
					Case 0 'TOTAL RUN TIME.
						Call Check_Time_Parameters(0, NewValue, ForceAbort)
						If Not (ForceAbort) Then
							TimeP.End_Renamed = NewValue
						End If
					Case 1 'FIRST POINT DISPLAYED.
						Call Check_Time_Parameters(1, NewValue, ForceAbort)
						If Not (ForceAbort) Then
							TimeP.Init = NewValue
						End If
					Case 2 'TIME STEP.
						Call Check_Time_Parameters(2, NewValue, ForceAbort)
						If Not (ForceAbort) Then
							TimeP.Step_Renamed = NewValue
						End If
				End Select
				If (Raise_Dirty_Flag) Then
					'THROW DIRTY FLAG.
					Call DirtyStatus_Throw()
				End If
			End If
			'REFRESH WINDOW.
			Call frmMain_Refresh()
		End If
	End Sub


	'UPGRADE_WARNING: Event txttimeunits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txttimeunits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTimeUnits.SelectedIndexChanged
		Dim Index As Short = txtTimeUnits.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtTimeUnits(Index)
		Call unitsys_control_cbox_click(Ctl)
	End Sub
	Private Sub txttimeunits_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTimeUnits.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtTimeUnits.GetIndex(eventSender)
		KeyAscii = Global_TextKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub


	Private Sub txtWater_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWater.Enter
		Dim Index As Short = txtWater.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtTime(Index)
		Dim StatusMessagePanel As String
		Call unitsys_control_txtx_gotfocus(txtWater(Index))
		Select Case Index
			Case 0
				StatusMessagePanel = "Type in the Fluid Temperature"
			Case 1
				StatusMessagePanel = "Type in the Fluid Pressure"
		End Select
		Call GenericStatus_Set(StatusMessagePanel)
	End Sub
	Private Sub txtWater_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWater.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtWater.GetIndex(eventSender)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtWater_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWater.Leave
		Dim Index As Short = txtWater.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtWater(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		Dim ForceAbort As Short
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS
		Select Case Index
			Case 0 'TEMPERATURE (degC).
				Val_Low = 0.01
				Val_High = 100.0#
			Case 1 'PRESSURE (atm).
				Val_Low = 0.001
				Val_High = 100.0#
		End Select
		NewValue_Okay = False
		Raise_Dirty_Flag = True
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		Call GenericStatus_Set("")
		If (NewValue_Okay) Then
			If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
				''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
				Raise_Dirty_Flag = False
			End If
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				Select Case Index
					Case 0 'TEMPERATURE (degC).
						Bed.Temperature = NewValue
						'Call Update_Display_Water
					Case 1 'PRESSURE (atm).
						Bed.Pressure = NewValue
						'Call Update_Display_Water
				End Select
				Call Update_FluidDensity(Bed.Temperature, Bed.Pressure, Bed.WaterDensity)
				Call Update_FluidViscosity(Bed.Temperature, Bed.WaterViscosity)
				If (Raise_Dirty_Flag) Then
					'THROW DIRTY FLAG.
					Call DirtyStatus_Throw()
				End If
			End If
			'REFRESH WINDOW.
			Call frmMain_Refresh()
		End If
	End Sub

	Private Sub mnuFile_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles mnuFile.DropDownItemClicked
		Dim Index As Short = mnuFileItem.GetIndex(e.ClickedItem)
		mnuFile.HideDropDown()
		Select Case Index
			Case 0 'New
				If (file_query_unload()) Then
					Call Avoid_Weird_Focus_Problem()
					Call file_new()
				End If
			Case 1 'Open ...
				If (file_query_unload()) Then
					Call Avoid_Weird_Focus_Problem()
					Call File_OpenAs("")
				End If
			Case 2 'Save
				If (Filename = "") Then
					Call Avoid_Weird_Focus_Problem()
					Call File_SaveAs("")
				Else
					Call Avoid_Weird_Focus_Problem()
					Call File_SaveAs(Filename)
				End If
			Case 3 'Save As ...
				Call Avoid_Weird_Focus_Problem()
				Call File_SaveAs("")
			Case 6 'Select Printer ...
				'UPGRADE_WARNING: Couldn't resolve default property of object CommonDialog1.ShowPrinter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'CommonDialog1.ShowPrinter()
				'Case 85:      'Print ...
				'  frmPrint.Show 1
			Case 191 To 194 'Last-few-files list
				If (file_query_unload()) Then
					'If (mnuFileItem(Index).Visible) Then
					If (True) Then
						Call File_OpenAs(OldFiles(1, Index - 190))
					End If
				End If
			Case 200 'Exit
				'NOTE: MDIForm_QueryUnload() TAKES CAKE OF THIS.
				'If we do it here, _two_ message boxes will pop up
				'when the user has data which needs saving !
				'If (file_query_unload()) Then
				'  Unload Me
				'End If
				Me.Close()
		End Select
	End Sub

	'Private Sub spnNumberOfBeds_SpinUp(sender As Object, e As EventArgs)
	'Call spnNumberOfBeds_SpinUp()
	'End Sub

	'Private Sub spnNumberOfBeds_SpinDown(sender As Object, e As EventArgs)
	'Call spnNumberOfBeds_SpinDown()
	'End Sub

	'Private Sub _spnPoint_0_SpinDown(sender As Object, e As EventArgs) Handles _spnPoint_0.SpinDown
	'Call spnPoint_SpinDown(0)
	'End Sub

	'Private Sub _spnPoint_0_SpinUp(sender As Object, e As EventArgs) Handles _spnPoint_0.SpinUp
	'Call spnPoint_SpinUp(0)
	'End Sub

	'Private Sub _spnPoint_1_SpinDown(sender As Object, e As EventArgs) Handles _spnPoint_1.SpinDown
	'Call spnPoint_SpinDown(1)
	'End Sub

	'Private Sub _spnPoint_1_SpinUp(sender As Object, e As EventArgs) Handles _spnPoint_1.SpinUp
	'Call spnPoint_SpinUp(1)
	'End Sub

	Private Sub ssframe_FixedBed_Enter(sender As Object, e As EventArgs)

	End Sub
	Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
		'this numericupdown1 is used for "number of radial elements"
		Bed.NumberOfBeds = NumericUpDown1.Value

	End Sub

	Private Sub NumericUpDown1_Leave(sender As Object, e As EventArgs) Handles NumericUpDown1.Leave
		Call DirtyStatus_Throw()
		Call frmMain_Refresh()
	End Sub

	Private Sub NumericUpDown1_Enter(sender As Object, e As EventArgs) Handles NumericUpDown1.Enter
		Dim StatusMessagePanel As String
		StatusMessagePanel = "Type in the number of axial elements (PSDM only)"
		Call GenericStatus_Set(StatusMessagePanel)
	End Sub

	Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
		'axial points
		MC = NumericUpDown2.Value

	End Sub

	Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
		'radial points
		NC = NumericUpDown3.Value
	End Sub

	Private Sub NumericUpDown2_Enter(sender As Object, e As EventArgs) Handles NumericUpDown2.Enter
		Dim StatusMessagePanel As String
		StatusMessagePanel = "Type in the number of collocation points in the axial direction (PSDM only)"
		Call GenericStatus_Set(StatusMessagePanel)
	End Sub

	Private Sub NumericUpDown3_Enter(sender As Object, e As EventArgs) Handles NumericUpDown3.Enter
		Dim StatusMessagePanel As String
		StatusMessagePanel = "Type in the number of collocation points in the radial direction (PSDM only)"
		Call GenericStatus_Set(StatusMessagePanel)
	End Sub

	Private Sub _lblWaterUnit_1_Click(sender As Object, e As EventArgs) Handles _lblWaterUnit_1.Click

	End Sub

	Private Sub _lblWaterUnit_0_Click(sender As Object, e As EventArgs) Handles _lblWaterUnit_0.Click

	End Sub

	Private Sub _txtWater_0_TextChanged(sender As Object, e As EventArgs) Handles _txtWater_0.TextChanged

	End Sub

	Private Sub _txtWater_1_TextChanged(sender As Object, e As EventArgs) Handles _txtWater_1.TextChanged

	End Sub

	Private Sub _lblWater_1_Click(sender As Object, e As EventArgs) Handles _lblWater_1.Click

	End Sub

	Private Sub _lblWater_0_Click(sender As Object, e As EventArgs) Handles _lblWater_0.Click

	End Sub

	Private Sub _cmdNote_1_Click(sender As Object, e As EventArgs) Handles _cmdNote_1.Click

	End Sub

	Private Sub _lblText_0_Click(sender As Object, e As EventArgs) Handles _lblText_0.Click

	End Sub

	Private Sub _lblText_1_Click(sender As Object, e As EventArgs) Handles _lblText_1.Click

	End Sub

	Private Sub _mnuFileItem_1_Click(sender As Object, e As EventArgs) Handles _mnuFileItem_1.Click

	End Sub

	Private Sub frmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class