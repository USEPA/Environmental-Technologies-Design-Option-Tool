Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmFouling
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Dim Raise_Dirty_Flag As Boolean
	
	
	
	
	
	
	Const frmFouling_decl_end As Boolean = True
	
	
	Sub frmFouling_Go(ByRef OUT_Raise_Dirty_Flag As Boolean)
		Raise_Dirty_Flag = False
		Me.ShowDialog()
		OUT_Raise_Dirty_Flag = Raise_Dirty_Flag
	End Sub


	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdCancelOK_1.Enabled = False
		End If
	End Sub


	Sub Populate_cboType()
		Dim i As Short
		i = False
		Call Load_Correlations_Water(i)
		If i Then Number_Water_Correlations = 0
		cboType.Items.Clear()
		For i = 1 To Number_Water_Correlations
			cboType.Items.Add(Trim(Correlations_For_Water(i).Name))
		Next i
		'If (DemoMode) Then
		'  cboType.ListIndex = 0
		'Else
		If Number_Water_Correlations > 0 Then cboType.SelectedIndex = Set_Number_Correlation_Water() - 1
		'End If
	End Sub
	Sub Populate_cboCorrel()
		Dim i As Short
		Dim J As Short
		i = False
		Call Load_Correlation_Compounds(i)
		If i Then Number_Correlations_Compounds = 0
		For i = 0 To Number_Component - 1
			cboCorrel(i).Visible = True
			cboCorrel(i).Items.Clear()
			For J = 1 To Number_Correlations_Compounds
				cboCorrel(i).Items.Add(Trim(Correlations_For_Classes(J).Name))
			Next J
			'If (DemoMode) Then
			'  ' Set it to Halogenated Alkenes
			'  If (0 = StrComp(lblName(i).Caption, "trichloroethylene", 1)) Then cboCorrel(i).ListIndex = 1
			'  ' Set it to Aromatics
			'  If (0 = StrComp(lblName(i).Caption, "benzene", 1)) Then cboCorrel(i).ListIndex = 3
			'  ' Set it to Aromatics
			'  If (0 = StrComp(lblName(i).Caption, "1,2-dichlorobenzene", 1)) Then cboCorrel(i).ListIndex = 3
			'Else
			cboCorrel(i).SelectedIndex = (Set_Number_Correlation(i + 1) - 1)
			'End If
			'UPGRADE_WARNING: Couldn't resolve default property of object chkUse().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkUse(i).Visible = True

			'UPGRADE_WARNING: Couldn't resolve default property of object chkUse(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object chkUse(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

			'K_reduction not properly defined, returns false, commented out to prevent disabling
			'change to .value instead of .enable?
			chkUse(i).Checked = Component(i + 1).K_Reduction
			lblName(i).Visible = True
			lblName(i).Text = Trim(Component(i + 1).Name)
		Next i
	End Sub


	'UPGRADE_WARNING: Event cboCorrel.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCorrel_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCorrel.SelectedIndexChanged
		Dim Index As Short = cboCorrel.GetIndex(eventSender)
		'If (DemoMode) Then
		'    ' Set it to Halogenated Alkenes
		'    If (0 = StrComp(Trim$(lblName(index).Caption), "trichloroethylene", 1)) Then cboCorrel(index).ListIndex = 1
		'    ' Set it to Aromatics
		'    If (0 = StrComp(Trim$(lblName(index).Caption), "benzene", 1)) Then cboCorrel(index).ListIndex = 3
		'    ' Set it to Aromatics
		'    If (0 = StrComp(Trim$(lblName(index).Caption), "1,2-dichlorobenzene", 1)) Then cboCorrel(index).ListIndex = 3
		'End If
	End Sub


	'UPGRADE_WARNING: Event cboType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.SelectedIndexChanged
		Dim msg As String
		Static old_index As Short
		'  If (DemoMode) Then
		'    If (0 = StrComp(Trim$(cboType.Text), "Organic Free Water", 1)) Then
		'      old_index% = 0
		'      Exit Sub
		'    End If
		'    If (0 = StrComp(Trim$(cboType.Text), "Groundwater from the city of Karlsruhe, Germany", 1)) Then
		'      old_index% = 3
		'      Exit Sub
		'    End If
		'    msg$ = "            " + cboType.Text + NL + NL
		'    msg$ = msg$ + "Is not a valid Water Type in the Demonstration version."
		'    MsgBox msg$
		'    cboType.ListIndex = old_index%
		'  End If
	End Sub


	Private Sub chkUse_Click(ByRef Index As Short)
		Dim Is_Invalid As Boolean
		'DE-APPLY CORRELATION IF USER HAS NOT PROPERLY
		'SELECTED A CORRELATION.
		'UPGRADE_WARNING: Couldn't resolve default property of object chkUse(Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (CBool(chkUse(Index).Enabled) = True) Then
			Is_Invalid = False
			If (cboCorrel(Index).SelectedIndex < 0) Then
				Is_Invalid = True
			Else
				If (VB6.GetItemString(cboCorrel(Index), cboCorrel(Index).SelectedIndex) = "") Then
					Is_Invalid = True
				End If
			End If
			If (Is_Invalid) Then

				If chkUse(Index).Checked = True Then
					Call Show_Error("You must select a correlation " & "type before you can apply fouling for this chemical.")
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object chkUse(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chkUse(Index).Checked = False
				Exit Sub
			End If
		End If
	End Sub


	Private Sub cmdCancelOK_Click(ByRef Index As Short)
		'This code is not used

		Dim i As Short
		Dim msg As String
		Dim IsInvalid As Boolean
		Select Case Index
			Case 0 'CANCEL.
				Raise_Dirty_Flag = False
				Me.Close()
			Case 1 'OK.
				IsInvalid = True
				If (cboType.SelectedIndex >= 0) Then
					If (cboType.Items.Count >= 1) Then
						If (Trim(VB6.GetItemString(cboType, cboType.SelectedIndex)) <> "") Then
							IsInvalid = False
						End If
					End If
				End If
				If (IsInvalid) Then
					Call Show_Error("You must first select a water correlation type.")
					Exit Sub
				End If
				'      If (DemoMode) Then
				'        If (0 = StrComp(Trim$(cboType.Text), "Organic Free Water")) Then GoTo DEMO_00_CONTINUE
				'        If (0 = StrComp(Trim$(cboType.Text), "Groundwater from the city of Karlsruhe, Germany", 1)) Then GoTo DEMO_00_CONTINUE
				'        msg$ = "In Demonstration version you can only use two types of water:" + NL + NL
				'        msg$ = msg$ + Chr$(9) + "- Organic Free Water" + NL
				'        msg$ = msg$ + Chr$(9) + "- Groundwater from the city of Karlsruhe, Germany" + NL
				'        MsgBox msg$
				'        Exit Sub
				'      End If
				'DEMO_00_CONTINUE:
				For i = 1 To Number_Component
					If cboCorrel(i - 1).SelectedIndex > -1 Then
						Component(i).Correlation.Name = Trim(VB6.GetItemString(cboCorrel(i - 1), cboCorrel(i - 1).SelectedIndex))
						'UPGRADE_WARNING: Couldn't resolve default property of object chkUse(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Component(i).K_Reduction = chkUse(i - 1).Enabled
						Component(i).Correlation.Coeff(1) = Correlations_For_Classes(cboCorrel(i - 1).SelectedIndex + 1).Coeff(1)
						Component(i).Correlation.Coeff(2) = Correlations_For_Classes(cboCorrel(i - 1).SelectedIndex + 1).Coeff(2)
					Else
						Component(i).K_Reduction = False
					End If
				Next i
				If cboType.SelectedIndex = -1 Then cboType.SelectedIndex = 0
				Bed.Water_Correlation.Name = Correlations_For_Water(cboType.SelectedIndex + 1).Name
				For i = 1 To 4
					Bed.Water_Correlation.Coeff(i) = Correlations_For_Water(cboType.SelectedIndex + 1).Coeff(i)
				Next i
				'
				' STORE SIGNAL TO RAISE DIRTY FLAG AND THEN EXIT.
				Raise_Dirty_Flag = True
				Me.Close()
		End Select
	End Sub


	Private Sub cmdEdit_Click()
		Call frmFoulingWaterDatabase.frmFoulingWaterDatabase_Edit()
		Call Populate_cboType()
	End Sub
	Private Sub cmdEditCompo_Click()
		Call frmFoulingCompoundDatabase.frmFoulingCompoundDatabase_Edit()
		Call Populate_cboCorrel()
	End Sub


	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub

	Private Sub frmFouling_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Short
		Dim J As Short

		rs.FindAllControls(Me)

		'If (DemoMode) Then
		'  cmdEdit.Enabled = False
		'  cmdEditCompo.Enabled = False
		'End If
		'Me.HelpContextID = Hlp_Fouling_of


		Call Populate_cboType()

		Call Populate_cboCorrel()



		For i = Number_Component To Number_Compo_Max - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object chkUse().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkUse(i).Visible = False
			lblName(i).Visible = False
			cboCorrel(i).Visible = False
		Next i


		'UPGRADE_WARNING: Couldn't resolve default property of object cmdEditCompo.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		cmdEditCompo.Top = VB6.PixelsToTwipsY(lblName(Number_Component - 1).Top) + VB6.PixelsToTwipsY(lblName(Number_Component - 1).Height) + VB6.TwipsPerPixelY * 10
		'UPGRADE_WARNING: Couldn't resolve default property of object fraCompo.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdEditCompo.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdEditCompo.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		fraCompo.Height = cmdEditCompo.Top + cmdEditCompo.Height + VB6.TwipsPerPixelY * 10
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK().Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fraCompo.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fraCompo.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		_cmdCancelOK_1.Top = fraCompo.Top + fraCompo.Height + VB6.TwipsPerPixelY * 10
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK().Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fraCompo.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fraCompo.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		_cmdCancelOK_0.Top = fraCompo.Top + fraCompo.Height + VB6.TwipsPerPixelY * 10
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK(1).Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK().Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		Height = VB6.TwipsToPixelsY(_cmdCancelOK_1.Top + _cmdCancelOK_1.Height + VB6.TwipsPerPixelY * 35)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdEditCompo.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdEditCompo.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fraCompo.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		cmdEditCompo.Left = (fraCompo.Width - cmdEditCompo.Width) / 2
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdEdit.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdEdit.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fraWater.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		cmdEdit.Left = (fraWater.Width - cmdEdit.Width) / 2
		'UPGRADE_WARNING: Couldn't resolve default property of object fraWater.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		cboType.Left = VB6.TwipsToPixelsX((fraWater.Width - VB6.PixelsToTwipsX(cboType.Width)) / 2)
		'		Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) / 2 - VB6.PixelsToTwipsY(Height) / 2)
		'		Me.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 2 - VB6.PixelsToTwipsX(Width) / 2)
		Call CenterOnForm(Me, frmMain)
		'
		' DEMO SETTINGS.
		'
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub


	Private Sub Load_Correlation_Compounds(ByRef flag As Short)
		Dim N, f, i As Short
		On Error GoTo Error_In_Reading_Corr
		f = FreeFile()
		FileOpen(f, Database_Path & "\corr_com.txt", OpenMode.Input)
		Input(f, N)
		If N > Max_Number_Correlation_Compo Then
			flag = True
			FileClose((f))
			Call Show_Error("Too many correlations in the file.")
			Exit Sub
		End If

		For i = 1 To N
			Correlations_For_Classes(i).Initialize()
		Next
		For i = 1 To N
			Input(f, Correlations_For_Classes(i).Name)
			Input(f, Correlations_For_Classes(i).Coeff(1))
			Input(f, Correlations_For_Classes(i).Coeff(2))
		Next i
		FileClose((f))
		Number_Correlations_Compounds = N
		flag = False
		Exit Sub
Error_In_Reading_Corr:
		Call Show_Error("Error while reading the file containing correlations.")
		flag = True
		Resume Exit_Corr_Compound
Exit_Corr_Compound:
	End Sub
	Private Sub Load_Correlations_Water(ByRef flag As Short)
		Dim N, f, i As Short
		On Error GoTo Error_In_Reading_WCorr
		f = FreeFile()
		FileOpen(f, Database_Path & "\water_co.txt", OpenMode.Input)
		Input(f, N)
		If N > Max_Number_Water_Correlations Then
			flag = True
			FileClose((f))
			Call Show_Error("Too many correlations in the file.")
			Exit Sub
		End If

		For i = 1 To N
			Correlations_For_Water(i).Initialize()   'Shang add
		Next i

		For i = 1 To N
			Input(f, Correlations_For_Water(i).Name)
			Input(f, Correlations_For_Water(i).Coeff(1))
			Input(f, Correlations_For_Water(i).Coeff(2))
			Input(f, Correlations_For_Water(i).Coeff(3))
			Input(f, Correlations_For_Water(i).Coeff(4))
		Next i
		FileClose((f))
		Number_Water_Correlations = N
		flag = False
		Exit Sub
Error_In_Reading_WCorr:
		Call Show_Error("Error while reading the file containing correlations.")
		flag = True
		FileClose((f))
		Resume Exit_Corr_Water
Exit_Corr_Water:
	End Sub
	Private Function Set_Number_Correlation(ByRef i As Short) As Short
		Dim ST As String
		Dim J As Short
		ST = Component(i).Correlation.Name
		For J = 1 To Number_Correlations_Compounds
			If Trim(ST) = Trim(Correlations_For_Classes(J).Name) Then
				Set_Number_Correlation = J
				Exit Function
			Else
				Set_Number_Correlation = 0
			End If
		Next J
	End Function
	Private Function Set_Number_Correlation_Water() As Short
		Dim ST As String
		Dim J As Short
		ST = Bed.Water_Correlation.Name
		For J = 1 To Number_Water_Correlations
			If Trim(ST) = Trim(Correlations_For_Water(J).Name) Then
				Set_Number_Correlation_Water = J
				Exit Function
			Else
				Set_Number_Correlation_Water = 0
			End If
		Next J
	End Function

	Private Sub cmdEdit_ClickEvent(sender As Object, e As EventArgs)
		Call frmFoulingWaterDatabase.frmFoulingWaterDatabase_Edit()
		Call Populate_cboType()
	End Sub

	Private Sub cmdEditCompo_ClickEvent(sender As Object, e As EventArgs)
		Call frmFoulingCompoundDatabase.frmFoulingCompoundDatabase_Edit()
		Call Populate_cboCorrel()
	End Sub

	Private Sub _cmdCancelOK_1_ClickEvent(sender As Object, e As EventArgs)
		Dim i As Short
		'	Dim msg As String
		Dim IsInvalid As Boolean
		IsInvalid = True
		If (cboType.SelectedIndex >= 0) Then
			If (cboType.Items.Count >= 1) Then
				If (Trim(VB6.GetItemString(cboType, cboType.SelectedIndex)) <> "") Then
					IsInvalid = False
				End If
			End If
		End If
		If (IsInvalid) Then
			Call Show_Error("You must first select a water correlation type.")
			Exit Sub
		End If
		'      If (DemoMode) Then
		'        If (0 = StrComp(Trim$(cboType.Text), "Organic Free Water")) Then GoTo DEMO_00_CONTINUE
		'        If (0 = StrComp(Trim$(cboType.Text), "Groundwater from the city of Karlsruhe, Germany", 1)) Then GoTo DEMO_00_CONTINUE
		'        msg$ = "In Demonstration version you can only use two types of water:" + NL + NL
		'        msg$ = msg$ + Chr$(9) + "- Organic Free Water" + NL
		'        msg$ = msg$ + Chr$(9) + "- Groundwater from the city of Karlsruhe, Germany" + NL
		'        MsgBox msg$
		'        Exit Sub
		'      End If
		'DEMO_00_CONTINUE:
		For i = 1 To Number_Component
			If cboCorrel(i - 1).SelectedIndex > -1 Then
				Component(i).Correlation.Name = Trim(VB6.GetItemString(cboCorrel(i - 1), cboCorrel(i - 1).SelectedIndex))
				'UPGRADE_WARNING: Couldn't resolve default property of object chkUse(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Component(i).K_Reduction = chkUse(i - 1).Checked
				'^changed from .enabled to .value
				Component(i).Correlation.Coeff(1) = Correlations_For_Classes(cboCorrel(i - 1).SelectedIndex + 1).Coeff(1)
				Component(i).Correlation.Coeff(2) = Correlations_For_Classes(cboCorrel(i - 1).SelectedIndex + 1).Coeff(2)
			Else
				Component(i).K_Reduction = False
			End If
		Next i
		If cboType.SelectedIndex = -1 Then cboType.SelectedIndex = 0
		Bed.Water_Correlation.Name = Correlations_For_Water(cboType.SelectedIndex + 1).Name
		For i = 1 To 4
			Bed.Water_Correlation.Coeff(i) = Correlations_For_Water(cboType.SelectedIndex + 1).Coeff(i)
		Next i
		'
		' STORE SIGNAL TO RAISE DIRTY FLAG AND THEN EXIT.
		Raise_Dirty_Flag = True
		Me.Dispose()


	End Sub

	Private Sub _cmdCancelOK_0_ClickEvent(sender As Object, e As EventArgs)
		Raise_Dirty_Flag = False
		Me.Dispose()  'Dispose Shang
	End Sub

	Private Sub cmdEdit_Enter(sender As Object, e As EventArgs)
		'Call frmFoulingWaterDatabase.frmFoulingWaterDatabase_Edit()
		Call Populate_cboType()
	End Sub

	Private Sub cmdEditCompo_Enter(sender As Object, e As EventArgs)
		'Call frmFoulingCompoundDatabase.frmFoulingCompoundDatabase_Edit()
		Call Populate_cboCorrel()
	End Sub

	Private Sub _chkUse_0_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(0)
	End Sub

	Private Sub _chkUse_1_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(1)
	End Sub

	Private Sub _chkUse_2_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(2)
	End Sub

	Private Sub _chkUse_3_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(3)
	End Sub

	Private Sub _chkUse_4_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(4)
	End Sub

	Private Sub _chkUse_5_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(5)
	End Sub

	Private Sub _chkUse_6_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(6)
	End Sub

	Private Sub _chkUse_7_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(7)
	End Sub

	Private Sub _chkUse_8_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(8)
	End Sub

	Private Sub _chkUse_9_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		Call chkUse_Click(9)
	End Sub

	Private Sub _cmdCancelOK_1_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_1.Click
		'ok

		Dim i As Short
		'	Dim msg As String
		Dim IsInvalid As Boolean
		IsInvalid = True
		If (cboType.SelectedIndex >= 0) Then
			If (cboType.Items.Count >= 1) Then
				If (Trim(VB6.GetItemString(cboType, cboType.SelectedIndex)) <> "") Then
					IsInvalid = False
				End If
			End If
		End If
		If (IsInvalid) Then
			Call Show_Error("You must first select a water correlation type.")
			Exit Sub
		End If
		'      If (DemoMode) Then
		'        If (0 = StrComp(Trim$(cboType.Text), "Organic Free Water")) Then GoTo DEMO_00_CONTINUE
		'        If (0 = StrComp(Trim$(cboType.Text), "Groundwater from the city of Karlsruhe, Germany", 1)) Then GoTo DEMO_00_CONTINUE
		'        msg$ = "In Demonstration version you can only use two types of water:" + NL + NL
		'        msg$ = msg$ + Chr$(9) + "- Organic Free Water" + NL
		'        msg$ = msg$ + Chr$(9) + "- Groundwater from the city of Karlsruhe, Germany" + NL
		'        MsgBox msg$
		'        Exit Sub
		'      End If
		'DEMO_00_CONTINUE:
		For i = 1 To Number_Component
			If cboCorrel(i - 1).SelectedIndex > -1 Then
				Component(i).Correlation.Name = Trim(VB6.GetItemString(cboCorrel(i - 1), cboCorrel(i - 1).SelectedIndex))
				'UPGRADE_WARNING: Couldn't resolve default property of object chkUse(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Component(i).K_Reduction = chkUse(i - 1).Checked
				'^changed from .enabled to .value
				Component(i).Correlation.Coeff(1) = Correlations_For_Classes(cboCorrel(i - 1).SelectedIndex + 1).Coeff(1)
				Component(i).Correlation.Coeff(2) = Correlations_For_Classes(cboCorrel(i - 1).SelectedIndex + 1).Coeff(2)
			Else
				Component(i).K_Reduction = False
			End If
		Next i
		If cboType.SelectedIndex = -1 Then cboType.SelectedIndex = 0
		Bed.Water_Correlation.Name = Correlations_For_Water(cboType.SelectedIndex + 1).Name
		For i = 1 To 4
			Bed.Water_Correlation.Coeff(i) = Correlations_For_Water(cboType.SelectedIndex + 1).Coeff(i)
		Next i
		'
		' STORE SIGNAL TO RAISE DIRTY FLAG AND THEN EXIT.
		Raise_Dirty_Flag = False
		Me.Dispose()  'Dispose Shang

	End Sub

	Private Sub _cmdCancelOK_0_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_0.Click
		'cancel
		Raise_Dirty_Flag = False
		Me.Dispose()  'Dispose Shang

	End Sub

	Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_0.CheckedChanged
		Call chkUse_Click(0)
	End Sub

	Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_1.CheckedChanged
		Call chkUse_Click(1)
	End Sub

	Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_2.CheckedChanged
		Call chkUse_Click(2)
	End Sub

	Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_3.CheckedChanged
		Call chkUse_Click(3)
	End Sub

	Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_4.CheckedChanged
		Call chkUse_Click(4)
	End Sub

	Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_5.CheckedChanged
		Call chkUse_Click(5)
	End Sub

	Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_6.CheckedChanged
		Call chkUse_Click(6)
	End Sub

	Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_7.CheckedChanged
		Call chkUse_Click(7)
	End Sub

	Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_8.CheckedChanged
		Call chkUse_Click(8)
	End Sub

	Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles _chkUse_9.CheckedChanged
		Call chkUse_Click(9)
	End Sub



	Private Sub ECTC_Click(sender As Object, e As EventArgs) Handles cmdEditCompo.Click
		Call frmFoulingCompoundDatabase.frmFoulingCompoundDatabase_Edit()
		Call Populate_cboCorrel()
	End Sub

	Private Sub Edit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
		Call frmFoulingWaterDatabase.frmFoulingWaterDatabase_Edit()
		Call Populate_cboType()
	End Sub

	Private Sub Picture1_Click(sender As Object, e As EventArgs) Handles Picture1.Click

	End Sub

	Private Sub frmFouling_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class