Option Strict Off
Option Explicit On
Friend Class frmFoulingCompoundDatabase
	Inherits System.Windows.Forms.Form
	
	'UPGRADE_WARNING: Array Local_Correlation may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
	Dim Local_Correlation(Max_Number_Correlation_Compo) As Correlation_Compound_Type



	Dim FORM_MODE As Short
	Const FORM_MODE_VIEW As Short = 1
	Const FORM_MODE_EDIT As Short = 2
	Const FORM_MODE_ADDNEW As Short = 3
	
	Dim HALT_LSTCORRELATIONS As Boolean
	
	'//////// COMMUNICATIONS WITH frmFoulingCompoundDatabase: /////////////////////////////////////////////////
	Private Structure frmFoulingCompoundDatabase_Record_Type
		Dim A1 As Double
		Dim A2 As Double
		'UPGRADE_NOTE: Name was upgraded to Name_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Name_Renamed As String
	End Structure
	Dim Local_Record As frmFoulingCompoundDatabase_Record_Type
	
	
	
	
	
	
	Const frmFoulingCompoundDatabase_declarations_end As Boolean = True
	
	
	
	Sub frmFoulingCompoundDatabase_Edit()
		Me.ShowDialog()
	End Sub
	
	
	Sub Populate_lstCorrelations()
		Dim SAVE_INDEX As Short
		Dim i As Short
		If (lstCorrelations.SelectedIndex >= 0) Then
			SAVE_INDEX = lstCorrelations.SelectedIndex
		Else
			SAVE_INDEX = 0
		End If
		HALT_LSTCORRELATIONS = True
		lstCorrelations.Items.Clear()
		For i = 1 To Number_Correlations_Compounds
			lstCorrelations.Items.Add(Local_Correlation(i).Name)
		Next i
		HALT_LSTCORRELATIONS = False
		If (SAVE_INDEX > lstCorrelations.Items.Count - 1) Then
			SAVE_INDEX = lstCorrelations.Items.Count - 1
		End If
		If (SAVE_INDEX >= 0) And (SAVE_INDEX <= lstCorrelations.Items.Count - 1) Then
			lstCorrelations.SelectedIndex = SAVE_INDEX
		End If
	End Sub
	Sub frmFoulingCompoundDatabase_Repopulate_Values()
		'Dim Frm As System.Windows.Forms.Form
		'Frm = Me
		Dim Frm As frmFoulingCompoundDatabase
		Frm = Me

		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtCoeff could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtCoeff(1), Local_Record.A1)
		'UPGRADE_ISSUE: Control txtCoeff could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtCoeff(2), Local_Record.A2)
		'TEXT DATA.
		'UPGRADE_ISSUE: Control txtName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        Frm.txtName.Text = Trim(Local_Record.Name_Renamed)
	End Sub
	Sub frmFoulingCompoundDatabase_Refresh()
		'Dim Frm As System.Windows.Forms.Form
		'Frm = Me
		Dim Frm As frmFoulingCompoundDatabase
		Frm = Me

		Dim TextLocked As Boolean
		'REPOPULATE VALUES.
		Call frmFoulingCompoundDatabase_Repopulate_Values()
		'LOCK/UNLOCK TEXTBOXES AND LISTBOX.
		TextLocked = (FORM_MODE = FORM_MODE_VIEW)
		txtCoeff(1).ReadOnly = TextLocked
		txtCoeff(2).ReadOnly = TextLocked
		txtName.ReadOnly = TextLocked
		lstCorrelations.Enabled = TextLocked
		'DISABLE/ENABLE BUTTONS.
		Select Case FORM_MODE
			Case FORM_MODE_VIEW
				'UPGRADE_ISSUE: Control lstCorrelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				If Frm.lstCorrelations.Items.Count = 0 Then
					'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					Frm._cmdRecord_0.Enabled = True 'NEW.
					'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					Frm._cmdRecord_1.Enabled = False 'EDIT.
					'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					Frm._cmdRecord_2.Enabled = False 'DELETE.
				Else
					'UPGRADE_ISSUE: Control lstCorrelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					If Frm.lstCorrelations.Items.Count >= Max_Number_Correlation_Compo Then
						'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						Frm._cmdRecord_0.Enabled = False 'NEW.
						'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						Frm._cmdRecord_1.Enabled = True 'EDIT.
						'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						Frm._cmdRecord_2.Enabled = True 'DELETE.
					Else
						'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						Frm._cmdRecord_0.Enabled = True 'NEW.
						'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						Frm._cmdRecord_1.Enabled = True 'EDIT.
						'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						Frm._cmdRecord_2.Enabled = True 'DELETE.
					End If
                End If
				'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdRecord_3.Enabled = False 'SAVE.
				'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdRecord_4.Enabled = False 'CANCEL EDIT.
				'UPGRADE_ISSUE: Control cmdCancelOK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdCancelOK_0.Enabled = True 'CANCEL.
				'UPGRADE_ISSUE: Control cmdCancelOK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdCancelOK_1.Enabled = True 'OK.
			Case FORM_MODE_EDIT, FORM_MODE_ADDNEW
				'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdRecord_0.Enabled = False 'NEW.
				'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdRecord_1.Enabled = False 'EDIT.
				'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdRecord_2.Enabled = False 'DELETE.
				'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdRecord_3.Enabled = True 'SAVE.
				'UPGRADE_ISSUE: Control cmdRecord could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdRecord_4.Enabled = True 'CANCEL EDIT.
				'UPGRADE_ISSUE: Control cmdCancelOK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdCancelOK_0.Enabled = False 'CANCEL.
				'UPGRADE_ISSUE: Control cmdCancelOK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm._cmdCancelOK_1.Enabled = False 'OK.
		End Select
	End Sub
	Sub frmFoulingCompoundDatabase_PopulateUnits()
		Call unitsys_register(Me, lblDesc(1), txtCoeff(1), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(2), txtCoeff(2), Nothing, "", "", "", "", "", 100#, False)
	End Sub
	
	
	Private Sub Load_Compound_Correlations(ByRef flag As Short)
		Dim N, f, i As Short
		On Error GoTo Error_In_Reading_Corr
		f = FreeFile
		FileOpen(f, Database_Path & "\corr_com.txt", OpenMode.Input)
		Input(f, N)
		If N > Max_Number_Correlation_Compo Then
			flag = True
			FileClose((f))
			Call Show_Error("Too many correlations in the file.")
			Exit Sub
		End If
		For i = 1 To N
			Local_Correlation(i).Initialize()
		Next
		For i = 1 To N
			Input(f, Local_Correlation(i).Name)
			Input(f, Local_Correlation(i).Coeff(1))
			Input(f, Local_Correlation(i).Coeff(2))
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
	Sub Store_Compound_Correlations()
		Dim f As Short
		Dim i As Short
		On Error GoTo Error_In_Writing_File
		f = FreeFile
		FileOpen(f, Database_Path & "\corr_com.txt", OpenMode.Output)
		WriteLine(f, Number_Correlations_Compounds)
		For i = 1 To Number_Correlations_Compounds
			WriteLine(f, Local_Correlation(i).Name, Local_Correlation(i).Coeff(1), Local_Correlation(i).Coeff(2))
		Next i
		FileClose((f))
		Exit Sub
Error_In_Writing_File: 
		Call Show_Error("Error writing to file.")
		FileClose((f))
		Resume Exit_Here
Exit_Here: 
	End Sub
	Sub Load_Local_Record(ByRef RecNum As Short)
		Local_Record.Name_Renamed = Local_Correlation(RecNum).Name
		Local_Record.A1 = Local_Correlation(RecNum).Coeff(1)
		Local_Record.A2 = Local_Correlation(RecNum).Coeff(2)
	End Sub
	Sub Store_Local_Record(ByRef RecNum As Short)
		Local_Correlation(RecNum).Name = Local_Record.Name_Renamed
		Local_Correlation(RecNum).Coeff(1) = Local_Record.A1
		Local_Correlation(RecNum).Coeff(2) = Local_Record.A2
	End Sub
	
	
	Private Sub cmdCancelOK_Click(ByRef Index As Short)
		Dim flag As Short
		Dim k, resp, f, j As Short
		Dim RetVal As Short
		Select Case Index
			Case 0 'CANCEL.
				RetVal = MsgBox("Are you sure you want to exit without " & "saving the database ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Exit Without Saving Database ?")
				If (RetVal = MsgBoxResult.No) Then Exit Sub
				Call Load_Compound_Correlations(flag)
				If flag Then Exit Sub
				Me.Close()
			Case 1 'OK.
				RetVal = MsgBox("Are you sure you want to " & "save the database ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Save Database ?")
				If (RetVal = MsgBoxResult.No) Then Exit Sub
				Call Store_Compound_Correlations()
				Me.Close()
				Exit Sub
		End Select
	End Sub
	
	
	Private Sub cmdRecord_Click(ByRef Index As Short)
		Dim RetVal As Short
		Dim New_Rec_Index As Short
		Dim Del_Rec_Index As Short
		Dim Edit_Rec_Index As Short
		Dim i As Short
		Select Case Index
			Case 0 'NEW. ///////////////////////////////////////////////////////////////////////
				If (FORM_MODE <> FORM_MODE_VIEW) Then Exit Sub
				If (lstCorrelations.Items.Count >= Max_Number_Correlation_Compo) Then
					Exit Sub
				End If
				FORM_MODE = FORM_MODE_ADDNEW
				'SET DEFAULT SETTINGS FOR NEW RECORD.
				Local_Record.Name_Renamed = "New Compound Correlation"
				Local_Record.A1 = 1#
				Local_Record.A2 = 0#
				'REFRESH WINDOW.
				Call frmFoulingCompoundDatabase_Refresh()
			Case 1 'EDIT. //////////////////////////////////////////////////////////////////////
				If (FORM_MODE <> FORM_MODE_VIEW) Then Exit Sub
				Edit_Rec_Index = lstCorrelations.SelectedIndex + 1
				If (Edit_Rec_Index < 1) Or (Edit_Rec_Index > Number_Correlations_Compounds) Then
					Call Show_Error("You must first select a correlation.")
					Exit Sub
				End If
				FORM_MODE = FORM_MODE_EDIT
				'REFRESH WINDOW.
				Call frmFoulingCompoundDatabase_Refresh()
			Case 2 'DELETE. ////////////////////////////////////////////////////////////////////
				If (FORM_MODE <> FORM_MODE_VIEW) Then Exit Sub
				Del_Rec_Index = lstCorrelations.SelectedIndex + 1
				If (Del_Rec_Index < 1) Or (Del_Rec_Index > Number_Correlations_Compounds) Then
					Call Show_Error("You must first select a correlation.")
					Exit Sub
				End If
				For i = Del_Rec_Index To Number_Correlations_Compounds - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Local_Correlation(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Local_Correlation(i) = Local_Correlation(i + 1)
				Next i
				Number_Correlations_Compounds = Number_Correlations_Compounds - 1
				'REPOPULATE LISTBOX.
				Call Populate_lstCorrelations()
				'REFRESH WINDOW.
				Call frmFoulingCompoundDatabase_Refresh()
			Case 3 'SAVE. //////////////////////////////////////////////////////////////////////
				If (FORM_MODE = FORM_MODE_VIEW) Then Exit Sub
				Select Case FORM_MODE
					Case FORM_MODE_EDIT
						Call Store_Local_Record(lstCorrelations.SelectedIndex + 1)
					Case FORM_MODE_ADDNEW
						Number_Correlations_Compounds = Number_Correlations_Compounds + 1
						New_Rec_Index = Number_Correlations_Compounds
						For i = 1 To New_Rec_Index
							Local_Correlation(i).Initialize()
						Next
						Call Store_Local_Record(New_Rec_Index)
				End Select
				FORM_MODE = FORM_MODE_VIEW
				'REPOPULATE LISTBOX.
				Call Populate_lstCorrelations()
				lstCorrelations.SelectedIndex = lstCorrelations.Items.Count - 1
				'REFRESH WINDOW.
				Call frmFoulingCompoundDatabase_Refresh()
			Case 4 'CANCEL EDIT. ///////////////////////////////////////////////////////////////
				If (FORM_MODE = FORM_MODE_VIEW) Then Exit Sub
				FORM_MODE = FORM_MODE_VIEW
				'REPOPULATE LISTBOX.
				Call Populate_lstCorrelations()
				'REFRESH WINDOW.
				Call frmFoulingCompoundDatabase_Refresh()
		End Select
	End Sub
	
	
	Private Sub frmFoulingCompoundDatabase_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i, j As Short
		Call CenterOnForm(Me, frmFouling)
		i = False

		Call Load_Compound_Correlations(i)
		If i Then Number_Correlations_Compounds = 0
		Call Populate_lstCorrelations()
		If (Number_Correlations_Compounds >= 1) Then
			Call Load_Local_Record(1)
		End If
		FORM_MODE = FORM_MODE_VIEW
		'POPULATE UNIT CONTROLS.
		Call frmFoulingCompoundDatabase_PopulateUnits()
		'REFRESH WINDOW.
		Call frmFoulingCompoundDatabase_Refresh()
	End Sub
	Private Sub frmFoulingCompoundDatabase_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub
	
	
	'UPGRADE_WARNING: Event lstCorrelations.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstCorrelations_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCorrelations.SelectedIndexChanged
		Dim ThisIndex As Short
		If (HALT_LSTCORRELATIONS) Then Exit Sub
		ThisIndex = lstCorrelations.SelectedIndex + 1
		If (ThisIndex <= lstCorrelations.Items.Count) Then
			Call Load_Local_Record(lstCorrelations.SelectedIndex + 1)
		End If
		'REFRESH WINDOW.
		Call frmFoulingCompoundDatabase_Refresh()
	End Sub
	
	
	Private Sub txtCoeff_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoeff.Enter
		Dim Index As Short = txtCoeff.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtCoeff(Index)
		Call unitsys_control_txtx_gotfocus(Ctl)
	End Sub
	Private Sub txtCoeff_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoeff.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtCoeff.GetIndex(eventSender)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtCoeff_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoeff.Leave
		Dim Index As Short = txtCoeff.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtCoeff(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
		Select Case Index
			Case 1 : Val_Low = -1E+20 : Val_High = 1E+20
			Case 2 : Val_Low = -1E+20 : Val_High = 1E+20
		End Select
		NewValue_Okay = False
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		If (NewValue_Okay) Then
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				Select Case Index
					Case 1 : Local_Record.A1 = NewValue
					Case 2 : Local_Record.A2 = NewValue
				End Select
				'RAISE DIRTY FLAG IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					''THROW DIRTY FLAG.
					'Call frmCompoProp_DirtyStatus_Throw
				End If
				'REFRESH WINDOW.
				Call frmFoulingCompoundDatabase_Refresh()
			End If
		End If
	End Sub
	
	
	Private Sub txtName_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtName.Enter
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtName
		Call Global_GotFocus(Ctl)
	End Sub
	Private Sub txtName_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtName.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_TextKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtName_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtName.Leave
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtName
		Dim OldValueStr As String
		'HANDLE STRING FIELDS.
		OldValueStr = Trim(Local_Record.Name_Renamed)
		If (Trim(Ctl.Text) = "") Then
			Ctl.Text = OldValueStr
			'Call Show_Error("You must enter a non-blank string for the carbon name.")
			'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
			'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
		Else
			If (Trim(OldValueStr) <> Trim(Ctl.Text)) Then
				Local_Record.Name_Renamed = Trim(Ctl.Text)
				''THROW DIRTY FLAG.
				'Call DirtyStatus_Throw
			End If
		End If
		Call Global_LostFocus(Ctl)
		'Call GenericStatus_Set("")
	End Sub

	Private Sub _cmdCancelOK_1_ClickEvent(sender As Object, e As EventArgs)
		Dim RetVal As Short
		RetVal = MsgBox("Are you sure you want to " & "save the database ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Save Database ?")
		If (RetVal = MsgBoxResult.No) Then Exit Sub
		Call Store_Compound_Correlations()
		Me.Dispose()
		Exit Sub
	End Sub

	Private Sub _cmdCancelOK_0_ClickEvent(sender As Object, e As EventArgs)
		Dim flag As Short
		Dim RetVal As Short
		RetVal = MsgBox("Are you sure you want to exit without " & "saving the database ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Exit Without Saving Database ?")
		If (RetVal = MsgBoxResult.No) Then Exit Sub
		Call Load_Compound_Correlations(flag)
		If flag Then Exit Sub
		Me.Dispose()
	End Sub

	Private Sub _cmdRecord_0_ClickEvent(sender As Object, e As EventArgs)
		Call cmdRecord_Click(0)
	End Sub

	Private Sub _cmdRecord_1_ClickEvent(sender As Object, e As EventArgs)
		Call cmdRecord_Click(1)
	End Sub

	Private Sub _cmdRecord_2_ClickEvent(sender As Object, e As EventArgs)
		Call cmdRecord_Click(2)
	End Sub

	Private Sub _cmdRecord_3_ClickEvent(sender As Object, e As EventArgs)
		Call cmdRecord_Click(3)
	End Sub


	Private Sub _cmdRecord_4_ClickEvent(sender As Object, e As EventArgs)
		Call cmdRecord_Click(4)
	End Sub

	Private Sub OK_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_1.Click
		Dim RetVal As Short
		RetVal = MsgBox("Are you sure you want to " & "save the database ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Save Database ?")
		If (RetVal = MsgBoxResult.No) Then Exit Sub
		Call Store_Compound_Correlations()
		Me.Dispose()
		Exit Sub
	End Sub

	Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_0.Click
		Dim flag As Short
		Dim RetVal As Short
		RetVal = MsgBox("Are you sure you want to exit without " & "saving the database ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Exit Without Saving Database ?")
		If (RetVal = MsgBoxResult.No) Then Exit Sub
		Call Load_Compound_Correlations(flag)
		If flag Then Exit Sub
		Me.Dispose()
	End Sub

	Private Sub New_Click(sender As Object, e As EventArgs) Handles _cmdRecord_0.Click
		Call cmdRecord_Click(0)
	End Sub

	Private Sub Edit_Click(sender As Object, e As EventArgs) Handles _cmdRecord_1.Click
		Call cmdRecord_Click(1)
	End Sub

	Private Sub Delete_Click(sender As Object, e As EventArgs) Handles _cmdRecord_2.Click
		Call cmdRecord_Click(2)
	End Sub

	Private Sub Save_Click(sender As Object, e As EventArgs) Handles _cmdRecord_3.Click
		Call cmdRecord_Click(3)
	End Sub

	Private Sub CancelEdit_Click(sender As Object, e As EventArgs) Handles _cmdRecord_4.Click
		Call cmdRecord_Click(4)
	End Sub
End Class