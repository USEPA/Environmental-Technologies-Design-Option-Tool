Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmFileNote
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer


	Dim NoteText As String
	Dim Rollback_NoteText As String
	Dim RaiseDirtyFlag As Boolean
	Dim Rollback_RaiseDirtyFlag As Boolean
	
	Dim FORM_MODE As Short
	Const FORM_MODE_VIEW As Short = 1
	Const FORM_MODE_EDIT As Short = 2
	
	
	
	
	Const frmFileNote_declarations_end As Boolean = True
	
	
	Sub frmFileNote_Run(ByRef IO_NoteText As String, ByRef O_RaiseDirtyFlag As Boolean)
		NoteText = IO_NoteText
		''''NoteText = Parser_ReplaceStrings(NoteText, Chr$(255), Chr$(13) & Chr$(10))
		Me.ShowDialog()
		If (RaiseDirtyFlag) Then
			IO_NoteText = NoteText
			''''MsgBox NoteText
			''''IO_NoteText = Parser_ReplaceStrings(IO_NoteText, Chr$(13) & Chr$(10), Chr$(255))
		End If
		O_RaiseDirtyFlag = RaiseDirtyFlag
	End Sub
	
	
	Sub frmFileNote_Refresh()
		Select Case FORM_MODE
			Case FORM_MODE_VIEW
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_0.Enabled = True
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_1.Enabled = True
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_2.Enabled = True
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_3.Enabled = False
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_4.Enabled = False
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_5.Enabled = False
				txtData.ReadOnly = True
				lblInstructions.Text = ""
			Case FORM_MODE_EDIT
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_0.Enabled = False
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_1.Enabled = False
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_2.Enabled = False
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_3.Enabled = True
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_4.Enabled = True
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdButton().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				_cmdButton_5.Enabled = True
				txtData.ReadOnly = False
				lblInstructions.Text = "You may enter up to 500 characters of text.  " & "Line breaks are acceptable."
		End Select
		txtData.Text = NoteText
	End Sub


	Private Sub cmdButton_Click(ByRef Index As Short)
		'Private Sub cmdbutton_click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
		Dim RetVal As Short
		'Dim Index As Short
		'Index = Array.IndexOf(cmdButton, eventSender)
		Select Case Index
			Case 0 'DELETE.
				RetVal = MsgBox("Are you sure you want to delete this " & "file note ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Delete File Note ?")
				If (RetVal = MsgBoxResult.No) Then Exit Sub
				NoteText = ""
				RaiseDirtyFlag = True
				Call frmFileNote_Refresh()
			Case 1 'EDIT.
				'STORE OLD VALUE IF USER CANCELS EDIT.
				Rollback_NoteText = NoteText
				Rollback_RaiseDirtyFlag = RaiseDirtyFlag
				FORM_MODE = FORM_MODE_EDIT
				Call frmFileNote_Refresh()
			Case 2 'CLOSE.
				Me.Close()
				Exit Sub
			Case 3 'SAVE.
				FORM_MODE = FORM_MODE_VIEW
				Call frmFileNote_Refresh()
				''temp
				'Open "c:\test.out" For Output As #1
				'Dim i As Integer
				'For i = 1 To Len(NoteText)
				'  Print #1, Asc(Mid$(NoteText, i, 1))
				'Next i
				'Close #1
			Case 4 'CANCEL EDIT.
				FORM_MODE = FORM_MODE_VIEW
				NoteText = Rollback_NoteText
				RaiseDirtyFlag = Rollback_RaiseDirtyFlag
				Call frmFileNote_Refresh()
			Case 5 'INSERT DATE/TIME.
				txtData.Text = txtData.Text & Now
				Call txtData_Leave(txtData, New System.EventArgs())
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

	Private Sub frmFileNote_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'MISC INITS.
		rs.FindAllControls(Me)

		Call CenterOnForm(Me, frmMain)
		FORM_MODE = FORM_MODE_VIEW
		Call frmFileNote_Refresh()
		RaiseDirtyFlag = False
	End Sub


	Private Sub txtData_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Enter
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtData
		Call Global_GotFocus(Ctl)
		'FORCE BACKGROUND COLOR BACK TO WHITE.
		Ctl.BackColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 255, 255))
	End Sub
	Private Sub txtData_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtData.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_MultilineTextKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtData_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Leave
		'Dim NewValue_Okay As Integer
		'Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtData
		'Dim Val_Low As Double
		'Dim Val_High As Double
		'Dim Raise_Dirty_Flag As Boolean
		'Dim Too_Small As Integer
		Dim OldValueStr As String
		'HANDLE STRING FIELDS.
		OldValueStr = Trim(NoteText)
		'NOTE: ZERO-LENGTH STRINGS ARE ALLOWED.
		If (Trim(OldValueStr) <> Trim(Ctl.Text)) Then
			NoteText = Trim(Ctl.Text)
			RaiseDirtyFlag = True
		End If
		Call Global_LostFocus(Ctl)
		Call frmFileNote_Refresh()
		'Call GenericStatus_Set("")
		Exit Sub
		'  End If
		'  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
		'  Select Case index
		'    Case 0: Val_Low = 1E-20: Val_High = 1E+20
		'    Case 1: Val_Low = 0#: Val_High = 1E+20
		'    Case 2: Val_Low = 0#: Val_High = 1E+20
		'    Case 3: Val_Low = 1E-20: Val_High = 1E+20
		'    Case 4: Val_Low = 1E-20: Val_High = 1E+20
		'    Case 5: Val_Low = 0#: Val_High = 1E+20
		'    Case 6: Val_Low = 0#: Val_High = 1E+20
		'  End Select
		'  NewValue_Okay = False
		'  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
		'    NewValue_Okay = True
		'  End If
		'  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		'  If (NewValue_Okay) Then
		'    If (Raise_Dirty_Flag) Then
		'      'STORE TO MEMORY.
		'      Select Case index
		'        Case 0:         'FREUNDLICH K.
		'          frmEditIsothermData_Record.k = NewValue
		'        Case 1:         'MINIMUM CONCENTRATION.
		'          frmEditIsothermData_Record.Cmin = NewValue
		'        Case 2:         'MINIMUM pH.
		'          frmEditIsothermData_Record.pHmin = NewValue
		'        Case 3:         'TEMPERATURE.
		'          frmEditIsothermData_Record.Tmin = NewValue
		'        Case 4:         'FREUNDLICH 1/n.
		'          frmEditIsothermData_Record.OneOverN = NewValue
		'        Case 5:         'MAXIMUM CONCENTRATION.
		'          frmEditIsothermData_Record.Cmax = NewValue
		'        Case 6:         'MAXIMUM pH.
		'          frmEditIsothermData_Record.pHmax = NewValue
		'      End Select
		'      'RAISE DIRTY FLAG IF NECESSARY.
		'      If (Raise_Dirty_Flag) Then
		'        ''THROW DIRTY FLAG.
		'        'Call frmCompoProp_DirtyStatus_Throw
		'      End If
		'      'REFRESH WINDOW.
		'      Call frmEditIsothermData_Refresh
		'    End If
		'  End If
	End Sub

	Private Sub _cmdButton_0_Click(sender As Object, e As EventArgs) Handles _cmdButton_0.Click
		Call cmdButton_Click(0)
	End Sub

	Private Sub _cmdButton_1_Click(sender As Object, e As EventArgs) Handles _cmdButton_1.Click
		Call cmdButton_Click(1)
	End Sub

	Private Sub _cmdButton_2_Click(sender As Object, e As EventArgs) Handles _cmdButton_2.Click
		Call cmdButton_Click(2)
	End Sub

	Private Sub _cmdButton_5_Click(sender As Object, e As EventArgs) Handles _cmdButton_5.Click
		Call cmdbutton_click(5)
	End Sub

	Private Sub _cmdButton_3_Click(sender As Object, e As EventArgs) Handles _cmdButton_3.Click
		Call cmdButton_Click(3)
	End Sub

	Private Sub _cmdButton_4_Click(sender As Object, e As EventArgs) Handles _cmdButton_4.Click
		Call cmdButton_Click(4)
	End Sub

	Private Sub frmFileNote_RegionChanged(sender As Object, e As EventArgs) Handles Me.RegionChanged

	End Sub

	Private Sub frmFileNote_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub

	'	Private Sub _cmdButton_1_ClickEvent(sender As Object, e As EventArgs) Handles _cmdButton_1.ClickEvent
	'	Call cmdButton_Click(1)
	'	End Sub

	'Private Sub _cmdButton_0_ClickEvent(sender As Object, e As EventArgs) Handles _cmdButton_0.ClickEvent
	'Call cmdButton_Click(0)
	'End Sub


	'Private Sub _cmdButton_3_ClickEvent(sender As Object, e As EventArgs) Handles _cmdButton_3.ClickEvent
	'Call cmdButton_Click(3)
	'End Sub

	'Private Sub _cmdButton_4_ClickEvent(sender As Object, e As EventArgs) Handles _cmdButton_4.ClickEvent
	'Call cmdButton_Click(4)
	'End Sub

	'Private Sub _cmdButton_5_ClickEvent(sender As Object, e As EventArgs) Handles _cmdButton_5.ClickEvent
	'Call cmdButton_Click(5)
	'End Sub

	'	Private Sub _cmdButton_2_ClickEvent(sender As Object, e As EventArgs) Handles _cmdButton_2.ClickEvent
	'	Call cmdButton_Click(2)
	'	End Sub
End Class