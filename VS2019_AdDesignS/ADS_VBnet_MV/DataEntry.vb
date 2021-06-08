Option Strict Off
Option Explicit On
Module DataEntry
	
	
	
	Function Format_It(ByVal X As Double, ByVal N As Short) As String
		Dim s As String
		Select Case N
			Case 2
				Select Case System.Math.Abs(X)
					Case Is < 0.1
						s = VB6.Format(X, "0.00E+00")
					Case Is > 100#
						s = VB6.Format(X, "0.00E+00")
					Case Else
						s = VB6.Format(X, "0.00")
				End Select
			Case 3
				Select Case System.Math.Abs(X)
					Case Is < 0.1
						s = VB6.Format(X, "0.000E+00")
					Case Is > 100#
						s = VB6.Format(X, "0.000E+00")
					Case Else
						s = VB6.Format(X, "0.000")
				End Select
		End Select
		Format_It = s
	End Function
	
	
	Public Function Global_ReadOnlyKeyPress(ByVal KeyAscii As Short) As Short
		Dim KeyStroke As Short
		Select Case KeyAscii
			' CONTROL CHARACTERS: ^C, <BS>, ^V, ^X, ^Z
			Case 3, 8, 22, 24, 26
				KeyStroke = KeyAscii
			Case 13 ' <Enter> -> <Tab>
				System.Windows.Forms.SendKeys.SendWait("{TAB}")
			Case Else
				KeyStroke = 0
		End Select
		Global_ReadOnlyKeyPress = KeyStroke
	End Function
	Public Function Global_Numeric0123456789KeyPress(ByVal KeyAscii As Short) As Short
		Dim KeyStroke As Short
		Select Case KeyAscii
			Case 48 To 57
				KeyStroke = KeyAscii
				' CONTROL CHARACTERS: ^C, <BS>, ^V, ^X, ^Z
			Case 3, 8, 22, 24, 26
				KeyStroke = KeyAscii
			Case 13 ' <Enter> -> <Tab>
				System.Windows.Forms.SendKeys.SendWait("{TAB}")
			Case Else
				KeyStroke = 0
		End Select
		Global_Numeric0123456789KeyPress = KeyStroke
	End Function
	Public Function Global_NumericKeyPress(ByVal KeyAscii As Short) As Short
		Dim KeyStroke As Short
		'
		' THIS FUNCTION ONLY ALLOWS THE USER TO
		' ENTER NUMERIC VALUES INTO A TEXT BOX
		'
		Select Case KeyAscii
			' ASCII CHARACTERS:  +, -, ., 0-9, E
			Case 43, 45, 46, 48 To 57, 69
				KeyStroke = KeyAscii
				' CONTROL CHARACTERS: ^C, <BS>, ^V, ^X, ^Z
			Case 3, 8, 22, 24, 26
				KeyStroke = KeyAscii
			Case 101 ' e -> E
				KeyStroke = 69
			Case 13 ' <Enter> -> <Tab>
				System.Windows.Forms.SendKeys.SendWait("{TAB}")
			Case Else
				KeyStroke = 0
		End Select
		Global_NumericKeyPress = KeyStroke
	End Function
	Public Function Global_MultilineTextKeyPress(ByVal KeyAscii As Short) As Short
		Dim KeyStroke As Short
		Select Case KeyAscii
			'Case 13: ' <Enter> -> <Tab>
			'  SendKeys "{TAB}", True
			'  KeyStroke% = 0
			Case 34 'double-quote character (")
				KeyStroke = 0
			Case Else
				KeyStroke = KeyAscii
		End Select
		Global_MultilineTextKeyPress = KeyStroke
	End Function
	Public Function Global_TextKeyPress(ByVal KeyAscii As Short) As Short
		Dim KeyStroke As Short
		Select Case KeyAscii
			Case 13 ' <Enter> -> <Tab>
				System.Windows.Forms.SendKeys.SendWait("{TAB}")
				KeyStroke = 0
			Case 34 'double-quote character (")
				KeyStroke = 0
			Case Else
				KeyStroke = KeyAscii
		End Select
		Global_TextKeyPress = KeyStroke
	End Function
	Public Sub Global_GotFocus(ByRef gObject As TextBoxBase)
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.BackColor = Color.FromArgb(0, 220, 220)
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.SelectionStart = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.SelectionLength = Len(gObject.Text)
	End Sub
	Public Sub Global_LostFocus(ByRef gObject As TextBox)
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.BackColor = Color.FromArgb(255, 255, 255)
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.SelectionLength = 0
	End Sub


	Function GetDoubleFormat(ByVal Value As Double) As String
		Dim AbsValue As Double
		AbsValue = System.Math.Abs(Value)
		Select Case AbsValue
			Case 0#
				GetDoubleFormat = "0"
			Case Is < 0.001
				GetDoubleFormat = "0.00E+00"
			Case Is < 0.01
				GetDoubleFormat = "0.00E+00"
			Case Is < 0.1
				GetDoubleFormat = "0.0000"
			Case Is < 1
				GetDoubleFormat = "0.000"
			Case Is < 10
				GetDoubleFormat = "0.00"
			Case Is < 100
				GetDoubleFormat = "0.0"
			Case Is < 100000
				GetDoubleFormat = "0"
				'Case Is < 1000# * 1000# * 1000#
				'  GetDoubleFormat = "###,###,###,###"
			Case Else
				GetDoubleFormat = "0.00E+00"
		End Select
	End Function
	Function GetDoubleFormatLonger(ByVal Value As Double) As String
		Dim AbsValue As Double
		AbsValue = System.Math.Abs(Value)
		Select Case AbsValue
			Case 0#
				GetDoubleFormatLonger = "0"
			Case Is < 0.001
				GetDoubleFormatLonger = "0.00000E+00"
			Case Is < 0.01
				GetDoubleFormatLonger = "0.00000E+00"
			Case Is < 0.1
				GetDoubleFormatLonger = "0.0000000"
			Case Is < 1#
				GetDoubleFormatLonger = "0.000000"
			Case Is < 10#
				GetDoubleFormatLonger = "0.00000"
			Case Is < 100#
				GetDoubleFormatLonger = "0.0000"
			Case Is < 1000#
				GetDoubleFormatLonger = "0.000"
			Case Is < 10000#
				GetDoubleFormatLonger = "0.00"
			Case Is < 100000#
				GetDoubleFormatLonger = "0.0"
			Case Is < 1000000#
				GetDoubleFormatLonger = "0"
			Case Is < 100000000#
				GetDoubleFormatLonger = "0"
				'Case Is < 1000# * 1000# * 1000#
				'  GetDoubleFormat = "###,###,###,###"
			Case Else
				GetDoubleFormatLonger = "0.00000E+00"
		End Select
	End Function
	Public Function NumberToMFBString(ByRef Value As Object) As String
		Dim pformat As String
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case VarType(Value)
			Case VariantType.Integer, VariantType.Short
				pformat = "0"
			Case VariantType.Double, VariantType.Single
				'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pformat = GetDoubleFormat(CDbl(Value))
			Case VariantType.String
				'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NumberToMFBString = Value
				Exit Function
		End Select
		'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NumberToMFBString = VB6.Format(Value, pformat)
	End Function
	Public Sub AssignTextAndTag(ByRef gObject As Object, ByRef Value As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.Text = NumberToMFBString(Value)
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.Tag = gObject.Text
	End Sub
	Public Sub AssignCaptionAndTag(ByRef gObject As Label, ByRef Value As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.Text = NumberToMFBString(Value)
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object gObject.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gObject.Tag = gObject.Text
	End Sub


	'Sub Test_LostFocus_WithUnits( _
	''    text_box As TextBox, _
	''    ConversionFactor As Double, _
	''    LB As Double, _
	''    ub As Double, _
	''    dummy As Double, _
	''    VName As String, _
	''    Range As String, _
	''    Flag_OK As Integer)
	'  On Error GoTo ErrHandlerLost2
	'  '-- Get number into standard units
	'  If (ConversionFactor < 0) Then
	'    dummy = ReverseTemperatureConversion(-ConversionFactor - 1, CDbl(text_box))
	'  Else
	'    dummy = CDbl(text_box) / ConversionFactor
	'  End If
	'
	'  '-- Do range checking
	'  If (dummy < LB) Or (dummy > ub) Then GoTo RangeVariable2
	'  Flag_OK = True
	'  GoTo Exit_Test_LostFocus2
	'
	'ErrHandlerLost2:
	'  MsgBox "Invalid Data Value" & Chr$(10) & _
	''      "The wrong value is replaced by the previous one.", _
	''      vbExclamation, "Error in " & VName
	'  text_box.Text = Temp_Text
	'  Resume Active_Variable2
	'
	'Active_Variable2:
	'  '-- Don't need the next 3 lines any more!
	'  If (text_box.Enabled) Then
	'    text_box.SetFocus
	'  End If
	'
	'  Flag_OK = False
	'  GoTo Exit_Test_LostFocus2
	'
	'RangeVariable2:
	'  MsgBox "Value Out Of Range" & Range & Chr$(10) & "The wrong value is replaced by the previous one.", vbExclamation, "Error in " & VName
	'  text_box = Temp_Text
	'  GoTo Active_Variable2
	'
	'Exit_Test_LostFocus2:
	'End Sub
End Module