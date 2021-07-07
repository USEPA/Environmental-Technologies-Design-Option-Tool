Attribute VB_Name = "DataEntry"
Option Explicit



Function Format_It(ByVal X As Double, ByVal N As Integer) As String
Dim s As String
  Select Case N
  Case 2
    Select Case Abs(X)
     Case Is < 0.1
      s = Format$(X, "0.00E+00")
     Case Is > 100#
      s = Format$(X, "0.00E+00")
     Case Else
      s = Format$(X, "0.00")
     End Select
  Case 3
    Select Case Abs(X)
     Case Is < 0.1
      s = Format$(X, "0.000E+00")
     Case Is > 100#
      s = Format$(X, "0.000E+00")
     Case Else
      s = Format$(X, "0.000")
     End Select
  End Select
  Format_It = s
End Function


Public Function Global_ReadOnlyKeyPress(ByVal KeyAscii%) As Integer
Dim KeyStroke%
  Select Case KeyAscii%
    ' CONTROL CHARACTERS: ^C, <BS>, ^V, ^X, ^Z
    Case 3, 8, 22, 24, 26
      KeyStroke% = KeyAscii%
    Case 13 ' <Enter> -> <Tab>
      SendKeys "{TAB}", True
    Case Else
      KeyStroke% = 0
  End Select
  Global_ReadOnlyKeyPress = KeyStroke%
End Function
Public Function Global_Numeric0123456789KeyPress(ByVal KeyAscii%) As Integer
Dim KeyStroke%
  Select Case KeyAscii%
    Case 48 To 57:
      KeyStroke% = KeyAscii%
    ' CONTROL CHARACTERS: ^C, <BS>, ^V, ^X, ^Z
    Case 3, 8, 22, 24, 26
      KeyStroke% = KeyAscii%
    Case 13 ' <Enter> -> <Tab>
      SendKeys "{TAB}", True
    Case Else
      KeyStroke% = 0
  End Select
  Global_Numeric0123456789KeyPress = KeyStroke%
End Function
Public Function Global_NumericKeyPress(ByVal KeyAscii%) As Integer
Dim KeyStroke%
  '
  ' THIS FUNCTION ONLY ALLOWS THE USER TO
  ' ENTER NUMERIC VALUES INTO A TEXT BOX
  '
  Select Case KeyAscii%
    ' ASCII CHARACTERS:  +, -, ., 0-9, E
    Case 43, 45, 46, 48 To 57, 69
      KeyStroke% = KeyAscii%
    ' CONTROL CHARACTERS: ^C, <BS>, ^V, ^X, ^Z
    Case 3, 8, 22, 24, 26
      KeyStroke% = KeyAscii%
    Case 101 ' e -> E
      KeyStroke% = 69
    Case 13 ' <Enter> -> <Tab>
      SendKeys "{TAB}", True
    Case Else
      KeyStroke% = 0
  End Select
  Global_NumericKeyPress = KeyStroke%
End Function
Public Function Global_MultilineTextKeyPress(ByVal KeyAscii%) As Integer
Dim KeyStroke%
  Select Case KeyAscii%
    'Case 13: ' <Enter> -> <Tab>
    '  SendKeys "{TAB}", True
    '  KeyStroke% = 0
    Case 34:  'double-quote character (")
      KeyStroke% = 0
    Case Else:
      KeyStroke% = KeyAscii%
  End Select
  Global_MultilineTextKeyPress = KeyStroke%
End Function
Public Function Global_TextKeyPress(ByVal KeyAscii%) As Integer
Dim KeyStroke%
  Select Case KeyAscii%
    Case 13: ' <Enter> -> <Tab>
      SendKeys "{TAB}", True
      KeyStroke% = 0
    Case 34:  'double-quote character (")
      KeyStroke% = 0
    Case Else:
      KeyStroke% = KeyAscii%
  End Select
  Global_TextKeyPress = KeyStroke%
End Function
Public Sub Global_GotFocus(gObject As Object)
  gObject.BackColor = RGB(0, 220, 220)
  gObject.SelStart = 0
  gObject.SelLength = Len(gObject.Text)
End Sub
Public Sub Global_LostFocus(gObject As Object)
  gObject.BackColor = RGB(255, 255, 255)
  gObject.SelLength = 0
End Sub


Function GetDoubleFormat(ByVal Value As Double) As String
Dim AbsValue As Double
  AbsValue = Abs(Value)
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
  AbsValue = Abs(Value)
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
Public Function NumberToMFBString(Value As Variant) As String
Dim pformat$
  Select Case VarType(Value)
    Case vbLong, vbInteger
      pformat$ = "0"
    Case vbDouble, vbSingle
      pformat$ = GetDoubleFormat(CDbl(Value))
    Case vbString
      NumberToMFBString = Value
      Exit Function
  End Select
  NumberToMFBString = Format$(Value, pformat$)
End Function
Public Sub AssignTextAndTag(gObject As Object, Value As Variant)
  gObject.Text = NumberToMFBString(Value)
  gObject.Tag = gObject.Text
End Sub
Public Sub AssignCaptionAndTag(gObject As Object, Value As Variant)
  gObject.Caption = NumberToMFBString(Value)
  gObject.Tag = gObject.Caption
End Sub

