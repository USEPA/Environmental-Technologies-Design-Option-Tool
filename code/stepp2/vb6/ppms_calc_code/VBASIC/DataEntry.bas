Attribute VB_Name = "DataEntry"
Option Explicit





Const DataEntry_decl_end = True


Function Format_Numerical_Value( _
    in_Dbl As Double) _
    As String
Dim Use_Format_Type As Integer
Dim sFormat As String
Dim in_SigFigCount As Integer
  With PrefEnvironment
    If (in_Dbl > 1000#) Then
      Use_Format_Type = .NumFormat_Greater1000
    ElseIf (in_Dbl < 0.001) Then
      Use_Format_Type = .NumFormat_Less0_001
    Else
      Use_Format_Type = .NumFormat_Other
    End If
  End With
  sFormat = ""
  in_SigFigCount = 0
  Select Case Use_Format_Type
    Case NUMFORMAT_3SIGFIG: in_SigFigCount = 3
    Case NUMFORMAT_4SIGFIG: in_SigFigCount = 4
    Case NUMFORMAT_5SIGFIG: in_SigFigCount = 5
    Case NUMFORMAT_6SIGFIG: in_SigFigCount = 6
    Case NUMFORMAT_EXP3: sFormat = "0.000E+00"
    Case NUMFORMAT_EXP4: sFormat = "0.0000E+00"
    Case NUMFORMAT_EXP5: sFormat = "0.00000E+00"
    Case NUMFORMAT_3PASTDEC: sFormat = "0.000"
    Case NUMFORMAT_4PASTDEC: sFormat = "0.0000"
    Case NUMFORMAT_5PASTDEC: sFormat = "0.00000"
  End Select
  If (in_SigFigCount <> 0) Then
    sFormat = GetDoubleFormat_VarSigFigs(in_Dbl, in_SigFigCount)
  End If
  Format_Numerical_Value = Format$(in_Dbl, sFormat)
End Function


Function Format_It( _
    ByVal x As Double, _
    ByVal N As Integer) _
    As String
Dim s As String
  Select Case N
  Case 2
    Select Case Abs(x)
     Case Is < 0.1
      s = Format$(x, "0.00E+00")
     Case Is > 100#
      s = Format$(x, "0.00E+00")
     Case Else
      s = Format$(x, "0.00")
     End Select
  Case 3
    Select Case Abs(x)
     Case Is < 0.1
      s = Format$(x, "0.000E+00")
     Case Is > 100#
      s = Format$(x, "0.000E+00")
     Case Else
      s = Format$(x, "0.000")
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


Function GetDoubleFormat_VarSigFigs( _
    in_Value As Double, _
    in_SigFigCount As Integer) _
    As String
Dim AbsValue As Double
Dim sReturn As String
  AbsValue = Abs(in_Value)
  Select Case in_SigFigCount
    '
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////  3 SIGNIFICANT FIGURES  //////////////////////////////////////////////////////////////////////////////////////////////////
    Case 3:
      Select Case AbsValue
        Case 0#: sReturn = "0"
        Case Is < 0.001: sReturn = "0.00E+00"
        Case Is < 0.01: sReturn = "0.00E+00"
        Case Is < 0.1: sReturn = "0.0000"
        Case Is < 1#: sReturn = "0.000"
        Case Is < 10#: sReturn = "0.00"
        Case Is < 100#: sReturn = "0.0"
        Case Is < 1000#: sReturn = "0"
        ''''Case Is < 10000#: sReturn = "0"
        ''''Case Is < 100000#: sReturn = "0"
        'Case Is < 1000# * 1000# * 1000#
        '  sReturn = "###,###,###,###"
        Case Else: sReturn = "0.00E+00"
      End Select
    '
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////  4 SIGNIFICANT FIGURES  //////////////////////////////////////////////////////////////////////////////////////////////////
    Case 4:
      Select Case AbsValue
        Case 0#: sReturn = "0"
        Case Is < 0.001: sReturn = "0.000E+00"
        Case Is < 0.01: sReturn = "0.000E+00"
        Case Is < 0.1: sReturn = "0.00000"
        Case Is < 1#: sReturn = "0.0000"
        Case Is < 10#: sReturn = "0.000"
        Case Is < 100#: sReturn = "0.00"
        Case Is < 1000#: sReturn = "0.0"
        Case Is < 10000#: sReturn = "0"
        Case Else: sReturn = "0.000E+00"
      End Select
    '
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////  5 SIGNIFICANT FIGURES  //////////////////////////////////////////////////////////////////////////////////////////////////
    Case 5:
      Select Case AbsValue
        Case 0#: sReturn = "0"
        Case Is < 0.001: sReturn = "0.0000E+00"
        Case Is < 0.01: sReturn = "0.0000E+00"
        Case Is < 0.1: sReturn = "0.000000"
        Case Is < 1#: sReturn = "0.00000"
        Case Is < 10#: sReturn = "0.0000"
        Case Is < 100#: sReturn = "0.000"
        Case Is < 1000#: sReturn = "0.00"
        Case Is < 10000#: sReturn = "0.0"
        Case Is < 100000#: sReturn = "0"
        Case Else: sReturn = "0.0000E+00"
      End Select
    '
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////  6 SIGNIFICANT FIGURES  //////////////////////////////////////////////////////////////////////////////////////////////////
    Case 6:
      Select Case AbsValue
        Case 0#: sReturn = "0"
        Case Is < 0.001: sReturn = "0.00000E+00"
        Case Is < 0.01: sReturn = "0.00000E+00"
        Case Is < 0.1: sReturn = "0.0000000"
        Case Is < 1#: sReturn = "0.000000"
        Case Is < 10#: sReturn = "0.00000"
        Case Is < 100#: sReturn = "0.0000"
        Case Is < 1000#: sReturn = "0.000"
        Case Is < 10000#: sReturn = "0.00"
        Case Is < 100000#: sReturn = "0.0"
        Case Is < 1000000#: sReturn = "0"
        ''''Case Is < 100000000#: sReturn = "0"
        'Case Is < 1000# * 1000# * 1000#
        '  GetDoubleFormat = "###,###,###,###"
        Case Else: sReturn = "0.00000E+00"
      End Select
  End Select
  GetDoubleFormat_VarSigFigs = sReturn
End Function


Function GetDoubleFormat(ByVal value As Double) As String
  GetDoubleFormat = GetDoubleFormat_VarSigFigs(value, 3)
''''Dim AbsValue As Double
''''  AbsValue = Abs(value)
''''  Select Case AbsValue
''''    Case 0#
''''      GetDoubleFormat = "0"
''''    Case Is < 0.001
''''      GetDoubleFormat = "0.00E+00"
''''    Case Is < 0.01
''''      GetDoubleFormat = "0.00E+00"
''''    Case Is < 0.1
''''      GetDoubleFormat = "0.0000"
''''    Case Is < 1
''''      GetDoubleFormat = "0.000"
''''    Case Is < 10
''''      GetDoubleFormat = "0.00"
''''    Case Is < 100
''''      GetDoubleFormat = "0.0"
''''    Case Is < 100000
''''      GetDoubleFormat = "0"
''''    'Case Is < 1000# * 1000# * 1000#
''''    '  GetDoubleFormat = "###,###,###,###"
''''    Case Else
''''      GetDoubleFormat = "0.00E+00"
''''  End Select
End Function
Function GetDoubleFormatLonger(ByVal value As Double) As String
  GetDoubleFormatLonger = GetDoubleFormat_VarSigFigs(value, 6)
''''Dim AbsValue As Double
''''  AbsValue = Abs(value)
''''  Select Case AbsValue
''''    Case 0#
''''      GetDoubleFormatLonger = "0"
''''    Case Is < 0.001
''''      GetDoubleFormatLonger = "0.00000E+00"
''''    Case Is < 0.01
''''      GetDoubleFormatLonger = "0.00000E+00"
''''    Case Is < 0.1
''''      GetDoubleFormatLonger = "0.0000000"
''''    Case Is < 1#
''''      GetDoubleFormatLonger = "0.000000"
''''    Case Is < 10#
''''      GetDoubleFormatLonger = "0.00000"
''''    Case Is < 100#
''''      GetDoubleFormatLonger = "0.0000"
''''    Case Is < 1000#
''''      GetDoubleFormatLonger = "0.000"
''''    Case Is < 10000#
''''      GetDoubleFormatLonger = "0.00"
''''    Case Is < 100000#
''''      GetDoubleFormatLonger = "0.0"
''''    Case Is < 1000000#
''''      GetDoubleFormatLonger = "0"
''''    Case Is < 100000000#
''''      GetDoubleFormatLonger = "0"
''''    'Case Is < 1000# * 1000# * 1000#
''''    '  GetDoubleFormat = "###,###,###,###"
''''    Case Else
''''      GetDoubleFormatLonger = "0.00000E+00"
''''  End Select
End Function
Public Function NumberToMFBString(value As Variant) As String
Dim pformat$
  Select Case VarType(value)
    Case vbLong, vbInteger
      pformat$ = "0"
    Case vbDouble, vbSingle
      pformat$ = GetDoubleFormat(CDbl(value))
    Case vbString
      NumberToMFBString = value
      Exit Function
  End Select
  NumberToMFBString = Format$(value, pformat$)
End Function
Public Sub AssignTextAndTag(gObject As Object, value As Variant)
  gObject.Text = NumberToMFBString(value)
  gObject.Tag = gObject.Text
End Sub
Public Sub AssignCaptionAndTag(gObject As Object, value As Variant)
  gObject.Caption = NumberToMFBString(value)
  gObject.Tag = gObject.Caption
End Sub

