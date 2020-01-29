Attribute VB_Name = "DataEntryHandlers"
'testing2
Option Explicit

Public Const ERROR_REPLACE_VALUE = True
Public Const ERROR_LEAVE_VALUE = False


Public DataEntryError As Boolean
Public PreviousErrorObject As Object
Public EntryErrorMessage$
Public EntryErrorReplaceValue As Boolean

Public Sub AssignTextAndTag_WithRange(gObject_text As Object, Value As Variant, xmin As Double, xmax As Double)

  Call AssignTextAndTag(gObject_text, Value)
  gObject_text.LinkItem = Trim$(Str$(xmin))
  gObject_text.DataField = Trim$(Str$(xmax))

End Sub

Public Sub AssignTag_OptionChecks(gObject As Object, Value As Variant)
  gObject.Tag = Value
End Sub

Public Sub Global_GotFocus(gObject As Object)

gObject.BackColor = RGB(0, 220, 220)
gObject.SelStart = 0
gObject.SelLength = Len(gObject.Text)

End Sub


Public Sub Global_LostFocus(gObject As Object)

gObject.BackColor = RGB(255, 255, 255)
gObject.SelLength = 0

End Sub

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
    Case 13 ' <Enter>
        SendKeys "{TAB}", True
    Case Else
        KeyStroke% = 0
End Select

Global_NumericKeyPress = KeyStroke%

End Function


Public Function Global_TextKeyPress(ByVal KeyAscii%) As Integer
Dim KeyStroke%

Select Case KeyAscii%
  Case 13:
    SendKeys "{TAB}", True
    KeyStroke% = 0
  Case Asc(Chr$(34)):     'double-quote character
    KeyStroke% = 0
  Case Else:
    KeyStroke% = KeyAscii%
End Select
    
Global_TextKeyPress = KeyStroke%
'  KeyStroke% = keyascii%''
'
'  If (keyascii% = 13) Then
'    SendKeys "{TAB}", True
'  End If
End Function


Function NumberToCommaDelimited(Value As Double) As String
  NumberToCommaDelimited = _
      Format$(Value, _
      "###,###,###,###,###,##0")
End Function
'Function NumberToCurrency(Value As Double) As String
'  Select Case NowProj.CurrencyStandard_Type
'    Case CURRENCYSTANDARD_TYPE_1:
'      NumberToCurrency = Format$( _
'          Value, _
'          "$#,##0;($#,##0)")
'    Case CURRENCYSTANDARD_TYPE_1K:
'      NumberToCurrency = Format$( _
'          Value / 1000#, _
'          "$#,##0K;($#,##0K)")
'    Case CURRENCYSTANDARD_TYPE_1M:
'      NumberToCurrency = Format$( _
'          Value / 1000# / 1000#, _
'          "$#,##0M;($#,##0M)")
'    Case CURRENCYSTANDARD_TYPE_1B:
'      NumberToCurrency = Format$( _
'          Value / 1000# / 1000# / 1000#, _
'          "$#,##0B;($#,##0B)")
'  End Select
'End Function
'Function Get_DCF_DISPLAY_FACTOR() As Double
'  Select Case NowProj.CurrencyDCF_Type
'    Case CURRENCYDCF_TYPE_1:
'      Get_DCF_DISPLAY_FACTOR = 1#
'    Case CURRENCYDCF_TYPE_1K:
'      Get_DCF_DISPLAY_FACTOR = 1000#
'    Case CURRENCYDCF_TYPE_1M:
'      Get_DCF_DISPLAY_FACTOR = 1000# * 1000#
'    Case CURRENCYDCF_TYPE_1B:
'      Get_DCF_DISPLAY_FACTOR = 1000# * 1000# * 1000#
'  End Select
'End Function
'Function Get_DCF_DISPLAY_FACTOR_AsString() As String
'  Select Case NowProj.CurrencyDCF_Type
'    Case CURRENCYDCF_TYPE_1:
'      Get_DCF_DISPLAY_FACTOR_AsString = "1"
'    Case CURRENCYDCF_TYPE_1K:
'      Get_DCF_DISPLAY_FACTOR_AsString = "1000"
'    Case CURRENCYDCF_TYPE_1M:
'      Get_DCF_DISPLAY_FACTOR_AsString = "1,000,000"
'    Case CURRENCYDCF_TYPE_1B:
'      Get_DCF_DISPLAY_FACTOR_AsString = "1,000,000,000"
'  End Select
'End Function
'Function NumberToCurrency0(Value As Double) As String
'  NumberToCurrency0 = Format$(Value, _
'      "$###,###,##0;($###,###,##0)")
'End Function
'Function NumberToCurrency2(Value As Double) As String
'  NumberToCurrency2 = _
'      Format$(Value, _
'      "$###,###,##0.00;($###,###,##0.00)")
'End Function


Function NumberToPercentage(Value As Double) As String

  NumberToPercentage = Format$(Value, "0.0%")

End Function


Public Sub SetDataEntryError(gObject As Object)

' THIS IS THE GENERIC MESSAGE FOR THE DATA ENTRY ERROR
' SETTING IT HERE TO SAVE MEMORY
' THE SETENTRYERRORMESSAGE FUNCTION SHOULD BE CALLED BEFORE
' THIS FUNCTION IF YOU DO NOT WANT THIS MESSAGE TO APPEAR
If (EntryErrorMessage$ = "") Then
    EntryErrorReplaceValue = True
    EntryErrorMessage$ = _
          "You have entered data into the text field that" + vbCrLf _
        + "could not be converted to a valid number!"
End If

' RETURN TO PREVIOUS GOOD VALUE
If (EntryErrorReplaceValue) Then
    gObject.Text = gObject.Tag
    EntryErrorMessage$ = EntryErrorMessage$ + vbCrLf + vbCrLf _
        + "The Value Will Be Replaced With the Last Value."
End If

' SET VARIABLES FOR CHECKING IN THE NEXT GOTFOCUS EVENT
DataEntryError = True
Set PreviousErrorObject = gObject

End Sub

Public Sub SetEntryErrorMessage(ErrMsg$, ReplacementType As Boolean)

EntryErrorMessage$ = ErrMsg$
EntryErrorReplaceValue = ReplacementType

End Sub


Public Function DisplayDataEntryError() As Boolean

If (DataEntryError) Then
    MsgBox EntryErrorMessage$, vbCritical, "Error Message"
    EntryErrorMessage$ = ""
    DataEntryError = False
    PreviousErrorObject.SetFocus
    Set PreviousErrorObject = Nothing
End If

End Function



Function ValueHasChanged(gObject As Object) As Boolean

ValueHasChanged = IIf(StrComp(gObject.Tag, gObject.Text, 1) = 0, False, True)

End Function


Function ValidateRange(Value As Variant, dbl_LB As Variant, dbl_UB As Variant) As Boolean
    Dim strLB$, strUB$, msg$
    
If ((Value < dbl_LB) Or (Value > dbl_UB)) Then
    ValidateRange = False
    strLB$ = NumberToMFBString(dbl_LB)
    strUB$ = NumberToMFBString(dbl_UB)
    msg$ = "The Value Is Out Of the Range:" + vbCrLf + vbCrLf _
         + vbTab + "Minimum = " + strLB$ + vbCrLf _
         + vbTab + "Maximum = " + strUB$
    Call SetEntryErrorMessage(msg$, ERROR_REPLACE_VALUE)
Else
    ValidateRange = True
End If

End Function

Function IsValidNumber(gObject As Object, VariableType%) As Boolean
    Dim tmpVar As Variant
    
On Error GoTo VALIDATION_ERROR
Select Case VariableType%
    Case vbInteger
        tmpVar = CInt(gObject.Text)
    Case vbLong
        tmpVar = CLng(gObject.Text)
    Case vbSingle
        tmpVar = CSng(gObject.Text)
    Case vbDouble
        tmpVar = CDbl(gObject.Text)
End Select

IsValidNumber = True
Exit Function

VALIDATION_ERROR:
    IsValidNumber = False

End Function

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
    Case Is < 1000
        GetDoubleFormat = "0"
    Case Is < 1000# * 1000# * 1000#
        GetDoubleFormat = "0"
        'GetDoubleFormat = "###,###,###,###"
    Case Else
        GetDoubleFormat = "0.00E+00"
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

Public Sub AssignTag_Scrollbox(gObject As Object, Value As Variant)

  gObject.Tag = Value

End Sub


Public Sub AssignTextAndTag(gObject As Object, Value As Variant)
gObject.Text = NumberToMFBString(Value)
gObject.Tag = gObject.Text
End Sub


Public Sub AssignTextAndTag_Slider(gObject_text As Object, gObject_slider As Object, Value As Variant, xmin As Double, xmax As Double)
Dim smallest_sensible_step_size As Double

  Call AssignTextAndTag_WithRange(gObject_text, Value, xmin, xmax)
  
  gObject_slider.Min = xmin
  gObject_slider.Max = xmax
  smallest_sensible_step_size = (xmax - xmin) / 100#
  If (gObject_slider.SmallChange > smallest_sensible_step_size) Or _
     (gObject_slider.LargeChange > smallest_sensible_step_size) Then
    gObject_slider.SmallChange = smallest_sensible_step_size
    gObject_slider.LargeChange = smallest_sensible_step_size
  End If
  gObject_slider.Value = gObject_text.Text
  
End Sub


Public Sub AssignTextAndTag_SpinButton( _
    gObject_text As Object, _
    gObject_SpinButton As Object, _
    Value As Variant, _
    xmin As Double, _
    xmax As Double)
'Dim smallest_sensible_step_size As Double
  Call AssignTextAndTag_WithRange(gObject_text, Value, xmin, xmax)
  'gObject_SpinButton.min = xmin
  'gObject_SpinButton.max = xmax
  'smallest_sensible_step_size = (xmax - xmin) / 100#
  'If (gObject_slider.SmallChange > smallest_sensible_step_size) Or _
  '   (gObject_slider.LargeChange > smallest_sensible_step_size) Then
  '  gObject_slider.SmallChange = smallest_sensible_step_size
  '  gObject_slider.LargeChange = smallest_sensible_step_size
  'End If
  'gObject_SpinButton.Value = gObject_text.Text
End Sub

