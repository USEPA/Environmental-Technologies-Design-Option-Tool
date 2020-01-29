Attribute VB_Name = "PrintModule"
Option Explicit

'Global Const USE_FONTNAME = "arial"
Global Const USE_FONTNAME = "courier new"
Global Const USE_FONTSIZE = 8
Global Const USE_FORMAT_CURRENCYSTANDARD = "$#,##0_);[Red]($#,##0)"
Global Const USE_FORMAT_CURRENCYDIGITSPAST2 = "$#,##0.00_);[Red]($#,##0.00)"


Const PrintModule_declarations_end = 0


Sub PrintCell( _
    f1 As Control, _
    r As Integer, _
    c As Integer, _
    v As Variant, _
    do_italics As Boolean, _
    do_bold As Boolean, _
    do_rightjustify)
Dim use_HAlign As Integer
  f1.EntryRC(r, c) = v
  f1.SetSelection r, c, r, c
  f1.SetFont _
      USE_FONTNAME, _
      USE_FONTSIZE, _
      do_bold, _
      do_italics, _
      False, _
      False, _
      QBColor(0), _
      False, _
      False
  use_HAlign = F1HAlignLeft
  If (do_rightjustify) Then use_HAlign = F1HAlignRight
  f1.SetAlignment _
      use_HAlign, _
      False, _
      F1VAlignBottom, _
      0
End Sub
Sub PrintCell_CurrencyStandard( _
    f1 As Control, _
    r As Integer, _
    c As Integer, _
    v As Variant)
  Call PrintCell(f1, r, c, v, False, False, True)
  f1.NumberFormat = USE_FORMAT_CURRENCYSTANDARD
End Sub
Sub PrintCell_CurrencyDigitsPast2( _
    f1 As Control, _
    r As Integer, _
    c As Integer, _
    v As Variant)
  Call PrintCell(f1, r, c, v, False, False, True)
  f1.NumberFormat = USE_FORMAT_CURRENCYDIGITSPAST2
End Sub
Sub PrintCell_QuantityStandard( _
    f1 As Control, _
    r As Integer, _
    c As Integer, _
    v As Variant)
Dim AbsValue As Double
Dim GetDoubleFormat As String
AbsValue = Abs(Val(v))
  Select Case AbsValue
    Case 0#
      GetDoubleFormat = "0"
    'Case Is < 0.001
    '  GetDoubleFormat = "0.00E+00"
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
    Case Else
      GetDoubleFormat = "0.00E+00"
  End Select
  Call PrintCell(f1, r, c, v, False, False, True)
  f1.NumberFormat = GetDoubleFormat
End Sub
Sub Print_Border0(f1 As Control, r1 As Integer, c1 As Integer, r2 As Integer, c2 As Integer, num_top_rows As Integer, num_left_cols As Integer)
Dim cc As Variant
  cc = QBColor(0)
  f1.SetSelection r1, c1, r2, c2
  f1.SetBorder 5, 0, 0, 0, 0, 1, cc, cc, cc, cc, cc
  If (r1 <> r2) Then
    f1.SetSelection r1 + num_top_rows, c1, r2, c2
    f1.SetBorder 5, -1, -1, -1, -1, 1, cc, cc, cc, cc, cc
  End If
  f1.SetSelection r1, c1 + num_left_cols, r2, c2
  f1.SetBorder 5, -1, -1, -1, -1, 1, cc, cc, cc, cc, cc
End Sub
Sub Print_Border(f1 As Control, r1 As Integer, c1 As Integer, r2 As Integer, c2 As Integer)
  Call Print_Border0(f1, r1, c1, r2, c2, 1, 2)
End Sub


Sub Print_Inputs(f1 As Control, proj As Project_Type, _
  SheetIdx As Integer, in_file As String)
Dim i As Integer
Dim f As Integer
Dim x As Integer
Dim r As Integer
Dim USER_OPTION_DRAWBORDERS As Integer
Dim MAX_COLUMN As Long
Dim MAX_COLUMNWIDTH As Long
Dim in_file1 As String
Dim in_file2 As String
Dim out_file As String
Dim strHeader As String
Dim strValue As String
Dim StrDescription As String
Dim textline1 As String
Dim textline2 As String
Dim textline3 As String
Dim StrTextLine As String
Dim NumArgs As Integer
Dim is_nom As Boolean
Dim rthis As Integer
Dim this_name As String
Dim strTextArg As String


  f1.Sheet = SheetIdx
    
  USER_OPTION_DRAWBORDERS = True
  
  r = 2
  Call PrintCell(f1, r + 0, 3, "Filename:", True, True, True)
  Call PrintCell(f1, r + 1, 3, "Unused:", True, True, True)
  Call PrintCell(f1, r + 2, 3, "Printed:", True, True, True)
  Call PrintCell(f1, r + 0, 4, Current_Filename, False, False, False)
  Call PrintCell(f1, r + 1, 4, "Unused", False, False, False)
  Call PrintCell(f1, r + 2, 4, Now, False, False, False)
  r = r + 3     'move to immediately after this section.
  
  If FileExists(in_file) Then
    f = FreeFile
    Open in_file For Input As #f
    Do While Not EOF(f)
    Line Input #f, textline1
    Select Case (InStr(1, textline1, "=======") > 0)
      Case True
        StrTextLine = Trim$(textline1)
        NumArgs = Parser_GetNumArgs(" ", StrTextLine)
        If NumArgs < 1 Then
          Call Show_Message00("File is corrupt, please call Vendor", _
            vbExclamation, App.Title)
        Else
          For i = 2 To NumArgs
            Call Parser_GetArg(" ", StrTextLine, i, strTextArg)
            strHeader = strHeader + " " + strTextArg
          Next i
          Call PrintCell(f1, r + 1, 1, strHeader, False, True, False)
          r = r + 1
        End If
       
        
      Case False
       If Val(textline1) = 0 Then
          If textline1 = "" Then r = r + 1
          Call PrintCell(f1, r, 2, textline1, False, False, False)
       Else
          Call PrintCell(f1, r, 1, textline1, False, True, False)
          r = r + 1
       End If
       
      End Select
      
    Loop
    
    Close #f
  End If

  'RESIZE THE COLUMNS.
  MAX_COLUMN = 20
  MAX_COLUMNWIDTH = 3000
  For i = 1 To MAX_COLUMN
    f1.ColWidth(i) = MAX_COLUMNWIDTH
  Next i

  'LAST STEP: RETURN CURSOR TO POSITION 1,1.
  f1.SetSelection 1, 1, 1, 1

End Sub


Sub Print_Outputs(f1 As Control, proj As Project_Type, _
  SheetIdx As Integer, out_file As String)
Dim r As Integer
Dim USER_OPTION_DRAWBORDERS As Boolean
Dim sname As String
Dim MAX_COLUMN As Long
Dim MAX_COLUMNWIDTH As Long
Dim f As String
Dim textline1 As String

  f1.Sheet = SheetIdx
    
  USER_OPTION_DRAWBORDERS = True
  
  'SECTION: "TOP HEADER".
  r = 2
  Call PrintCell(f1, r + 0, 3, "Filename:", True, True, True)
  Call PrintCell(f1, r + 1, 3, "Printed:", True, True, True)
  Call PrintCell(f1, r + 0, 4, Current_Filename, False, False, False)
  Call PrintCell(f1, r + 1, 4, Now, False, False, False)
  r = r + 3     'move to immediately after this section.
  
  If FileExists(out_file) Then
    f = FreeFile
    Open out_file For Input As #f
    Do While Not EOF(f)
    Line Input #f, textline1
    Call PrintCell(f1, r + 1, 1, textline1, False, True, False)
    r = r + 1
    Loop
  End If

End Sub


Sub PrintTo_f1book(f1 As Control, proj As Project_Type)
Dim i As Integer
Dim j As Integer
Dim SortIndex_CaseNames() As Integer
Dim Case_Index As Integer
Dim Sheet_Index As Integer

Dim SheetIdx_Inputs As Integer
Dim SheetIdx_Outputs As Integer
Dim NumSheets_Outputs As Integer
Dim SheetIdx_ThisOutput As Integer
Dim NumSheets_Total As Integer
Dim in_file As String
Dim out_file As String
Dim Any_Error As Boolean


  'SET NUMBER OF SHEETS.
  f1.NumSheets = 3
  
  'SET DEFAULT FONT.
  f1.SetDefaultFont USE_FONTNAME, USE_FONTSIZE
  
  If (frmPrint_DO_INPUTS) Then
    'PRINT THE INPUTS.
    in_file = "exes\input1.dat"
    f1.SheetName(1) = "Input1"
    Call Print_Inputs(f1, proj, 1, in_file)
    in_file = "exes\input2.dat"
    f1.SheetName(2) = "Input2"
    Call Print_Inputs(f1, proj, 2, in_file)
  End If
 ' If (frmPrint_DO_OUTPUTS) Then
      out_file = "exes\output.txt"
      f1.SheetName(3) = "Output"
      Call Print_Outputs(f1, proj, 3, out_file)
'  End If

  'RETURN TO SHEET #1.
  f1.Sheet = 1

End Sub

