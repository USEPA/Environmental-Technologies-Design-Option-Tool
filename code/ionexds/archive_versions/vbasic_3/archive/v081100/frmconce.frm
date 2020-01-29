VERSION 2.00
Begin Form frmConcentrations 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Influent Concentrations"
   ClientHeight    =   5385
   ClientLeft      =   1590
   ClientTop       =   2205
   ClientWidth     =   7755
   Height          =   6075
   Left            =   1530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7755
   Top             =   1575
   Width           =   7875
   Begin SSFrame Frame3D1 
      Caption         =   "Variable Influent File Name"
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   2100
      ShadowColor     =   1  'Black
      TabIndex        =   8
      Top             =   3480
      Width           =   3375
      Begin CommandButton cmdInputFile 
         Caption         =   "Input New File Name"
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Width           =   3075
      End
      Begin Label lblInputFile 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   3015
      End
   End
   Begin CommonDialog CMDialog1 
      Left            =   5220
      Top             =   2400
   End
   Begin SSCommand cdmEdit 
      Caption         =   "Fill Down"
      Height          =   372
      Index           =   3
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   1092
   End
   Begin SSCommand cdmEdit 
      Caption         =   "Paste"
      Height          =   372
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   1092
   End
   Begin SSCommand cdmEdit 
      Caption         =   "Copy"
      Height          =   372
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   1092
   End
   Begin SSCommand cdmEdit 
      Caption         =   "Cut"
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1092
   End
   Begin SSCommand cmdOK 
      Caption         =   "OK"
      Height          =   492
      Left            =   6300
      TabIndex        =   3
      Top             =   4680
      Width           =   1332
   End
   Begin TextBox txtCel 
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   2520
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   972
   End
   Begin Grid Grid1 
      Height          =   2475
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin SSCommand cmdCancel 
      Caption         =   "Cancel"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1212
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
      End
      Begin Menu mnuFileItem 
         Caption         =   "&Open..."
         Index           =   1
      End
      Begin Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin Menu mnuFileItem 
         Caption         =   "Save &As..."
         Index           =   3
      End
   End
   Begin Menu mnuEdit 
      Caption         =   "&Edit"
      Begin Menu mnuEditItem 
         Caption         =   "&Paste "
         Index           =   0
         Shortcut        =   ^V
      End
   End
End
Option Explicit
Option Base 1

Dim Shifting As Integer, X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
Dim TempStr As String, SaveAs  As Integer, Filename_Concentration As String
Dim Temp_Array() As String
Dim OriginalVarInfluentFile As String

Sub cdmEdit_Click (Index As Integer)
Dim i As Integer, j As Integer
   Select Case Index
     Case 0
       cdmEdit(2).Enabled = True
       X1 = Grid1.SelStartCol
       X2 = Grid1.SelEndCol
       Y1 = Grid1.SelStartRow
       Y2 = Grid1.SelEndRow

       ReDim Temp_Array(Grid1.SelEndCol - Grid1.SelStartCol + 1, Grid1.SelEndRow - Grid1.SelStartRow + 1)
       For i = X1 To X2
         Grid1.Col = i
         For j = Y1 To Y2
          Grid1.Row = j
          Temp_Array(i - X1 + 1, j - Y1 + 1) = Grid1.Text
          Grid1.Text = ""
         Next j
       Next i
     Case 1
       cdmEdit(2).Enabled = True
       X1 = Grid1.SelStartCol
       X2 = Grid1.SelEndCol
       Y1 = Grid1.SelStartRow
       Y2 = Grid1.SelEndRow

       ReDim Temp_Array(Grid1.SelEndCol - Grid1.SelStartCol + 1, Grid1.SelEndRow - Grid1.SelStartRow + 1)
       For i = X1 To X2
         Grid1.Col = i
         For j = Y1 To Y2
          Grid1.Row = j
          Temp_Array(i - X1 + 1, j - Y1 + 1) = Grid1.Text
         Next j
       Next i
     Case 2
      
      If ((Grid1.SelEndRow - Grid1.SelStartRow) <> (Y2 - Y1)) Or ((Grid1.SelEndCol - Grid1.SelStartCol) <> (X2 - X1)) Then
        MsgBox "The Selected zone does not fit with the data in memory", MB_ICONEXCLAMATION, Application_Name
        Exit Sub
      Else
       For i = Grid1.SelStartCol To Grid1.SelEndCol
         Grid1.Col = i
         For j = Grid1.SelStartRow To Grid1.SelEndRow
          Grid1.Row = j
          Grid1.Text = Temp_Array(i - Grid1.SelStartCol + 1, j - Grid1.SelStartRow + 1)
         Next j
       Next i
     End If
     Case 3
      TempStr = Grid1.Text
      For i = Grid1.SelStartRow + 1 To Grid1.SelEndRow
         Grid1.Row = i
         Grid1.Text = TempStr
      Next i
   End Select

End Sub

Sub cmdCancel_Click ()
  If Cations.Available And Anions.Available Then
  ElseIf Cations.Available Then
     VarInfluentFileCation = OriginalVarInfluentFile
  ElseIf Anions.Available Then

  End If

  Unload Me
End Sub

Sub cmdInputFile_Click ()
    Dim OldVarInfFile As String

    If Cations.Available And Anions.Available Then

    ElseIf Cations.Available Then
       OldVarInfFile = VarInfluentFileCation
       frmConcentrations!CMDialog1.Filename = VarInfluentFileCation
       Call SaveFileVariableInfluent(VarInfluentFileCation)

    ElseIf Anions.Available Then

    End If

    If VarInfluentFileCation = "" Then
       VarInfluentFileCation = OldVarInfFile
    Else
       Call PutInputFileName
    End If

End Sub

Sub cmdOK_Click ()
Dim i As Integer, Response As Integer
ReDim NData(7)  As Integer
Dim DFlag As Integer, f As Integer
Dim j As Integer, No_Var_Influent As Integer

  No_Var_Influent = False
  If Not (CountConc(1, Number_Influent_Points)) Then Exit Sub

  Grid1.Row = 1
  For i = 1 To Total_NumberOfComponents
   Grid1.Col = i
   If Grid1.Text = "" Then No_Var_Influent = True
  Next i
  If No_Var_Influent Then
   Response = MsgBox("There is no data for the first row" & Chr$(13) & "It will be assummed there is no variable influent concentration", MB_ICONSTOP + MB_OKCANCEL, Application_Name)
  End If
  Select Case Response
   Case IDOK
     GoTo NoInfluent_Conc
   Case IDCANCEL
     Exit Sub
  End Select

  For j = 1 To Total_NumberOfComponents + 1
    Grid1.Col = j
    NData(j) = 0
    Grid1.Row = 0
    Do Until Grid1.Text = "" Or Grid1.Row >= Number_Max_Influent_Points
     Grid1.Row = Grid1.Row + 1
     NData(j) = NData(j) + 1
    Loop
  Next j
  DFlag = False
  For i = 1 To Total_NumberOfComponents + 1
   For j = i + 1 To Total_NumberOfComponents + 1
    If NData(i) <> NData(j) Then DFlag = True
   Next j
  Next i
  If DFlag Then Response = MsgBox("There is not the same number of data in each column." & Chr$(13) & "It will be assummed that there is no variable influent concentrations", MB_ICONEXCLAMATION + MB_OKCANCEL, Application_Name)
  Select Case Response
     Case IDOK
     Case IDCANCEL
      Exit Sub
  End Select
  
  'Store Time values
On Error GoTo Time_Error
  Grid1.Col = 1
  For i = 1 To Number_Influent_Points
   Grid1.Row = i
   T_Influent(i) = CDbl(Grid1.Text) * 24# * 60# ' To convert from days to minutes
   If i > 1 Then
     If T_Influent(i) <= T_Influent(i - 1) Then GoTo Time_Error2
   End If
  Next i

  'Store influent conc.
On Error GoTo Conc_Error
  For j = 2 To Total_NumberOfComponents + 1
   Grid1.Col = j
   For i = 1 To Number_Influent_Points
     Grid1.Row = i
     C_Influent(j - 1, i) = CDbl(Grid1.Text)
   Next i
  Next j

  'Write influent concentrations to file
  If (Cations.Available And Anions.Available) Then

  ElseIf Cations.Available Then
On Error GoTo FileWritingError
SpecifyVarInfluentCation:
     If VarInfluentFileCation = "NONE" Then
        Call SaveFileVariableInfluent(VarInfluentFileCation)
     End If
     
     If VarInfluentFileCation <> "" Then
        Call SaveVariableInfluent(VarInfluentFileCation)
     Else
        MsgBox "A file name must be specified for storage of the variable influent data.", MB_ICONINFORMATION
        GoTo SpecifyVarInfluentCation
     End If
  ElseIf Anions.Available Then

  End If
  
  Unload Me
Exit_This_OK:
  Exit Sub

Time_Error:
 MsgBox "At least one value in time input (" & Format$(i, "0") & ") is not a real number." & Chr$(13) & "Return to the editing mode.", MB_ICONEXCLAMATION, Application_Name
 Resume Exit_This_OK:

Time_Error2:
 MsgBox "Time(" & Format$(i, "0") & ") < (=) Time(" & Format$(i - 1, "0") & ")" & Chr$(13) & "Return to editing mode", MB_ICONEXCLAMATION + MB_OK, Application_Name
 Exit Sub

Conc_Error:
 MsgBox "At least one value in concentration (" & Format$(i, "0") & "," & Format$(j, "0") & ") is not a real number." & Chr$(13) & "Return to the editing mode.", MB_ICONEXCLAMATION, Application_Name
 Resume Exit_This_OK:

FileWritingError:
 Resume SpecifyVarInfluentCation

NoInfluent_Conc:
   Number_Influent_Points = 0
   Unload Me
End Sub

Function CountConc (i As Integer, NPoints As Integer) As Integer

On Error GoTo Error_In_CountConc
  NPoints = 0
  Grid1.Col = i
  Grid1.Row = 1
  Do Until Grid1.Text = "" Or Grid1.Row = Number_Max_Influent_Points
    NPoints = NPoints + 1
    Grid1.Row = Grid1.Row + 1
  Loop
  If Grid1.Text <> "" Then NPoints = NPoints + 1

  CountConc = True
  Exit Function

Error_In_CountConc:
  CountConc = False
  MsgBox "Invalid data.", 64, Application_Name
  Resume Exit_CountConc
Exit_CountConc:
End Function

Function CutString () As Integer
Dim ClipString As String, Length As Integer
Dim CurrentPosition As Integer, PreviousPosition As Integer, Character As String * 1
Dim StringToTransfer As String, Row As Integer, Col As Integer

On Error GoTo Error_In_CutString

  ClipString = Clipboard.GetText()
  Length = Len(ClipString)

  If Length > 0 Then
    PreviousPosition = 1: CurrentPosition = 1
    Col = 1: Row = 1
    While PreviousPosition <= Length
      Character = Mid$(ClipString, CurrentPosition, 1)
      Select Case Asc(Character)
        Case 10
          CurrentPosition = CurrentPosition + 1
          PreviousPosition = CurrentPosition
        Case 13, 9
          StringToTransfer = Mid$(ClipString, PreviousPosition, CurrentPosition - PreviousPosition)
          If Not (PasteString(StringToTransfer, Row, Col)) Then
            MsgBox "Error while pasting data.", 64, Application_Name
          End If
          Col = Col Mod (Total_NumberOfComponents + 1) + 1
          If Col = 1 Then
           Row = Row + 1
           If Row > Number_Max_Influent_Points Then GoTo Too_Many_Points
          End If
          CurrentPosition = CurrentPosition + 1
          PreviousPosition = CurrentPosition
        Case Else
          CurrentPosition = CurrentPosition + 1
          Character = Mid$(ClipString, CurrentPosition, 1)
      End Select
    Wend
  Else
  End If


  CutString = True
  Exit Function
  
Too_Many_Points:
  CutString = True
  MsgBox "Too much data was selected. Only the first " & Format$(Number_Max_Influent_Points, "0") & " points were pasted.", 64, Application_Name
  GoTo Exit_CutString
Error_In_CutString:
  CutString = False
  Resume Exit_CutString
Exit_CutString:
End Function

Sub Form_Activate ()

    Call PutInputFileName

End Sub

Sub Form_Load ()

' Me.HelpContextID = Hlp_Influent_Concentrations
  cdmEdit(0).Enabled = True
  cdmEdit(1).Enabled = True
  cdmEdit(2).Enabled = False
  cdmEdit(3).Enabled = True

  If Cations.Available And Anions.Available Then
  ElseIf Cations.Available Then
     OriginalVarInfluentFile = VarInfluentFileCation
  ElseIf Anions.Available Then

  End If

  Call UpdateVarInf

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub

Sub Form_Resize ()
  If WindowState = 0 Then
    Grid1.Width = Width - 350
  End If
  If WindowState = 1 Then frmIonExchangeMain.WindowState = 1
    'If WindowState <> 0 Then Exit Sub
    'If (Grid1.Left + Grid1.Width) > (cdmEdit(3).Left + cdmEdit(3).Width) Then
    '  Width = Grid1.Left + Grid1.Width + 20 * Screen.TwipsPerPixelX
    'Else
    '  Width = cdmEdit(3).Left + cdmEdit(3).Width + 20 * Screen.TwipsPerPixelX
    'End If
    'If Height > (Grid1.Top + 90 * Screen.TwipsPerPixelY + cmdCancel.Height) Then
    ' Grid1.Height = Height - Grid1.Top - 90 * Screen.TwipsPerPixelY - cmdCancel.Height
    ' cmdCancel.Top = Grid1.Top + Grid1.Height + 15 * Screen.TwipsPerPixelY
    ' cmdOK.Top = Grid1.Top + Grid1.Height + 15 * Screen.TwipsPerPixelY
    'End If

    'Top = Screen.Height / 2 - Height / 2
    'Left = Screen.Width / 2 - Width / 2
End Sub

Sub Grid1_DblClick ()
  If Grid1.Row > 0 And Grid1.Col > 0 Then Grid1_KeyPress (13)
End Sub

Sub Grid1_KeyDown (Keycode As Integer, Shift As Integer)
  txtCel.Visible = False
  If Shift = 1 Then Shifting = True Else Shifting = False

End Sub

Sub Grid1_KeyPress (KeyAscii As Integer)
Dim Char As String
  Select Case KeyAscii
    Case 27
      txtCel.Text = Grid1.Text
    Case 9
      If Shifting Then
        If Grid1.Col > 1 Then
          Grid1.Col = Grid1.Col - 1
        End If
      Else
        If Grid1.Col < (Grid1.Cols - 1) Then
         Grid1.Col = Grid1.Col + 1
        End If
      End If
      Unselect
    Case Else
     If KeyAscii = 13 Then
       txtCel = Grid1.Text
       txtCel.SelStart = Len(txtCel.Text)
     Else
      Char = Chr$(KeyAscii)
      txtCel = Char
      txtCel.SelStart = 1
     End If
     ShowTextBox
     KeyAscii = 0
  End Select
End Sub

Sub Grid1_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtCel.Visible = False
End Sub

Sub Grid1_RowColChange ()
  txtCel.Text = Grid1.Text
End Sub

Sub mnuEditItem_Click (Index As Integer)
  Select Case Index
    Case 0
      If CutString() Then
      Else
       MsgBox "Impossible to paste data from the clipboard.", 64, Application_Name
      End If
  End Select
End Sub

Sub mnuFileItem_Click (Index As Integer)
Dim i As Integer, j As Integer, f As Integer, OldVarInfFile As String
  Select Case Index
    Case 0
      Grid1.Col = 0
      j = 0
      While j <= Total_NumberOfComponents
        Grid1.Col = Grid1.Col + 1
        Grid1.Row = 0
        i = 0
        While i <= Number_Max_Influent_Points - 1
          Grid1.Row = Grid1.Row + 1
          Grid1.Text = ""
          i = i + 1
        Wend
        j = j + 1
      Wend
    Case 1  'Open Variable Influent File
       If Cations.Available And Anions.Available Then
       ElseIf Cations.Available Then
          OldVarInfFile = VarInfluentFileCation
          Call OpenFileVariableInfluent(VarInfluentFileCation)
          If VarInfluentFileCation = "" Then
             VarInfluentFileCation = OldVarInfFile
          Else
             Call ReadVarInfluentConcs
             Call PutInputFileName
             Call UpdateVarInf
          End If
        End If
    Case 2   'Save
       Call SaveVariableInfluent(VarInfluentFileCation)
       Call PutInputFileName
    Case 3   'Save As
      If Cations.Available And Anions.Available Then
      ElseIf Cations.Available Then
         OldVarInfFile = VarInfluentFileCation
         Call SaveFileVariableInfluent(VarInfluentFileCation)
         If VarInfluentFileCation = "" Then
            VarInfluentFileCation = OldVarInfFile
         Else
            Call SaveVariableInfluent(VarInfluentFileCation)
            Call PutInputFileName
         End If
      ElseIf Anions.Available Then
      End If
      
      
  End Select
End Sub

Sub OpenFileVariableInfluent (VarInfluentFilename As String)

       On Error Resume Next
       frmConcentrations!CMDialog1.Filename = ""
       frmConcentrations!CMDialog1.DefaultExt = "var"
       frmConcentrations!CMDialog1.Filter = "Variable Influent (*.var)|*.var"
       frmConcentrations!CMDialog1.DialogTitle = "Ion Exchange Variable Influent Files"
       frmConcentrations!CMDialog1.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
       frmConcentrations!CMDialog1.Action = 1
       VarInfluentFilename$ = frmConcentrations!CMDialog1.Filename
       If Err = 32755 Then   'Cancel selected by user
          VarInfluentFilename$ = ""
       End If

End Sub

Function PasteString (StringToTransfer As String, Row As Integer, Col As Integer) As Integer

On Error GoTo Error_In_PasteString
  Grid1.Row = Row
  Grid1.Col = Col
  Grid1.Text = StringToTransfer
  PasteString = True
  Exit Function

Error_In_PasteString:
  PasteString = False
  Resume Exit_PasteString
Exit_PasteString:
End Function

Sub PutInputFileName ()
    Dim RightFile As String
    Dim TheLength As Integer
    Dim i As Integer
    Dim Done As Integer

    Done = False
    If Cations.Available And Anions.Available Then

    ElseIf Cations.Available Then
       
       For i = Len(VarInfluentFileCation) To 1 Step -1
           If Mid$(VarInfluentFileCation, i, 1) = "\" Then
              TheLength = Len(VarInfluentFileCation) - i
              RightFile = Right$(VarInfluentFileCation, TheLength)
              Done = True
              Exit For
           End If
       Next i

       If Not Done Then
          RightFile = VarInfluentFileCation
       End If

       lblInputFile = RightFile
       
    ElseIf Anions.Available Then

    End If

End Sub

Sub ShowTextBox ()
  Dim TestX As Integer, TestY As Integer
  Dim C As Integer
  txtCel.Visible = False
  txtCel.Height = Grid1.RowHeight(Grid1.Row) - (Screen.TwipsPerPixelY * 2)
  txtCel.Width = Grid1.ColWidth(Grid1.Col) - (Screen.TwipsPerPixelX * 2)

  TestX = Grid1.Left + Grid1.ColWidth(0) + (Screen.TwipsPerPixelX * 3)
  
  For C = Grid1.LeftCol To Grid1.Col - 1
    TestX = TestX + Grid1.ColWidth(C) + Screen.TwipsPerPixelX
  Next C

  TestY = Grid1.Top + Grid1.RowHeight(0) + (Screen.TwipsPerPixelY * 2)
  For C = Grid1.TopRow To Grid1.Row - 1
    TestY = TestY + Grid1.RowHeight(C) + Screen.TwipsPerPixelY
  Next C

  txtCel.Left = TestX
  txtCel.Top = TestY

  txtCel.ZOrder
  txtCel.Visible = True
  txtCel.SetFocus

End Sub

Sub txtCel_KeyDown (Keycode As Integer, Shift As Integer)
   If Shift = 1 Then
     Shifting = True
   Else
     Shifting = False
   End If
   Select Case Keycode
   Case 38
    txtCel_KeyPress (13)
    SendKeys "{UP}"
   Case 40
    txtCel_KeyPress (13)
    SendKeys "{DOWN}"
   End Select
End Sub

Sub txtCel_KeyPress (KeyAscii As Integer)
  Select Case KeyAscii
    Case 13, 9
     Grid1.Text = txtCel.Text
     txtCel.Visible = False
     Grid1.SetFocus
     If KeyAscii = 9 And Grid1.Col < Grid1.Cols - 1 Then
       If Shifting Then
         If Grid1.Col > 1 Then
           Grid1.Col = Grid1.Col - 1
         End If
       Else
         If Grid1.Col < (Grid1.Cols - 1) Then
           Grid1.Col = Grid1 + 1
          End If
       End If
       Unselect
     End If
     KeyAscii = 0
     Case 27
     KeyAscii = 0
     txtCel.Visible = False
     Grid1.SetFocus
   End Select

End Sub

Sub Unselect ()
  If Grid1.Visible = False Then Exit Sub
  Grid1.SetFocus
  Select Case Grid1.Col
    Case 1
      SendKeys "{RIGHT}{LEFT}"
    Case Grid1.Cols - 1
      SendKeys "{LEFT}{RIGHT}"
    Case Else
      SendKeys "{LEFt}{RIGHT}"
  End Select
End Sub

Sub UpdateVarInf ()
Dim i As Integer, j As Integer, TB As String, CB As String
Dim temp As String, LF  As String, C As Integer, SetWidth As Integer

  TB = Chr$(9)
  CB = Chr$(13)
 
  Grid1.Cols = 1
  Grid1.Rows = 1
  temp = TB & "Time (days)"

  For i = 1 To Total_NumberOfComponents
    temp = temp & TB & Trim$(Ion(i).Name) & " (mg/L)"
  Next i
  Grid1.AddItem temp
  
  For i = 0 To Total_NumberOfComponents + 1
   Grid1.FixedAlignment(i) = 2
  Next i
  
  For i = 1 To Number_Max_Influent_Points
    temp = Format(i, "0") & TB
    Grid1.AddItem temp
  Next i
  Grid1.RemoveItem 0

  Grid1.FixedRows = 1
  Grid1.FixedCols = 1
  Grid1.ColWidth(0) = 30 * Screen.TwipsPerPixelX
  For C = 1 To Grid1.Cols - 1
   Grid1.ColWidth(C) = 120 * Screen.TwipsPerPixelX
  Next C
  SetWidth = Screen.TwipsPerPixelX * 19
  For C = 0 To Grid1.Cols - 1
    SetWidth = SetWidth + Grid1.ColWidth(C) + (Screen.TwipsPerPixelX * 2)
  Next C
  Grid1.Width = SetWidth

  If Number_Influent_Points > 0 Then
   For i = 1 To Number_Influent_Points
      Grid1.Row = i
      Grid1.Col = 1
      Grid1.Text = T_Influent(i) / 24# / 60#    'Convert form min. to days
    For j = 2 To Total_NumberOfComponents + 1
      Grid1.Col = j
      Grid1.Text = C_Influent(j - 1, i)
    Next j
   Next i
  End If
    
End Sub

