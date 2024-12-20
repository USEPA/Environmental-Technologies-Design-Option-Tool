VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlantData 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Compare prediction to data"
   ClientHeight    =   5220
   ClientLeft      =   1260
   ClientTop       =   1845
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Plantdat.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5220
   ScaleWidth      =   7350
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   5040
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox Grid1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   0
      ScaleHeight     =   5205
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open..."
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As..."
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   8
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Paste from Excel"
         Index           =   0
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuCompare 
      Caption         =   "&Compare Data"
      Begin VB.Menu mnuCompareItem 
         Caption         =   "To &PFPSDM Results"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmPlantData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1

Dim Shifting As Integer
Dim TempStr As String
Dim Filename_Data As String
Dim SaveAs As Integer

Private Function CountPoints(i As Integer, NPoints As Integer) As Integer

On Error GoTo Error_In_CountPoints

  CountPoints = False
  NPoints = 0
  Grid1.Col = i
  Grid1.Row = 1
  Do Until Grid1.Text = "" Or Grid1.Row = Number_Data_Points_Max
    NPoints = NPoints + 1
    Grid1.Row = Grid1.Row + 1
  Loop
  If Grid1.Text <> "" Then NPoints = NPoints + 1
  CountPoints = True
  Exit Function

Error_In_CountPoints:
  NPoints = 0
  CountPoints = False
  Resume Exit_CountPoints
Exit_CountPoints:
End Function

Private Function CutString() As Integer
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
          Col = Col Mod (NComponents + 1) + 1
          If Col = 1 Then
           Row = Row + 1
           If Row > Number_Data_Points_Max Then GoTo Too_Many_Points
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
  MsgBox "Too much data was selected. Only the first " & Format$(Number_Data_Points_Max, "0") & " points were pasted.", 64, Application_Name
  GoTo Exit_CutString
Error_In_CutString:
  CutString = False
  Resume Exit_CutString
Exit_CutString:
End Function

Private Sub Form_Load()
Dim i As Integer, j As Integer, TB As String, CB As String
Dim temp As String, LF  As String, C As Integer, SetWidth As Integer
Dim TimeUnits As Integer, factor As Double

  top = Screen.height / 2 - height / 2
  left = Screen.width / 2 - width / 2

'  Me.HelpContextID = Hlp_Compare_prediction
  TB = Chr$(9)
  CB = Chr$(13)
  NComponents = Results.NComponent
  Grid1.Cols = 1
  Grid1.Rows = 1

  TimeUnits = TimeUnitsOnGraphs
  If TimeUnits = 0 Then  'min
     temp = TB & "Time (min) "
     factor = 1#
  ElseIf TimeUnits = 1 Then   'sec
     temp = TB & "Time (sec) "
     factor = 1# * 60#
  ElseIf TimeUnits = 2 Then   'hrs
     temp = TB & "Time (hrs) "
     factor = 1# / 60#
  ElseIf TimeUnits = 3 Then   'days
     temp = TB & "Time (days)"
     factor = 1# / 60# / 24#
  End If

  For i = 1 To NComponents
    temp = temp & TB & Trim$(Results.Component(i).Name)
  Next i
  Grid1.AddItem temp
  
  For i = 0 To NComponents + 1
   Grid1.FixedAlignment(i) = 2
  Next i
  
  For i = 1 To Number_Data_Points_Max
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
  Grid1.width = SetWidth

  'Width = Grid1.Width + 8 * Screen.TwipsPerPixelX

  top = Screen.height / 2 - height / 2
  left = Screen.width / 2 - width / 2

  mnuFileItem(2).Enabled = False

  'Display the current data available in memory
  If (Results.NComponent = NComponents) And (NData_Points > 0) Then
    Grid1.Col = 1
    Grid1.Row = 0
    For i = 1 To NData_Points
     Grid1.Row = Grid1.Row + 1
     Grid1.Text = Format$(T_Data_Points(i) * factor, "0.000E+00")
    Next i
    For j = 1 To NComponents
      Grid1.Col = j + 1
      Grid1.Row = 0
      For i = 1 To NData_Points
       Grid1.Row = Grid1.Row + 1
       Grid1.Text = Format$(C_Data_Points(j, i), "0.000E+00")
       
      Next i
    Next j
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Response As Integer, i As Integer, j As Integer
    
  If ReadPoints() Then
  Else
    Response = MsgBox("The points could not be read correctly. Do you still want to keep the points in memory?", MB_ICONSTOP + MB_YESNO, Application_Name)
    Select Case Response
      Case IDNO
        NData_Points = 0
        NComponents = 0
      Case IDYES
        Cancel = 1 'Do not unload
    End Select
  End If
End Sub

Private Sub Form_Resize()
  If WindowState = 0 Then
   Grid1.width = width - 8 * Screen.TwipsPerPixelX
   Grid1.height = height - 45 * Screen.TwipsPerPixelY
  End If
  If WindowState = 1 Then frmIonExchangeMain.WindowState = 1
End Sub

Private Sub Grid1_DblClick()
  If Grid1.Row > 0 And Grid1.Col > 0 Then Grid1_KeyPress (13)
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
  txtCel.visible = False
  If Shift = 1 Then Shifting = True Else Shifting = False
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
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

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtCel.visible = False
End Sub

Private Sub Grid1_RowColChange()
   txtCel.Text = Grid1.Text
End Sub

Private Sub LoadPoints()
Dim F As Integer, NPoints As Integer, i As Integer, j  As Integer
ReDim T(Number_Data_Points_Max) As Double, C(6, Number_Data_Points_Max) As Double
    F = FreeFile
On Error GoTo Error_In_Reading:
    CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
    CMDialog1.FilterIndex = 2
'------Begin Modification Hokanson: 12-Aug2000
    CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
    CMDialog1.Action = 1
   If CMDialog1.FileName = "" Then
     Exit Sub
   End If
   Filename_Data = CMDialog1.FileName
   mnuFileItem(2).Enabled = True
    Open Filename_Data For Input As F
    Input #F, NPoints
    For i = 1 To NPoints
      Select Case NComponents
        Case 1
          Input #F, T(i), C(1, i)
        Case 2
          Input #F, T(i), C(1, i), C(2, i)
        Case 3
          Input #F, T(i), C(1, i), C(2, i), C(3, i)
        Case 4
          Input #F, T(i), C(1, i), C(2, i), C(3, i), C(4, i)
        Case 5
          Input #F, T(i), C(1, i), C(2, i), C(3, i), C(4, i), C(5, i)
        Case 6
          Input #F, T(i), C(1, i), C(2, i), C(3, i), C(4, i), C(5, i), C(6, i)
      End Select
    Next i
    Close (F)
    Grid1.Row = 0
    Grid1.Col = 1
    For i = 1 To NPoints
      Grid1.Row = Grid1.Row + 1
      Grid1.Text = T(i)
    Next i

    Grid1.Col = 1
    For j = 1 To NComponents
      Grid1.Col = Grid1.Col + 1
      Grid1.Row = 0
      For i = 1 To NPoints
        Grid1.Row = Grid1.Row + 1
        Grid1.Text = C(j, i)
        'Grid1.Row = Grid1.Row + 1
      Next i
    Next j
    Exit Sub

Error_In_Reading:
'------Begin Modification Hokanson: 12-Aug2000
    If Err = 32755 Then   'Cancel selected by user
       Resume Exit_Load_Points
    End If
'------End Modification Hokanson: 11-Aug2000
    Dim temp As String, Error_Code As Integer
    temp = "Error " & Format$(Error_Code, "0") & " : " & Error$(Error_Code)
    Close (F)
    MsgBox "An error occured while reading the file." & Chr$(13) & temp, MB_ICONEXCLAMATION, Application_Name
    Resume Exit_Load_Points
Exit_Load_Points:
End Sub

Private Sub mnuCompareItem_Click(Index As Integer)
  Select Case Index
    Case 0 'PFPSDM Results
      If ReadPoints() Then
        If (NData_Points > 0) And (NComponents > 0) Then
         frmShow_Data_And_Prediction.Show 1
        Else
         MsgBox "The data available does not allow you to compare the prediction to the data.", 64, Application_Name
      End If
    Else
         MsgBox "The data available does not allow you to compare the prediction to the data.", 64, Application_Name

    End If

  End Select
End Sub

Private Sub mnuEditItem_Click(Index As Integer)
  Select Case Index
    Case 0 '
      If CutString() Then
      Else
       MsgBox "Impossible to paste data from the clipboard.", 64, Application_Name
      End If
  End Select
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
Dim i As Integer, j As Integer
  Select Case Index
    Case 0
      Grid1.Col = 0
      j = 0
      While j <= Results.NComponent
        Grid1.Col = Grid1.Col + 1
        Grid1.Row = 0
        i = 0
        While i <= Number_Data_Points_Max - 1
          Grid1.Row = Grid1.Row + 1
          Grid1.Text = ""
          i = i + 1
        Wend
        j = j + 1
      Wend
    Case 1
       LoadPoints
    Case 2
       SaveAs = False
       If Not (SavePoints()) Then Exit Sub
    Case 3
       SaveAs = True
       If Not (SavePoints()) Then Exit Sub
    Case 8
       Unload Me
  End Select
End Sub

Private Function PasteString(StringToTransfer As String, Row As Integer, Col As Integer) As Integer

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

Private Function ReadPoints() As Integer
Dim NPoints As Integer, i As Integer, j As Integer, temp As String, Error_Code As Integer
Dim factor As Double

  ReadPoints = False
  If CountPoints(1, NPoints) Then
    NData_Points = NPoints
  Else
    ReadPoints = False
    NData_Points = 0
    Exit Function
  End If
  Grid1.Col = 1
  Grid1.Row = 0

On Error GoTo Conversion_ErrorI

    If TimeUnitsOnGraphs = 0 Then   'min
       factor = 1#
    ElseIf TimeUnitsOnGraphs = 1 Then   'sec to min
       factor = 1# / 60#
    ElseIf TimeUnitsOnGraphs = 2 Then   'hrs to min
       factor = 1# * 60#
    ElseIf TimeUnitsOnGraphs = 3 Then   'days to min
       factor = 1# * 60# * 24#
    End If

  For i = 1 To NPoints
    Grid1.Row = Grid1.Row + 1
    T_Data_Points(i) = CDbl(Grid1.Text) * factor
    For j = 1 To NComponents
     Grid1.Col = Grid1.Col + 1
     C_Data_Points(j, i) = CDbl(Grid1.Text)
    Next j
    Grid1.Col = 1
  Next i

  ReadPoints = True
  Exit Function

Conversion_ErrorI:
  ReadPoints = False
  NData_Points = 0
  Error_Code = Err
  temp = "Error " & Format$(Error_Code, "0") & " : " & Error$(Error_Code)
  MsgBox "Error in data." & Chr$(13) & temp, MB_ICONEXCLAMATION, Application_Name
  Resume Exit_Read_Points
Exit_Read_Points:
End Function

Private Function SavePoints() As Integer
Dim F As Integer, NPoints As Integer, i As Integer, j As Integer
Dim Stemp As String, temp As String, Error_Code As Integer
Dim TemporaryName  As String, Previous_FileName_Data As String

  SavePoints = False
On Error GoTo Error_In_SavePoints
  If Not (CountPoints(1, NPoints)) Then
    MsgBox "No Data has been saved.", 64, Application_Name
    SavePoints = False
    Exit Function
  End If

  If (Trim$(Filename_Data) <> "") And Not (SaveAs) Then GoTo Save_File
  Previous_FileName_Data = Filename_Data
  CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
  CMDialog1.FilterIndex = 2
'------Begin Modification Hokanson: 12-Aug2000
  CMDialog1.CancelError = True
'------End Modification Hokanson: 12-Aug2000
  CMDialog1.Action = 2
  TemporaryName = CMDialog1.FileName
  If IsValidPath(TemporaryName, "C:") And CMDialog1.FileName <> "" Then
    TemporaryName = Mid$(TemporaryName, 1, Len(TemporaryName) - 1)
    Filename_Data = TemporaryName
  Else
    Filename_Data = Previous_FileName_Data
    CMDialog1.FileName = ""
    MsgBox "No data has been saved.", 64, Application_Name
    Exit Function
  End If

Save_File:
   F = FreeFile
   Grid1.Col = 1
   Grid1.Row = 0
   Open Filename_Data For Output As F
   Print #F, Format$(NPoints, "0")
   For i = 1 To NPoints
     Grid1.Row = Grid1.Row + 1
     Stemp = Format$(CDbl(Grid1.Text), "0.0000E+00")
     For j = 1 To NComponents
      Grid1.Col = Grid1.Col + 1
      Stemp = Stemp & "," & Format$(CDbl(Grid1.Text), "0.0000E+00")
     Next j
     Print #F, Stemp
     Grid1.Col = 1
   Next i
   Close (F)
   mnuFileItem(2).Enabled = True
   SavePoints = True
   Exit Function

Error_In_SavePoints:
   
'------Begin Modification Hokanson: 12-Aug2000
    If Err = 32755 Then   'Cancel selected by user
       Resume Exit_SavePoints
    End If
'------End Modification Hokanson: 12-Aug2000

   SavePoints = False
   If Err = 13 Then
     Close (F)
     MsgBox "The data entered are not valid data.", MB_ICONEXCLAMATION, Application_Name
   Else
     Close (F)
     Error_Code = Err
     temp = "Error" & Format$(Error_Code, "0") & " : " & Error$(Error_Code) & "."
     MsgBox "Error while saving the data." & Chr$(13) & temp, MB_ICONEXCLAMATION, Application_Name
   End If
   Resume Exit_SavePoints
Exit_SavePoints:
End Function

Private Sub ShowTextBox()
  Dim TestX As Integer, TestY As Integer
  Dim C As Integer
  txtCel.visible = False
  txtCel.height = Grid1.RowHeight(Grid1.Row) - (Screen.TwipsPerPixelY * 2)
  txtCel.width = Grid1.ColWidth(Grid1.Col) - (Screen.TwipsPerPixelX * 2)

  TestX = Grid1.left + Grid1.ColWidth(0) + (Screen.TwipsPerPixelX * 3)
  
  For C = Grid1.LeftCol To Grid1.Col - 1
    TestX = TestX + Grid1.ColWidth(C) + Screen.TwipsPerPixelX
  Next C

  TestY = Grid1.top + Grid1.RowHeight(0) + (Screen.TwipsPerPixelY * 2)
  For C = Grid1.TopRow To Grid1.Row - 1
    TestY = TestY + Grid1.RowHeight(C) + Screen.TwipsPerPixelY
  Next C

  txtCel.left = TestX
  txtCel.top = TestY

  txtCel.ZOrder
  txtCel.visible = True
  txtCel.SetFocus

End Sub

Private Sub txtCel_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 Then
     Shifting = True
   Else
     Shifting = False
   End If
   Select Case KeyCode
   Case 38
    txtCel_KeyPress (13)
    SendKeys "{UP}"
   Case 40
    txtCel_KeyPress (13)
    SendKeys "{DOWN}"
   End Select

End Sub

Private Sub txtCel_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13, 9
     Grid1.Text = txtCel.Text
     txtCel.visible = False
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
     txtCel.visible = False
     Grid1.SetFocus
   End Select


End Sub

Private Sub Unselect()
  If Grid1.visible = False Then Exit Sub
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

