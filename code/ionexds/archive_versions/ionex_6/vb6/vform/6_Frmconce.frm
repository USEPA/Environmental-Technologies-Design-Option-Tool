VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmConcentrations 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Influent Concentrations"
   ClientHeight    =   5385
   ClientLeft      =   1590
   ClientTop       =   2205
   ClientWidth     =   7755
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5385
   ScaleWidth      =   7755
   Begin VCIF1Lib.F1Book Grid1 
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "6_Frmconce.frx":0000
      TabIndex        =   9
      Top             =   480
      Width           =   5535
   End
   Begin Threed.SSCommand cmdEdit 
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   0
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Paste"
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Cancel"
   End
   Begin Threed.SSFrame frame3D1 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   1931
      _StockProps     =   14
      Caption         =   "Variable Influent File Name:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdInputFile 
         Appearance      =   0  'Flat
         Caption         =   "Input New File Name"
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   660
         Width           =   3075
      End
      Begin VB.Label lblInputFile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   6720
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   2520
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   972
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "OK"
   End
   Begin Threed.SSCommand cmdEdit 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Copy"
   End
   Begin Threed.SSCommand cmdEdit 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Cut"
   End
   Begin Threed.SSCommand cmdEdit 
      Height          =   375
      Index           =   4
      Left            =   3600
      TabIndex        =   10
      Top             =   0
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Insert"
   End
   Begin Threed.SSCommand cmdEdit 
      Height          =   375
      Index           =   5
      Left            =   4680
      TabIndex        =   11
      Top             =   0
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Delete"
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
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Paste "
         Index           =   0
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmConcentrations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1

Dim Shifting As Integer, x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
Dim TempStr As String, SaveAs  As Integer, Filename_Concentration As String
Dim Temp_Array() As String
Dim OriginalVarInfluentFile As String


Private Sub cmdEdit_Click(Index As Integer)
Dim i As Integer, j As Integer

  On Error GoTo ErrReset
  
   Select Case Index
     Case 0
       frmConcentrations.Grid1.EditCut
     
     Case 1
       frmConcentrations.Grid1.EditCopy
     
    Case 2
       frmConcentrations.Grid1.EditPasteValues    '' Only paste values
     
    Case 4
      Dim n1R As Integer
      Dim n2R As Integer
      Dim n1C As Integer
      Dim n2C As Integer
      
      n1R = frmConcentrations.Grid1.SelStartRow
      n2R = frmConcentrations.Grid1.SelEndRow
      n1C = frmConcentrations.Grid1.SelStartCol
      n2C = frmConcentrations.Grid1.SelEndCol
      
     frmConcentrations.Grid1.InsertRange n1R, n1C, n2R, n2C, F1ShiftRows
      
    Case 5
      Dim n1RD As Integer
      Dim n2RD As Integer
      Dim n1CD As Integer
      Dim n2CD As Integer
      
      n1RD = frmConcentrations.Grid1.SelStartRow
      n2RD = frmConcentrations.Grid1.SelEndRow
      n1CD = frmConcentrations.Grid1.SelStartCol
      n2CD = frmConcentrations.Grid1.SelEndCol
      
     frmConcentrations.Grid1.DeleteRange n1RD, n1CD, n2RD, n2CD, F1ShiftRows
      
   End Select
   
   Exit Sub
   
ErrReset:
  Call Show_Error("Data in clipboard is not valid.")
  Exit Sub
End Sub

Private Sub cmdCancel_Click()
  If Cations.Available And Anions.Available Then
  ElseIf Cations.Available Then
     NowProj.VarInfluentFileCation = OriginalVarInfluentFile
  ElseIf Anions.Available Then

  End If

  Unload Me
End Sub

Private Sub cmdInputFile_Click()
    Dim OldVarInfFile As String

    If Cations.Available And Anions.Available Then

    ElseIf Cations.Available Then
       OldVarInfFile = NowProj.VarInfluentFileCation
       frmConcentrations!CMDialog1.filename = NowProj.VarInfluentFileCation
       Call SaveFileVariableInfluent(NowProj.VarInfluentFileCation)

    ElseIf Anions.Available Then

    End If

    If NowProj.VarInfluentFileCation = "" Then
       NowProj.VarInfluentFileCation = OldVarInfFile
    Else
       Call PutInputFileName
    End If

End Sub

Private Sub cmdOK_Click()
Dim i As Integer, response As Integer
ReDim NDATA(7) As Integer
Dim DFlag As Integer, F As Integer
Dim j As Integer, No_Var_Influent As Integer

  No_Var_Influent = False
  If Not (CountConc(1, Number_Influent_Points)) Then Exit Sub

  Grid1.Row = 1
  For i = 1 To Total_NumberOfComponents
   Grid1.Col = i
   If Grid1.Text = "" Then No_Var_Influent = True
  Next i
  If No_Var_Influent Then
   response = MsgBox("There is no data for the first row" & Chr$(13) & "It will be assummed there is no variable influent concentration", MB_ICONSTOP + MB_OKCANCEL, App.title)
  End If
  Select Case response
   Case IDOK
     GoTo NoInfluent_Conc
   Case IDCANCEL
     Exit Sub
  End Select

  For j = 1 To Total_NumberOfComponents + 1
    Grid1.Col = j
    NDATA(j) = 0
    Grid1.Row = 0
    Do Until Grid1.Text = "" Or Grid1.Row >= Number_Max_Influent_Points
     Grid1.Row = Grid1.Row + 1
     NDATA(j) = NDATA(j) + 1
    Loop
  Next j
  DFlag = False
  For i = 1 To Total_NumberOfComponents + 1
   For j = i + 1 To Total_NumberOfComponents + 1
    If NDATA(i) <> NDATA(j) Then DFlag = True
   Next j
  Next i
  If DFlag Then response = MsgBox("There is not the same number of data in each column." & Chr$(13) & "It will be assummed that there is no variable influent concentrations", MB_ICONEXCLAMATION + MB_OKCANCEL, App.title)
  Select Case response
     Case IDOK
     Case IDCANCEL
      Exit Sub
  End Select
  
  'Store Time values
On Error GoTo Time_Error
  Grid1.Col = 1
  For i = 1 To Number_Influent_Points
   Grid1.Row = i
    T_Influent(i) = CDbl(Grid1.Text) * 24# * 60#  ' To convert from days to minutes
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
     If NowProj.VarInfluentFileCation = "NONE" Then
        Call SaveFileVariableInfluent(NowProj.VarInfluentFileCation)
     End If
     
     If NowProj.VarInfluentFileCation <> "" Then
        Call SaveVariableInfluent(NowProj.VarInfluentFileCation)
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
 MsgBox "At least one value in time input (" & Format$(i, "0") & ") is not a real number." & Chr$(13) & "Return to the editing mode.", MB_ICONEXCLAMATION, App.title
 Resume Exit_This_OK:

Time_Error2:
 MsgBox "Time(" & Format$(i, "0") & ") < (=) Time(" & Format$(i - 1, "0") & ")" & Chr$(13) & "Return to editing mode", MB_ICONEXCLAMATION + MB_OK, App.title
 Exit Sub

Conc_Error:
 MsgBox "At least one value in concentration (" & Format$(i, "0") & "," & Format$(j, "0") & ") is not a real number." & Chr$(13) & "Return to the editing mode.", MB_ICONEXCLAMATION, App.title
 Resume Exit_This_OK:

FileWritingError:
 Resume SpecifyVarInfluentCation

NoInfluent_Conc:
   Number_Influent_Points = 0
   Unload Me
End Sub

Private Function CountConc(i As Integer, NPoints As Integer) As Integer

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
  MsgBox "Invalid data.", 64, App.title
  Resume Exit_CountConc
Exit_CountConc:
End Function

Private Function CutString() As Integer
Dim ClipString As String, length As Integer
Dim CurrentPosition As Integer, PreviousPosition As Integer, Character As String * 1
Dim StringToTransfer As String, Row As Integer, Col As Integer

On Error GoTo Error_In_CutString

  ClipString = Clipboard.GetText()
  length = Len(ClipString)

  If length > 0 Then
    PreviousPosition = 1: CurrentPosition = 1
    Col = 1: Row = 1
    While PreviousPosition <= length
      Character = Mid$(ClipString, CurrentPosition, 1)
      Select Case Asc(Character)
        Case 10
          CurrentPosition = CurrentPosition + 1
          PreviousPosition = CurrentPosition
        Case 13, 9
          StringToTransfer = Mid$(ClipString, PreviousPosition, CurrentPosition - PreviousPosition)
          If Not (PasteString(StringToTransfer, Row, Col)) Then
            MsgBox "Error while pasting data.", 64, App.title
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
  MsgBox "Too much data was selected. Only the first " & Format$(Number_Max_Influent_Points, "0") & " points were pasted.", 64, App.title
  GoTo Exit_CutString
Error_In_CutString:
  CutString = False
  Resume Exit_CutString
Exit_CutString:
End Function

Private Sub Form_Activate()

    Call PutInputFileName

End Sub

Private Sub Form_Load()

' Me.HelpContextID = Hlp_Influent_Concentrations
  cmdEdit(0).Enabled = True   'cut
  cmdEdit(1).Enabled = True   'copy
  cmdEdit(2).Enabled = True   'paste could be from Excel, right?
  'removed Fill Down because Formula One does that by pulling down the handle
  cmdEdit(4).Enabled = True   'insert
  cmdEdit(5).Enabled = True   'delete

  Grid1.AllowFillRange = True   'fill down by dragging handle on cell
  If Cations.Available And Anions.Available Then
  ElseIf Cations.Available Then
     OriginalVarInfluentFile = NowProj.VarInfluentFileCation
  ElseIf Anions.Available Then

  End If
  Grid1.SheetName(1) = "  "
  Call UpdateVarInf

  top = Screen.height / 2 - height / 2
  left = Screen.width / 2 - width / 2

End Sub

Private Sub Form_Resize()
  If WindowState = 0 Then
    Grid1.width = width - 350
  End If
  If WindowState = 1 Then frmIonExchangeMain.WindowState = 1
    'If WindowState <> 0 Then Exit Sub
    'If (Grid1.Left + Grid1.Width) > (cmdEdit(3).Left + cmdEdit(3).Width) Then
    '  Width = Grid1.Left + Grid1.Width + 20 * Screen.TwipsPerPixelX
    'Else
    '  Width = cmdEdit(3).Left + cmdEdit(3).Width + 20 * Screen.TwipsPerPixelX
    'End If
    'If Height > (Grid1.Top + 90 * Screen.TwipsPerPixelY + cmdCancel.Height) Then
    ' Grid1.Height = Height - Grid1.Top - 90 * Screen.TwipsPerPixelY - cmdCancel.Height
    ' cmdCancel.Top = Grid1.Top + Grid1.Height + 15 * Screen.TwipsPerPixelY
    ' cmdOK.Top = Grid1.Top + Grid1.Height + 15 * Screen.TwipsPerPixelY
    'End If

    'Top = Screen.Height / 2 - Height / 2
    'Left = Screen.Width / 2 - Width / 2
End Sub



'Private Sub Grid1_DoDblClick()
'  If Grid1.Row > 0 And Grid1.Col > 0 Then Grid1_KeyPress (13)
'End Sub
'
'Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
'  txtCel.visible = False
'  If Shift = 1 Then Shifting = True Else Shifting = False
'
'End Sub

'Private Sub Grid1_KeyPress(KeyAscii As Integer)
'Dim Char As String
'  Select Case KeyAscii
'    Case 27
'      txtCel.Text = Grid1.Text
'    Case 9
'      If Shifting Then
'        If Grid1.Col > 1 Then
'          Grid1.Col = Grid1.Col - 1
'        End If
'      Else
'        If Grid1.Col < (Grid1.Cols - 1) Then
'         Grid1.Col = Grid1.Col + 1
'        End If
'      End If
'      Unselect
'    Case Else
'     If KeyAscii = 13 Then
'       txtCel = Grid1.Text
'       txtCel.SelStart = Len(txtCel.Text)
'     Else
'      Char = Chr$(KeyAscii)
'      txtCel = Char
'      txtCel.SelStart = 1
'     End If
'     ShowTextBox
'     KeyAscii = 0
'  End Select
'End Sub

'Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  txtCel.visible = False
'End Sub
'
'Private Sub Grid1_RowColChange()
'  txtCel.Text = Grid1.Text
'End Sub

Private Sub mnuEditItem_Click(Index As Integer)
  Select Case Index
    Case 0
      If CutString() Then
      Else
       MsgBox "Impossible to paste data from the clipboard.", 64, App.title
      End If
  End Select
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
Dim i As Integer, j As Integer, F As Integer
Dim OldVarInfFile As String, n1C As Integer, n1R As Integer
  Select Case Index
    Case 0
      n1C = 0
      j = 0
      While j <= Total_NumberOfComponents
        n1C = n1C + 1
        n1R = 0
        i = 0
        While i <= Number_Max_Influent_Points - 1
          n1R = n1R + 1
          Grid1.EntryRC(n1R, n1C) = ""
          i = i + 1
        Wend
        j = j + 1
      Wend
    Case 1  'Open Variable Influent File
       If Cations.Available And Anions.Available Then
       ElseIf Cations.Available Then
          OldVarInfFile = NowProj.VarInfluentFileCation
          Call OpenFileVariableInfluent(NowProj.VarInfluentFileCation)
          If NowProj.VarInfluentFileCation = "" Then
             NowProj.VarInfluentFileCation = OldVarInfFile
          Else
             Call ReadVarInfluentConcs
             Call PutInputFileName
             Call UpdateVarInf
          End If
        End If
    Case 2   'Save
       Call SaveVariableInfluent(NowProj.VarInfluentFileCation)
       Call PutInputFileName
    Case 3   'Save As
      If Cations.Available And Anions.Available Then
      ElseIf Cations.Available Then
         OldVarInfFile = NowProj.VarInfluentFileCation
         Call SaveFileVariableInfluent(NowProj.VarInfluentFileCation)
         If NowProj.VarInfluentFileCation = "" Then
            NowProj.VarInfluentFileCation = OldVarInfFile
         Else
            Call SaveVariableInfluent(NowProj.VarInfluentFileCation)
            Call PutInputFileName
         End If
      ElseIf Anions.Available Then
      End If
      
      
  End Select
End Sub

Private Sub OpenFileVariableInfluent(VarInfluentFileName As String)

       On Error Resume Next
       frmConcentrations!CMDialog1.filename = ""
       frmConcentrations!CMDialog1.DefaultExt = "var"
       frmConcentrations!CMDialog1.Filter = "Variable Influent (*.var)|*.var"
       frmConcentrations!CMDialog1.DialogTitle = "Ion Exchange Variable Influent Files"
       frmConcentrations!CMDialog1.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
'------Begin Modification Hokanson: 12-Aug2000
       frmConcentrations!CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
       frmConcentrations!CMDialog1.Action = 1
       VarInfluentFileName$ = frmConcentrations!CMDialog1.filename
       If Err = 32755 Then   'Cancel selected by user
          VarInfluentFileName$ = ""
       End If

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

Private Sub PutInputFileName()
    Dim RightFile As String
    Dim TheLength As Integer
    Dim i As Integer
    Dim Done As Integer

    Done = False
    If Cations.Available And Anions.Available Then

    ElseIf Cations.Available Then
       
       For i = Len(NowProj.VarInfluentFileCation) To 1 Step -1
           If Mid$(NowProj.VarInfluentFileCation, i, 1) = "\" Then
              TheLength = Len(NowProj.VarInfluentFileCation) - i
              RightFile = Right$(NowProj.VarInfluentFileCation, TheLength)
              Done = True
              Exit For
           End If
       Next i

       If Not Done Then
          RightFile = NowProj.VarInfluentFileCation
       End If

       lblInputFile = RightFile
       
    ElseIf Anions.Available Then

    End If

End Sub

Private Sub ShowTextBox()
  Dim TestX As Integer, TestY As Integer
  Dim c As Integer
  txtCel.visible = False
  txtCel.height = Grid1.RowHeight(Grid1.Row) - (Screen.TwipsPerPixelY * 2)
  txtCel.width = Grid1.ColWidth(Grid1.Col) - (Screen.TwipsPerPixelX * 2)

  TestX = Grid1.left + Grid1.ColWidth(0) + (Screen.TwipsPerPixelX * 3)
  
  For c = Grid1.LeftCol To Grid1.Col - 1
    TestX = TestX + Grid1.ColWidth(c) + Screen.TwipsPerPixelX
  Next c

  TestY = Grid1.top + Grid1.RowHeight(0) + (Screen.TwipsPerPixelY * 2)
  For c = Grid1.TopRow To Grid1.Row - 1
    TestY = TestY + Grid1.RowHeight(c) + Screen.TwipsPerPixelY
  Next c

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

Private Sub UpdateVarInf()
Dim i As Integer, j As Integer, TB As String, CB As String
Dim temp As String, LF  As String, c As Integer, SetWidth As Integer

  frmConcentrations.Grid1.MaxCol = Total_NumberOfComponents + 1
  frmConcentrations.Grid1.MaxRow = Number_Max_Influent_Points
  frmConcentrations.Grid1.HdrWidth = 3000
  frmConcentrations.Grid1.ColText(1) = "Time (days)"
  
  For i = 1 To Total_NumberOfComponents
       frmConcentrations.Grid1.ColText(i + 1) = _
           Trim$(Ion(i).Name) & " (mg/L)"
   Next i

  For c = 1 To frmConcentrations.Grid1.MaxCol
   Grid1.ColWidth(c) = 300 * Screen.TwipsPerPixelX
  Next c
  SetWidth = Screen.TwipsPerPixelX * 19
  For c = 1 To frmConcentrations.Grid1.MaxCol
    SetWidth = SetWidth + Grid1.ColWidth(c) + (Screen.TwipsPerPixelX * 2)
  Next c

  If Number_Influent_Points > 0 Then
   For i = 1 To Number_Influent_Points
      frmConcentrations.Grid1.EntryRC(i, 2) = T_Influent(i) / 24# / 60#     'Convert form min. to days
    For j = 2 To Total_NumberOfComponents + 1
      frmConcentrations.Grid1.EntryRC(i, j) = C_Influent(j - 1, i)
    Next j
   Next i
  End If
    
End Sub

