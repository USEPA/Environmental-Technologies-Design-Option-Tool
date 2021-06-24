VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmmethod 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "FRMMethod"
   ClientHeight    =   6255
   ClientLeft      =   630
   ClientTop       =   1320
   ClientWidth     =   8280
   ControlBox      =   0   'False
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
   Icon            =   "frmmetho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6255
   ScaleWidth      =   8280
   Begin MSFlexGridLib.MSFlexGrid GRDDataSources 
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1720
      _Version        =   65541
      Cols            =   4
   End
   Begin VB.ComboBox CMBPropInputs 
      Height          =   315
      Left            =   4830
      TabIndex        =   6
      Text            =   "CMBPropInputs"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox TXTMethodInfo 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmmetho.frx":030A
      Top             =   2400
      Width           =   4575
   End
   Begin VB.CommandButton CMDClose 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   5820
      Width           =   1335
   End
   Begin VB.CommandButton CMDAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5820
      Width           =   1335
   End
   Begin VB.ComboBox CMBBIPSel 
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Text            =   "CMBBIPSel"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.PictureBox FRMFT 
      Height          =   3615
      Left            =   4800
      ScaleHeight     =   3555
      ScaleWidth      =   3315
      TabIndex        =   8
      Top             =   2160
      Width           =   3375
      Begin VB.TextBox TXTCoeffE 
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Text            =   "TXTCoeffE"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox TXTCoeffD 
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Text            =   "TXTCoeffD"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox TXTCoeffC 
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "TXTCoeffC"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TXTCoeffB 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Text            =   "TXTCoeffB"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox TXTCoeffA 
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Text            =   "TXTCoeffA"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox TXTMaxT 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Text            =   "TXTMaxT"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TXTMinT 
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Text            =   "TXTMinT"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TXTEqnForm 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Text            =   "TXTEqnForm"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TXTCorrT 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Text            =   "TXTCorrT"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label LBLTUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "LBLTUnits"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   30
         Top             =   1470
         Width           =   255
      End
      Begin VB.Label LBLTUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "LBLTUnits"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   29
         Top             =   1110
         Width           =   255
      End
      Begin VB.Label LBLCoeffE 
         Caption         =   "Coefficient E"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3270
         Width           =   1335
      End
      Begin VB.Label LBLCoeffD 
         Caption         =   "Coefficient D"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2910
         Width           =   1335
      End
      Begin VB.Label LBLCoeffC 
         Caption         =   "Coefficient C"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2550
         Width           =   1335
      End
      Begin VB.Label LBLCoeffB 
         Caption         =   "Coefficient B"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2190
         Width           =   1335
      End
      Begin VB.Label LBLCoeffA 
         Caption         =   "Coefficient A"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Label LBLMaxT 
         Caption         =   "Maximum T"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label LBLMinT 
         Caption         =   "Minimum T"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label LBLEqnForm 
         Caption         =   "Equation Form"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label LBLTUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "LBLTUnits"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   11
         Top             =   390
         Width           =   255
      End
      Begin VB.Label LBLCorrT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Correlation T"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   390
         Width           =   1335
      End
   End
   Begin VB.Label LBLMethodInfo 
      Caption         =   "Method Information"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label LBLPropInputs 
      Caption         =   "Property Links for Predictive Method Inputs"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label LBLBIPSel 
      Caption         =   "UNIFAC Binary Interaction Parameter Database"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Label LBLPredictiveMethods 
      Alignment       =   2  'Center
      Caption         =   "Data Sources Available "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmmethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CMBBIPSel_Click()
    
    Dim i As Integer
    Dim TempValue As Double
    Dim TempIndex As Integer
'    Dim TempOpTVal As Double
    
    'Convert operating T to to K for calculations
    ' done in subroutines now
'    TempOpTVal = Cur_Info.OpT
'    If Trim(Cur_Info.OpTUnit) <> ConvertToDefault(-1) Then
'        Cur_Info.OpT = simple_t_convert(Trim(Cur_Info.OpTUnit), "K", Cur_Info.OpT)
'    End If

    Select Case CurProp
        Case 32
            TempIndex = CMBBIPSel.ListIndex + 1
            TempValue = CalcACwaterUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, TempIndex, Get_DefaultUnit(ACwater))
'msh        TempValue = CalcACwaterUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, TempIndex, ConvertToDefault(ACwater))
        Case 34
            TempIndex = CMBBIPSel.ListIndex + 1
'msh        TempValue = CalcACchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, TempIndex, ConvertToDefault(ACchem))
            TempValue = CalcACchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, TempIndex, Get_DefaultUnit(ACchem))
        Case 35
            TempIndex = CMBBIPSel.ListIndex + 1
'msh        TempValue = CalclogKowUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, TempIndex, ConvertToDefault(logKow))
            TempValue = CalclogKowUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, TempIndex, Get_DefaultUnit(logKow))
        Case 39
            TempIndex = CMBBIPSel.ListIndex + 1
            TempValue = CalcSchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), TempIndex, DefaultUnit(39))
        Case 40
            TempIndex = CMBBIPSel.ListIndex + 1
            TempValue = CalcSwaterUNIFAC(Cur_Info.OpT, InfoMethod(MW).value(InfoMethod(MW).CurMethod), TempIndex, DefaultUnit(40))
    End Select
        
    'Update current selected method
    Me!GRDDataSources.Col = 1
    For i = 1 To Me!GRDDataSources.Rows - 1
        Me!GRDDataSources.Row = i
        If Trim(Me!GRDDataSources.Text) = "UNIFAC" Then
            Me!GRDDataSources.Col = 2
            Me!GRDDataSources.Text = FormatVal(TempValue)
        End If
    Next i
            
    'Set operating T back to proper value
  '  Cur_Info.OpT = TempOpTVal
    
End Sub


Private Sub CMBPropInputs_Click()
    
    If CMBPropInputs.Text <> "None" Then
        CurProp = GetPropCode(Me!CMBPropInputs.Text)
        Call CreateMethodForm
    End If
    
End Sub

Private Sub CMDAccept_Click()
    
    Dim i As Integer
    Dim J As Integer
    Dim whichscreen As Integer
    Dim CurRow As Integer
    Dim temp_value As Double

    Screen.MousePointer = 11
                              
    'Set the modified flag
    WorkModified = True
    FRMMain.caption = "PEARLS:  " & SaveFileName & " modified"
    
    'Set the current method index
    Me!GRDDataSources.Col = 0
    For J = 1 To Me!GRDDataSources.Rows
        Me!GRDDataSources.Row = J
        If Trim(Me!GRDDataSources.Text) = "X" Then
            Me!GRDDataSources.Col = 1
            For i = 1 To NumMethods
                If Trim(Me!GRDDataSources.Text) = Trim(InfoMethod(CurProp).MethodName(i)) Then
                    InfoMethod(CurProp).CurMethod = i
                    If Trim(Me!GRDDataSources.Text) = "UNIFAC" Then
                        Select Case CurProp
                            Case 32
                                BIPIndex(1) = CMBBIPSel.ListIndex + 1
                                temp_value = CalcACwaterUNIFAC(Cur_Info.OpT, "K", BIPIndex(1), Get_DefaultUnit(Swater))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(32).value(4) = temp_value
                                    InfoMethod(32).Enabled(4) = True
                                End If
                                temp_value = CalcACchemUNIFAC(Cur_Info.OpT, "K", BIPIndex(1), Get_DefaultUnit(Swater))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(34).value(5) = temp_value
                                    InfoMethod(34).Enabled(5) = True
                                End If
                            Case 34
                                BIPIndex(1) = CMBBIPSel.ListIndex + 1
                                
                                temp_value = CalcACchemUNIFAC(Cur_Info.OpT, "K", BIPIndex(1), Get_DefaultUnit(Swater))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(34).value(5) = temp_value
                                    InfoMethod(34).Enabled(5) = True
                                End If
                                temp_value = CalcACwaterUNIFAC(Cur_Info.OpT, "K", BIPIndex(1), Get_DefaultUnit(Swater))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(32).value(4) = temp_value
                                    InfoMethod(32).Enabled(4) = True
                                End If
                            Case 35
                                BIPIndex(2) = CMBBIPSel.ListIndex + 1
                                temp_value = CalclogKowUNIFAC(Cur_Info.OpT, "K", BIPIndex(2), Get_DefaultUnit(Swater))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(35).value(4) = temp_value
                                    InfoMethod(35).Enabled(4) = True
                                End If
                            Case 39
                                BIPIndex(3) = CMBBIPSel.ListIndex + 1
                                temp_value = CalcSchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPIndex(3), DefaultUnit(39))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(39).value(7) = temp_value
                                    InfoMethod(39).Enabled(7) = True
                                End If
                                temp_value = CalcSwaterUNIFAC(Cur_Info.OpT, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPIndex(3), DefaultUnit(40))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(40).value(4) = temp_value
                                    InfoMethod(40).Enabled(4) = True
                                End If
                            Case 40
                                BIPIndex(3) = CMBBIPSel.ListIndex + 1
                                temp_value = CalcSwaterUNIFAC(Cur_Info.OpT, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPIndex(3), DefaultUnit(40))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(40).value(4) = temp_value
                                    InfoMethod(40).Enabled(4) = True
                                End If
                                temp_value = CalcSchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPIndex(3), DefaultUnit(39))
                                If temp_value <> 0# And temp_value <> ERROR_FLAG Then
                                    InfoMethod(39).value(7) = temp_value
                                    InfoMethod(39).Enabled(7) = True
                                End If
                        End Select
                    End If
                End If
            Next i
            GoTo NextStep:
        End If
    Next J
    
NextStep:
    
    'Set TFT temperature if changed
    If Me!TXTCorrT.Enabled = True Then
        InfoMethod(CurProp).TFT = Me!TXTCorrT.Text
    End If
    
    'Calculate f(T) values
    For i = 1 To NumMethods
        If InfoMethod(CurProp).EqNum(i) <> 0 Then
            InfoMethod(CurProp).value(i) = CalcFofT(CurProp, i)
        End If
    Next i
    
    'Check user input value if selected
    If InfoMethod(CurProp).CurMethod = 10 And InfoMethod(CurProp).value(10) = 0 Then
        MsgBox "Please specify a valid value for user input", 48, "No Value"
        Screen.MousePointer = 1
        Exit Sub
    End If
        
    'Recalculate properties
    Call RecalcPred
        
    'Update values
    Call DisplayProps
    
    Unload Me
    
    ScreenNum = ScreenNum - 1
        
   ' CurProp = PrevProp(ScreenNum)
    
    'Update previous screens if any (need to make sure it's a method form)
    For whichscreen = 0 To Forms.count - 1
        If Left(Trim(Forms(whichscreen).caption), 9) = "Property:" Then
            CurProp = get_prop_from_caption(Forms(whichscreen).caption)
    
   ' If ScreenNum > 0 Then
            CurRow = 1
            If Left(Trim(Forms(whichscreen).caption), 9) = "Property:" Then
                Forms(whichscreen)!GRDDataSources.Col = 2
                For i = 1 To NumMethods
                    If InfoMethod(CurProp).Enabled(i) = True Then
                        Forms(whichscreen)!GRDDataSources.Row = CurRow
                        Forms(whichscreen)!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(i))
                        CurRow = CurRow + 1
                    End If
                Next i
            End If
        End If
    Next whichscreen
    Screen.MousePointer = 1
    
End Sub



Private Sub CMDClose_Click()

    Dim ConvertFrom As String
    
    Unload Me

    ScreenNum = ScreenNum - 1
    
    'Reset onriginal TFT and unit
    ConvertFrom = Trim(InfoMethod(CurProp).TFTUnit)
    Call ConvertTFTUnits(ConvertFrom, Trim(TempTFTUnit))
    DefaultTFTUnit = TempTFTUnit

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = &H70 Or KeyAscii = 43 Or KeyAscii = 46 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 69 Or KeyAscii = 101 Or KeyAscii = &H25 Or KeyAscii = &H27 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub Form_Load()
 
    CenterForm Me
                   
End Sub



















Private Sub GRDDataSources_Click()

    Dim i As Integer
    Dim TempMethod As Integer
    Dim RowSelected As Integer
    Dim ColSelected As Integer
    
    RowSelected = Me!GRDDataSources.Row
    ColSelected = Me!GRDDataSources.Col
            
    'Set the current method index
    Me!GRDDataSources.Col = 1
    For i = 1 To NumMethods
        'here is where it figures what method was selected and stores in TempMethod
        If Trim(Me!GRDDataSources.Text) = Trim(InfoMethod(CurProp).MethodName(i)) Then TempMethod = i
    Next i
        
    'Check to see if this method can be selected
    If CurProp = logKow And InfoMethod(logKow).CurMethod = 5 And InfoMethod(Schem).CurMethod = 5 Then
        MsgBox "Yalkowsky correlation being used for solubility in water", 48, "Method Not Available"
        Exit Sub
    End If
    If CurProp = logKow And InfoMethod(logKow).CurMethod = 6 And InfoMethod(Schem).CurMethod = 5 Then
        MsgBox "Yalkowsky method being used for solubility in water", 48, "Method Not Available"
        Exit Sub
    End If
    If CurProp = Schem And InfoMethod(Schem).CurMethod = 6 And InfoMethod(logKow).CurMethod = 5 Then
        MsgBox "Kenaga and Goring method being used for log Kow", 48, "Method Not Available"
        Exit Sub
    End If
    If CurProp = Schem And InfoMethod(Schem).CurMethod = 6 And InfoMethod(logKow).CurMethod = 6 Then
        MsgBox "Hansch method being used for log Kow", 48, "Method Not Available"
        Exit Sub
    End If
    
    'Deselect the current selection
    Me!GRDDataSources.Col = 0
    For i = 1 To Me!GRDDataSources.Rows - 1
        Me!GRDDataSources.Row = i
        If Me!GRDDataSources.Text = "X" Then Me!GRDDataSources.Text = ""
    Next i
    
    'Select the new selection
    Me!GRDDataSources.Row = RowSelected
    Me!GRDDataSources.Col = 0
    Me!GRDDataSources.Text = "X"
        
    ' Sorry, some special cases (for now it's just antoine) added 3/30/98 by DMW
    If CurProp = VP And TempMethod = 4 Then
        Call load_frm_antoine
        FRMAntoine.Show 1
        
        Dim X
        Me!GRDDataSources.Col = 1
        For i = 1 To Me!GRDDataSources.Rows - 1
            Me!GRDDataSources.Row = i
            If Me!GRDDataSources.Text = "Antoine" Then
                Me!GRDDataSources.Col = 2
                Me!GRDDataSources.Text = InfoMethod(6).value(4)
                
                Me!GRDDataSources.Col = 3
                Me!GRDDataSources.Text = InfoMethod(6).Unit
                Exit For
            End If
        Next

    End If
    
    'Update screen
    If InfoMethod(CurProp).EqNum(TempMethod) <> 0 Then
        Call EnableFofT(Me)
        Call LoadFTInfo(Me, TempMethod)
    Else
        Call DisableFofT(Me)
    End If
        
    Me!TXTMethodInfo.Text = RefText(InfoMethod(CurProp).EqNum(TempMethod), InfoMethod(CurProp).MethodName(TempMethod))
    Call LoadPropertyInputs(Me, InfoMethod(CurProp).MethodName(TempMethod))
    
    'Set BIP database for properties using UNIFAC
    Me!GRDDataSources.Col = 1
    If Trim(Me!GRDDataSources.Text) = "UNIFAC" Then
        Select Case CurProp
            Case ACchem
                Me!LBLBIPSel.Enabled = True
                Me!CMBBIPSel.Enabled = True
                Me!CMBBIPSel.Text = Me!CMBBIPSel.List(BIPIndex(1) - 1)
            Case ACwater
                Me!LBLBIPSel.Enabled = True
                Me!CMBBIPSel.Enabled = True
                Me!CMBBIPSel.Text = Me!CMBBIPSel.List(BIPIndex(1) - 1)
            Case logKow
                Me!LBLBIPSel.Enabled = True
                Me!CMBBIPSel.Enabled = True
                Me!CMBBIPSel.Text = Me!CMBBIPSel.List(BIPIndex(2) - 1)
            Case Schem
                Me!LBLBIPSel.Enabled = True
                Me!CMBBIPSel.Enabled = True
                Me!CMBBIPSel.Text = Me!CMBBIPSel.List(BIPIndex(3) - 1)
            Case Swater
                Me!LBLBIPSel.Enabled = True
                Me!CMBBIPSel.Enabled = True
                Me!CMBBIPSel.Text = Me!CMBBIPSel.List(BIPIndex(3) - 1)
        End Select
    Else
        Me!LBLBIPSel.Enabled = False
        Me!CMBBIPSel.Enabled = False
    End If

End Sub

Private Sub GRDDataSources_DblClick()
   
    Dim RowSelected As Integer
    Dim ColSelected As Integer
    Dim dbtype As Integer
    Dim which_property As Integer
    Dim CurrentUnits As String
    
    which_property = GetPropCode(Right(Me.caption, Len(Me.caption) - 10))
    ' first make sure it's not an f(t) property
    'If is_f_of_t(which_property) Then
    '    Exit Sub
    'End If
    RowSelected = Me!GRDDataSources.Row
    ColSelected = Me!GRDDataSources.Col
 
    Me!GRDDataSources.Col = 1
    
    If Trim(Me!GRDDataSources.Text) = "911 Database" Then
        dbtype = 9
    Else
        If Trim(Me!GRDDataSources.Text) = "801 Database" Then
            dbtype = 8
        Else
            dbtype = 0
        End If
    End If
        
    GRDDataSources.Col = 3
    CurrentUnits = Trim$(GRDDataSources.Text)
    
    'Call Create911DBInfoForm(GetPropCode(Right(Me.caption, Len(Me.caption) - 10)))
    If dbtype = 8 Or dbtype = 9 Then
        Call Create911DBInfoForm(which_property, dbtype, CurrentUnits)
    End If
    
        
        
End Sub


Private Sub GRDDataSources_KeyPress(KeyAscii As Integer)

    Dim TempStr As String
    
    On Error GoTo NotValid
    
    If KeyAscii = 13 Then
        Me!GRDDataSources.Col = 1
        If Trim(Me!GRDDataSources.Text) = "User Input" Then
            Me!GRDDataSources.Col = 2
            InfoMethod(CurProp).value(10) = Val(Me!GRDDataSources.Text)
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            InfoMethod(CurProp).Enabled(10) = True
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf KeyAscii = 8 Then
        Me!GRDDataSources.Col = 1
        If Trim(Me!GRDDataSources.Text) = "User Input" Then
            Me!GRDDataSources.Col = 2
            TempStr = Me!GRDDataSources.Text
            If TempStr <> "" Then
                Me!GRDDataSources.Text = Mid(TempStr, 1, Len(Trim(TempStr)) - 1)
                InfoMethod(CurProp).value(10) = Val(Me!GRDDataSources.Text)
                InfoMethod(CurProp).Enabled(10) = True
            End If
            Exit Sub
        End If
    Else
        Me!GRDDataSources.Col = 1
        If Trim(Me!GRDDataSources.Text) = "User Input" Then
            Me!GRDDataSources.Col = 2
            Me!GRDDataSources.Text = Me!GRDDataSources.Text + Chr(KeyAscii)
            InfoMethod(CurProp).value(10) = Val(Me!GRDDataSources.Text)
            InfoMethod(CurProp).Enabled(10) = True
        End If
    End If

    Exit Sub
    
NotValid:
    MsgBox "Not a valid number, please enter a smaller value", 0, "Value Not Valid"
    Me!GRDDataSources.Col = 1
    If Trim(Me!GRDDataSources.Text) = "User Input" Then
        Me!GRDDataSources.Col = 2
        TempStr = Me!GRDDataSources.Text
        If TempStr <> "" Then
            Me!GRDDataSources.Text = Mid(TempStr, 1, Len(Trim(TempStr)) - 1)
            InfoMethod(CurProp).value(10) = Val(Me!GRDDataSources.Text)
            InfoMethod(CurProp).Enabled(10) = True
        End If
        Exit Sub
    End If
    
End Sub


Private Sub GRDDataSources_LostFocus()
    
    Dim Row As Integer
    
    For Row = 1 To Me!GRDDataSources.Rows - 1
        Me!GRDDataSources.Row = Row
        Me!GRDDataSources.Col = 1
        If Trim(Me!GRDDataSources.Text) = "User Input" Then
            Me!GRDDataSources.Col = 2
            InfoMethod(CurProp).value(10) = Val(Me!GRDDataSources.Text)
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            InfoMethod(CurProp).Enabled(10) = True
        End If
    Next Row

End Sub

Private Sub LBLTUnits_Click(Index As Integer)

    Dim i As Integer
       
    TFTConvert = True
    
    FRMUnits!CMBUnits.Clear

'msh FRMUnits!CMBUnits.AddItem ConvertToDefault(-1)
    FRMUnits!CMBUnits.AddItem Get_DefaultUnit(-1)

    i = 1
    Do While Unit1(i) <> "End"
'msh    If Trim(Unit1(i)) = Trim(ConvertToDefault(-1)) Then
        If Trim(Unit1(i)) = Trim(Get_DefaultUnit(-1)) Then
            FRMUnits!CMBUnits.AddItem Unit2(i)
'msh    ElseIf Trim(Unit2(i)) = Trim(ConvertToDefault(CurProp)) Then
        ElseIf Trim(Unit2(i)) = Trim(Get_DefaultUnit(CurProp)) Then
            FRMUnits!CMBUnits.AddItem Unit1(i)
        End If
            
        i = i + 1
    Loop
    
    FRMUnits!CMBUnits.Text = InfoMethod(CurProp).TFTUnit

    FRMUnits.Show 1

End Sub

Private Sub TXTCoeffA_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    
    If KeyAscii = 13 Then
        InfoMethod(CurProp).Coeff(10, 1) = Val(TXTCoeffA.Text)
        InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
        Me!GRDDataSources.Col = 1
        For i = 1 To GRDDataSources.Rows - 1
            If Me!GRDDataSources.Text = "User Input" Then
                Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            End If
        Next i
    End If

End Sub


Private Sub TXTCoeffA_LostFocus()

    Dim i As Integer
    
    InfoMethod(CurProp).Coeff(10, 1) = Val(TXTCoeffA.Text)
    InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
    Me!GRDDataSources.Col = 1
    For i = 1 To GRDDataSources.Rows - 1
        If Me!GRDDataSources.Text = "User Input" Then
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
        End If
    Next i

End Sub


Private Sub TXTCoeffB_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    
    If KeyAscii = 13 Then
        InfoMethod(CurProp).Coeff(10, 2) = Val(TXTCoeffB.Text)
        InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
        Me!GRDDataSources.Col = 1
        For i = 1 To GRDDataSources.Rows - 1
            If Me!GRDDataSources.Text = "User Input" Then
                Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            End If
        Next i
    End If
    
End Sub


Private Sub TXTCoeffB_LostFocus()

    Dim i As Integer
    
    InfoMethod(CurProp).Coeff(10, 2) = Val(TXTCoeffB.Text)
    InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
    Me!GRDDataSources.Col = 1
    For i = 1 To GRDDataSources.Rows - 1
        If Me!GRDDataSources.Text = "User Input" Then
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
        End If
    Next i

End Sub


Private Sub TXTCoeffC_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    
    If KeyAscii = 13 Then
        InfoMethod(CurProp).Coeff(10, 3) = Val(TXTCoeffC.Text)
        InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
        Me!GRDDataSources.Col = 1
        For i = 1 To GRDDataSources.Rows - 1
            If Me!GRDDataSources.Text = "User Input" Then
                Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            End If
        Next i
    End If

End Sub


Private Sub TXTCoeffC_LostFocus()

    Dim i As Integer
    
    InfoMethod(CurProp).Coeff(10, 3) = Val(TXTCoeffC.Text)
    InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
    Me!GRDDataSources.Col = 1
    For i = 1 To GRDDataSources.Rows - 1
        If Me!GRDDataSources.Text = "User Input" Then
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
        End If
    Next i

End Sub


Private Sub TXTCoeffD_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    
    If KeyAscii = 13 Then
        InfoMethod(CurProp).Coeff(10, 4) = Val(TXTCoeffD.Text)
        InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
        Me!GRDDataSources.Col = 1
        For i = 1 To GRDDataSources.Rows - 1
            If Me!GRDDataSources.Text = "User Input" Then
                Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            End If
        Next i
    End If

End Sub


Private Sub TXTCoeffD_LostFocus()

    Dim i As Integer
    
    InfoMethod(CurProp).Coeff(10, 4) = Val(TXTCoeffD.Text)
    InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
    Me!GRDDataSources.Col = 1
    For i = 1 To GRDDataSources.Rows - 1
        If Me!GRDDataSources.Text = "User Input" Then
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
        End If
    Next i

End Sub


Private Sub TXTCoeffE_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    
    If KeyAscii = 13 Then
        InfoMethod(CurProp).Coeff(10, 5) = Val(TXTCoeffE.Text)
        InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
        Me!GRDDataSources.Col = 1
        For i = 1 To GRDDataSources.Rows - 1
            If Me!GRDDataSources.Text = "User Input" Then
                Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            End If
        Next i
    End If

End Sub


Private Sub TXTCoeffE_LostFocus()

    Dim i As Integer
    
    InfoMethod(CurProp).Coeff(10, 5) = Val(TXTCoeffE.Text)
    InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
    Me!GRDDataSources.Col = 1
    For i = 1 To GRDDataSources.Rows - 1
        If Me!GRDDataSources.Text = "User Input" Then
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
        End If
    Next i

End Sub


Private Sub TXTCorrT_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    Dim CurRow As Integer
    Dim TempValue(NumMethods) As Double
    
    If KeyAscii = 13 Then
        InfoMethod(CurProp).TFT = Val(TXTCorrT.Text)
        For i = 1 To NumMethods
            If InfoMethod(CurProp).EqNum(i) <> 0 Then
                TempValue(i) = CalcFofT(CurProp, i)
            End If
        Next i
        Me!GRDDataSources.Col = 2
        CurRow = 1
        For i = 1 To NumMethods
            If InfoMethod(CurProp).Enabled(i) = True Then
                GRDDataSources.Row = CurRow
                Me!GRDDataSources.Row = CurRow
                Me!GRDDataSources.Text = FormatVal(TempValue(i))
                CurRow = CurRow + 1
            End If
        Next i
    End If
            
End Sub


Private Sub TXTCorrT_LostFocus()

    Dim i As Integer
    Dim CurRow As Integer
    Dim TempValue(NumMethods) As Double
    
    InfoMethod(CurProp).TFT = Val(TXTCorrT.Text)
    For i = 1 To NumMethods
        If InfoMethod(CurProp).EqNum(i) <> 0 Then
            TempValue(i) = CalcFofT(CurProp, i)
        End If
    Next i
    Me!GRDDataSources.Col = 2
    CurRow = 1
    For i = 1 To NumMethods
        If InfoMethod(CurProp).Enabled(i) = True Then
            GRDDataSources.Row = CurRow
            Me!GRDDataSources.Row = CurRow
            Me!GRDDataSources.Text = FormatVal(TempValue(i))
            CurRow = CurRow + 1
        End If
    Next i

End Sub


Private Sub TXTEqnForm_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyCode = 0
    
End Sub

Private Sub TXTEqnForm_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub


Private Sub TXTEqnForm_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyCode = 0
    
End Sub

Private Sub TXTMaxT_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    
    If KeyAscii = 13 Then
        InfoMethod(CurProp).MaxT(10) = Val(TXTMaxT.Text)
        InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
        Me!GRDDataSources.Col = 1
        For i = 1 To GRDDataSources.Rows - 1
            If Me!GRDDataSources.Text = "User Input" Then
                Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            End If
        Next i
    End If
      
End Sub


Private Sub TXTMaxT_LostFocus()

    Dim i As Integer
    
    InfoMethod(CurProp).MaxT(10) = Val(TXTMaxT.Text)
    InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
    Me!GRDDataSources.Col = 1
    For i = 1 To GRDDataSources.Rows - 1
        If Me!GRDDataSources.Text = "User Input" Then
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
        End If
    Next i

End Sub


Private Sub TXTMethodInfo_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub

Private Sub TXTMethodInfo_KeyPress(KeyAscii As Integer)
    
    KeyAscii = 0

End Sub

Private Sub TXTMethodInfo_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyCode = 0
    
End Sub


Private Sub TXTMinT_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    
    If KeyAscii = 13 Then
        InfoMethod(CurProp).MinT(10) = Val(TXTMinT.Text)
        InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
        Me!GRDDataSources.Col = 1
        For i = 1 To GRDDataSources.Rows - 1
            If Me!GRDDataSources.Text = "User Input" Then
                Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
            End If
        Next i
    End If

End Sub


Private Sub TXTMinT_LostFocus()

    Dim i As Integer
    
    InfoMethod(CurProp).MinT(10) = Val(TXTMinT.Text)
    InfoMethod(CurProp).value(10) = CalcFofT(CurProp, 10)
    Me!GRDDataSources.Col = 1
    For i = 1 To GRDDataSources.Rows - 1
        If Me!GRDDataSources.Text = "User Input" Then
            Me!GRDDataSources.Text = FormatVal(InfoMethod(CurProp).value(10))
        End If
    Next i

End Sub



Private Sub accept_block_5()
' NOT USED????
'Dim DBTbl As Recordset
'Dim property_name As String
'Dim J As Integer
'Dim I As Integer
'Dim prop_index As Integer
'Dim K As Integer
'Dim temp As Integer
'Dim prevtemp As Integer

' first check if it's a block 5 property
' first figure out which row is selected
    'For J = 1 To Me!GRDDataSources.Rows - 1
           ' Me!GRDDataSources.Row = J
           ' Me!GRDDataSources.Col = 0
           ' If Trim(Me!GRDDataSources.Text) = "X" Then
           '     Me!GRDDataSources.Col = 1
           '     Exit For
            'End If
   ' Next J
         
   ' Select Case CurProp
     '   Case UFL
       '     prop_index = 1
      '      property_name = "UFL"
      '  Case LFL
      '      prop_index = 2
      '      property_name = "LFL"
      '  Case FP
      '      prop_index = 3
     '       property_name = "FP"
      '  Case AIT
      '      prop_index = 4
      '      property_name = "AIT"
  '  End Select
    
   ' For K = 1 To 7
   '     If Trim(Me!GRDDataSources.Text) = Trim(infomethod(CurProp).MethodName(K)) Then
   '         temp = K
   '         Exit For
   '     End If
  '  Next K
            ' this needs temp to have been found, which means all method names need to be in there
  '  prevtemp = B5Preference(prop_index, 1)
  '  B5Preference(prop_index, 1) = temp
    'temp = prevtemp
   ' For K = 2 To 6
   '     If prevtemp <> B5Hierarchy(prop_index, 1) Then
  '          temp = B5Hierarchy(prop_index, K)
   '         B5Hierarchy(prop_index, K) = prevtemp
   '         prevtemp = temp
   '     Else
   '         Exit For
   '     End If
   'Next K
End Sub
