VERSION 5.00
Begin VB.Form frmunits 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Unit"
   ClientHeight    =   3390
   ClientLeft      =   5145
   ClientTop       =   2265
   ClientWidth     =   2130
   ClipControls    =   0   'False
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
   Icon            =   "frmunits.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3390
   ScaleWidth      =   2130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2910
      Width           =   855
   End
   Begin VB.ComboBox CMBUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   2715
      Left            =   0
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "CMBUnits"
      Top             =   0
      Width           =   2140
   End
End
Attribute VB_Name = "frmunits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CMBUnits_DblClick()
    
    Dim ConvertFrom As String
    
    ' if it's called by the preference form
    If SetDefaultUnit = True Then
        If CurProp = OptTemp Then
            DefaultTFTUnit = CMBUnits.List(CMBUnits.ListIndex)
            FRMPreferences!GRDDefaultUnits.Col = 1
            FRMPreferences!GRDDefaultUnits.Text = DefaultTFTUnit
        Else
            DefaultUnit(CurProp) = CMBUnits.List(CMBUnits.ListIndex)
            FRMPreferences!GRDDefaultUnits.Col = 1
            FRMPreferences!GRDDefaultUnits.Text = DefaultUnit(CurProp)
        End If
        Unload FRMUnits
        Exit Sub
    ' else if it's
    ElseIf TFTConvert = True Then
        ConvertFrom = DispMethod(CurProp).TFTUnit
        Call ConvertTFTUnits(ConvertFrom, CMBUnits.List(CMBUnits.ListIndex))
        DefaultTFTUnit = CMBUnits.List(CMBUnits.ListIndex)
        Unload FRMUnits
        Call LoadFTInfo(Forms(1), DispMethod(CurProp).CurMethod)
        Exit Sub
    ElseIf GraphConvert = True Then
        FRMGraphSet!TXTMinT.Text = FormatVal(Convert(Val(FRMGraphSet!TXTMinT.Text), OptPress, FRMGraphSet!LBLTFTUnit(0).caption, CMBUnits.List(CMBUnits.ListIndex), True)) 'Do not understand so I put in dummy value -2
        FRMGraphSet!TXTMaxT.Text = FormatVal(Convert(Val(FRMGraphSet!TXTMaxT.Text), OptPress, FRMGraphSet!LBLTFTUnit(1).caption, CMBUnits.List(CMBUnits.ListIndex), True)) 'Do not understand so I put in dummy value -2
        FRMGraphSet!LBLTFTUnit(0).caption = CMBUnits.List(CMBUnits.ListIndex)
        FRMGraphSet!LBLTFTUnit(1).caption = CMBUnits.List(CMBUnits.ListIndex)
        FRMGraphSet!TXTXAxis.Text = "T (" & CMBUnits.List(CMBUnits.ListIndex) & ")"
        Unload FRMUnits
        GraphConvert = False
        Exit Sub
    ElseIf CurProp = OptTemp Then
        ConvertFrom = Cur_Disp.OpTUnit
    ElseIf CurProp = OptPress Then
        ConvertFrom = Cur_Disp.OpPUnit
    Else
        ConvertFrom = DispMethod(CurProp).Unit
    End If
    
    If CurProp <> OptTemp And CurProp <> OptPress Then
        DefaultUnit(CurProp) = CMBUnits.List(CMBUnits.ListIndex)
    End If
    
    Call ConvertUnits(ConvertFrom, CMBUnits.List(CMBUnits.ListIndex))
    Call DisplayOneProp(CurProp)
    
    Unload FRMUnits

End Sub


Private Sub CMBUnits_KeyPress(KeyAscii As Integer)

    Dim ConvertFrom As String
    
    If SetDefaultUnit = True Then
        If CurProp = OptTemp Then
            DefaultTFTUnit = CMBUnits.List(CMBUnits.ListIndex)
            FRMPreferences!GRDDefaultUnits.Col = 1
            FRMPreferences!GRDDefaultUnits.Text = DefaultTFTUnit
        Else
            DefaultUnit(CurProp) = CMBUnits.List(CMBUnits.ListIndex)
            FRMPreferences!GRDDefaultUnits.Col = 1
            FRMPreferences!GRDDefaultUnits.Text = DefaultUnit(CurProp)
        End If
        Unload FRMUnits
        Exit Sub
    ElseIf TFTConvert = True Then
        ConvertFrom = DispMethod(CurProp).TFTUnit
        Call ConvertTFTUnits(ConvertFrom, CMBUnits.List(CMBUnits.ListIndex))
        DefaultTFTUnit = CMBUnits.List(CMBUnits.ListIndex)
        Unload FRMUnits
        Call LoadFTInfo(Forms(1), DispMethod(CurProp).CurMethod)
        Exit Sub
    ElseIf GraphConvert = True Then
        FRMGraphSet!TXTMinT.Text = FormatVal(Convert(Val(FRMGraphSet!TXTMinT.Text), OptPress, FRMGraphSet!LBLTFTUnit(0).caption, CMBUnits.List(CMBUnits.ListIndex), True)) 'FIX : -2 is dummy number
        FRMGraphSet!TXTMaxT.Text = FormatVal(Convert(Val(FRMGraphSet!TXTMaxT.Text), OptPress, FRMGraphSet!LBLTFTUnit(1).caption, CMBUnits.List(CMBUnits.ListIndex), True)) 'FIX : -2 is dummy number
        FRMGraphSet!LBLTFTUnit(0).caption = CMBUnits.List(CMBUnits.ListIndex)
        FRMGraphSet!LBLTFTUnit(1).caption = CMBUnits.List(CMBUnits.ListIndex)
        FRMGraphSet!TXTXAxis.Text = "T (" & CMBUnits.List(CMBUnits.ListIndex) & ")"
        Unload FRMUnits
        GraphConvert = False
        Exit Sub
    ElseIf CurProp = OptTemp Then
        ConvertFrom = Cur_Disp.OpTUnit
    ElseIf CurProp = OptPress Then
        ConvertFrom = Cur_Disp.OpPUnit
    Else
        ConvertFrom = DispMethod(CurProp).Unit
    End If
    
    If CurProp <> OptTemp And CurProp <> OptPress Then
        DefaultUnit(CurProp) = CMBUnits.List(CMBUnits.ListIndex)
    End If
    
    Call ConvertUnits(ConvertFrom, CMBUnits.List(CMBUnits.ListIndex))
    Call DisplayOneProp(CurProp)
    
    Unload FRMUnits
    
End Sub


Private Sub CMDCan_Click()

    FRMUnits.Hide
    
End Sub

Private Sub cmdok_Click()
                   
    Dim ConvertFrom As String
    
    ' if it's called from preference form
    If SetDefaultUnit = True Then
        If CurProp = OptTemp Then
            DefaultTFTUnit = CMBUnits.List(CMBUnits.ListIndex)
            FRMPreferences!GRDDefaultUnits.Col = 1
            FRMPreferences!GRDDefaultUnits.Text = DefaultTFTUnit
        Else
            DefaultUnit(CurProp) = CMBUnits.List(CMBUnits.ListIndex)
            FRMPreferences!GRDDefaultUnits.Col = 1
            FRMPreferences!GRDDefaultUnits.Text = DefaultUnit(CurProp)
        End If
        Unload FRMUnits
        Exit Sub
    ' else if it's converting a f(t) unit
    ElseIf TFTConvert = True Then
        ConvertFrom = DispMethod(CurProp).TFTUnit
        Call ConvertTFTUnits(ConvertFrom, CMBUnits.List(CMBUnits.ListIndex))
        DefaultTFTUnit = CMBUnits.List(CMBUnits.ListIndex)
        Unload FRMUnits
        Call LoadFTInfo(Forms(ScreenNum), DispMethod(CurProp).CurMethod)
        Exit Sub
    ' else if it's converting for graphing
    ElseIf GraphConvert = True Then
        FRMGraphSet!TXTMinT.Text = FormatVal(Convert(Val(FRMGraphSet!TXTMinT.Text), OptTemp, FRMGraphSet!LBLTFTUnit(0).caption, CMBUnits.List(CMBUnits.ListIndex), True))
        FRMGraphSet!TXTMaxT.Text = FormatVal(Convert(Val(FRMGraphSet!TXTMaxT.Text), OptTemp, FRMGraphSet!LBLTFTUnit(1).caption, CMBUnits.List(CMBUnits.ListIndex), True))
        FRMGraphSet!LBLTFTUnit(0).caption = CMBUnits.List(CMBUnits.ListIndex)
        FRMGraphSet!LBLTFTUnit(1).caption = CMBUnits.List(CMBUnits.ListIndex)
        FRMGraphSet!TXTXAxis.Text = "T (" & CMBUnits.List(CMBUnits.ListIndex) & ")"
        Unload FRMUnits
        GraphConvert = False
        Exit Sub
    ElseIf CurProp = OptTemp Then
        ConvertFrom = Cur_Disp.OpTUnit
    ElseIf CurProp = OptPress Then
        ConvertFrom = Cur_Disp.OpPUnit
    Else
        ConvertFrom = DispMethod(CurProp).Unit
    End If
    
    If CurProp <> OptTemp And CurProp <> OptPress Then
        DefaultUnit(CurProp) = CMBUnits.List(CMBUnits.ListIndex)
    End If
    
    Call ConvertUnits(ConvertFrom, CMBUnits.List(CMBUnits.ListIndex))
    Call DisplayOneProp(CurProp)
        
    Unload FRMUnits
        
End Sub







Private Sub Form_Load()
   
    CenterForm Me
     
End Sub


