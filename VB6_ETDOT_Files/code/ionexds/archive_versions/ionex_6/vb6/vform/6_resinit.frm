VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmResinPresaturantConditions 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Initial Resin Phase Presaturant Conditions"
   ClientHeight    =   5820
   ClientLeft      =   2010
   ClientTop       =   1500
   ClientWidth     =   7365
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
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Begin Threed.SSFrame fraPresaturant 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   8705
      _StockProps     =   14
      Caption         =   "Percentage of Each Ion Presaturating the Resin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5340
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5340
         TabIndex        =   11
         Top             =   1140
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   5340
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   5340
         TabIndex        =   9
         Top             =   1980
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   5340
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   5340
         TabIndex        =   7
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   5340
         TabIndex        =   6
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   5340
         TabIndex        =   5
         Top             =   3660
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   5340
         TabIndex        =   4
         Top             =   4020
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   5340
         TabIndex        =   3
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   24
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ion"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5340
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   1140
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   19
         Top             =   1980
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   18
         Top             =   2400
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   17
         Top             =   2820
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   16
         Top             =   3240
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   15
         Top             =   3660
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   14
         Top             =   4020
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   13
         Top             =   4440
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   5220
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   5220
      Width           =   1095
   End
End
Attribute VB_Name = "frmResinPresaturantConditions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Base 1
Option Explicit

Dim NumIonsSelected  As Integer
Dim Temp_Text As String

Dim Percentages(1 To MAX_CHEMICAL) As Double

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer

    If Cations.Available And Anions.Available Then

    ElseIf Cations.Available Then
       For i = 1 To NumSelectedCations
           j = Cations_Selected(i)
           NowProj.Resin.PresaturantPercentage(j) = Percentages(i)
       Next i
    ElseIf Anions.Available Then
       For i = 1 To NumSelectedAnions
           j = Anions_Selected(i)
           NowProj.Resin.PresaturantPercentage(j) = Percentages(i)
       Next i
    End If

    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer, j As Integer

    For i = 0 To 9
        lblIon(i).visible = False
        txtPercentage(i).visible = False
    Next i
    
    If Cations.Available And Anions.Available Then

    ElseIf Cations.Available Then
       
       For i = 1 To NumSelectedCations
           lblIon(i - 1).visible = True
           txtPercentage(i - 1).visible = True
           j = Cations_Selected(i)
           lblIon(i - 1) = NowProj.cation(j).Name
           Percentages(i) = NowProj.Resin.PresaturantPercentage(j)
           txtPercentage(i - 1) = Format$(Percentages(i), "0.00")
       Next i

       NumIonsSelected = NumSelectedCations

    ElseIf Anions.Available Then
       For i = 1 To NumSelectedAnions
           lblIon(i - 1).visible = True
           txtPercentage(i - 1).visible = True
           j = Anions_Selected(i)
           lblIon(i - 1) = NowProj.Anion(j).Name
           Percentages(i) = NowProj.Resin.PresaturantPercentage(j)
           txtPercentage(i - 1) = Format$(Percentages(i), "0.00")
       Next i

       NumIonsSelected = NumSelectedAnions

    End If

       fraPresaturant.height = lblIon(NumIonsSelected - 1).top + lblIon(NumIonsSelected - 1).height + 120
       cmdOK.top = fraPresaturant.top + fraPresaturant.height + 180
       cmdCancel.top = cmdOK.top
       frmResinPresaturantConditions.height = cmdOK.top + cmdOK.height + 540

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.width - frmResinPresaturantConditions.width) / 2, (Screen.height - frmResinPresaturantConditions.height) / 2
    End If
       
End Sub

Private Sub txtPercentage_GotFocus(index As Integer)

    If index = 0 Then Exit Sub

    Call TextGetFocus(txtPercentage(index), Temp_Text)
End Sub

Private Sub txtPercentage_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtPercentage_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer, SumPercentages As Double
    Dim IsError As Integer

    If index = 0 Then Exit Sub

    Call TextHandleError(IsError, txtPercentage(index), Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtPercentage(index).Text)
       OldValue = Percentages(index + 1)

       If (NewValue > 100#) Or (NewValue < 0#) Then
          txtPercentage(index) = Temp_Text
          Exit Sub
       End If

       SumPercentages = NewValue
       For i = 2 To NumIonsSelected
           If i <> (index + 1) Then SumPercentages = SumPercentages + Percentages(i)
       Next i

       If SumPercentages > 100# Then
          txtPercentage(index) = Temp_Text
          Exit Sub
       End If

       Percentages(index + 1) = NewValue
       Percentages(1) = 100# - SumPercentages
       txtPercentage(0) = Format$(Percentages(1), "0.00")
       
    End If

End Sub

