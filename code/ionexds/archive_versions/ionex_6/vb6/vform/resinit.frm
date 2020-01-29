VERSION 5.00
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
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   495
      Left            =   6000
      TabIndex        =   24
      Top             =   5220
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   5220
      Width           =   1095
   End
   Begin VB.PictureBox fraPresaturant 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   4815
      Left            =   360
      ScaleHeight     =   4785
      ScaleWidth      =   6645
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   5160
         TabIndex        =   22
         Top             =   4380
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   5160
         TabIndex        =   20
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   5160
         TabIndex        =   18
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   5160
         TabIndex        =   16
         Top             =   3180
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   5160
         TabIndex        =   14
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   5160
         TabIndex        =   12
         Top             =   2340
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   5160
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   8
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5160
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   3
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   300
         TabIndex        =   21
         Top             =   4380
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   300
         TabIndex        =   19
         Top             =   3960
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   17
         Top             =   3600
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   15
         Top             =   3180
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   300
         TabIndex        =   13
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   11
         Top             =   2340
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   9
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   1500
         Width           =   4695
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ion"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   300
         Width           =   4095
      End
      Begin VB.Label lblIon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   660
         Width           =   4695
      End
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
           Resin.PresaturantPercentage(j) = Percentages(i)
       Next i
    ElseIf Anions.Available Then
       For i = 1 To NumSelectedAnions
           j = Anions_Selected(i)
           Resin.PresaturantPercentage(j) = Percentages(i)
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
           lblIon(i - 1) = Cation(j).Name
           Percentages(i) = Resin.PresaturantPercentage(j)
           txtPercentage(i - 1) = Format$(Percentages(i), "0.00")
       Next i

       NumIonsSelected = NumSelectedCations

    ElseIf Anions.Available Then
       For i = 1 To NumSelectedAnions
           lblIon(i - 1).visible = True
           txtPercentage(i - 1).visible = True
           j = Anions_Selected(i)
           lblIon(i - 1) = Anion(j).Name
           Percentages(i) = Resin.PresaturantPercentage(j)
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

Private Sub txtPercentage_GotFocus(Index As Integer)

    If Index = 0 Then Exit Sub

    Call TextGetFocus(txtPercentage(Index), Temp_Text)
End Sub

Private Sub txtPercentage_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtPercentage_LostFocus(Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer, SumPercentages As Double
    Dim IsError As Integer

    If Index = 0 Then Exit Sub

    Call TextHandleError(IsError, txtPercentage(Index), Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtPercentage(Index).Text)
       OldValue = Percentages(Index + 1)

       If (NewValue > 100#) Or (NewValue < 0#) Then
          txtPercentage(Index) = Temp_Text
          Exit Sub
       End If

       SumPercentages = NewValue
       For i = 2 To NumIonsSelected
           If i <> (Index + 1) Then SumPercentages = SumPercentages + Percentages(i)
       Next i

       If SumPercentages > 100# Then
          txtPercentage(Index) = Temp_Text
          Exit Sub
       End If

       Percentages(Index + 1) = NewValue
       Percentages(1) = 100# - SumPercentages
       txtPercentage(0) = Format$(Percentages(1), "0.00")
       
    End If

End Sub

