VERSION 2.00
Begin Form frmResinPresaturantConditions 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Initial Resin Phase Presaturant Conditions"
   ClientHeight    =   5820
   ClientLeft      =   2010
   ClientTop       =   1500
   ClientWidth     =   7365
   Height          =   6225
   Left            =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Top             =   1155
   Width           =   7485
   Begin CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   6000
      TabIndex        =   24
      Top             =   5220
      Width           =   1095
   End
   Begin CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   5220
      Width           =   1095
   End
   Begin SSFrame fraPresaturant 
      Caption         =   "Percentage of Each Ion Presaturating the Resin"
      ForeColor       =   &H00000000&
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   9
         Left            =   5160
         TabIndex        =   22
         Top             =   4380
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   8
         Left            =   5160
         TabIndex        =   20
         Top             =   3960
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   7
         Left            =   5160
         TabIndex        =   18
         Top             =   3600
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   6
         Left            =   5160
         TabIndex        =   16
         Top             =   3180
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   5
         Left            =   5160
         TabIndex        =   14
         Top             =   2760
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   4
         Left            =   5160
         TabIndex        =   12
         Top             =   2340
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   3
         Left            =   5160
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   8
         Top             =   1500
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Height          =   285
         Index           =   1
         Left            =   5160
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin TextBox txtPercentage 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   3
         Top             =   660
         Width           =   1335
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   300
         TabIndex        =   21
         Top             =   4380
         Width           =   4695
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   300
         TabIndex        =   19
         Top             =   3960
         Width           =   4695
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   17
         Top             =   3600
         Width           =   4695
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   15
         Top             =   3180
         Width           =   4695
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   300
         TabIndex        =   13
         Top             =   2760
         Width           =   4695
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   11
         Top             =   2340
         Width           =   4695
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   9
         Top             =   1920
         Width           =   4695
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   1500
         Width           =   4695
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   1080
         Width           =   4695
      End
      Begin Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         Height          =   195
         Left            =   5160
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
      Begin Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ion"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   300
         Width           =   4095
      End
      Begin Label lblIon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   660
         Width           =   4695
      End
   End
End
Option Base 1
Option Explicit

Dim NumIonsSelected  As Integer
Dim Temp_Text As String

Dim Percentages(1 To MAX_CHEMICAL) As Double

Sub cmdCancel_Click ()
    Unload Me
End Sub

Sub cmdOK_Click ()
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

Sub Form_Load ()
    Dim i As Integer, j As Integer

    For i = 0 To 9
        lblIon(i).Visible = False
        txtPercentage(i).Visible = False
    Next i
    
    If Cations.Available And Anions.Available Then

    ElseIf Cations.Available Then
       
       For i = 1 To NumSelectedCations
           lblIon(i - 1).Visible = True
           txtPercentage(i - 1).Visible = True
           j = Cations_Selected(i)
           lblIon(i - 1) = Cation(j).Name
           Percentages(i) = Resin.PresaturantPercentage(j)
           txtPercentage(i - 1) = Format$(Percentages(i), "0.00")
       Next i

       NumIonsSelected = NumSelectedCations

    ElseIf Anions.Available Then
       For i = 1 To NumSelectedAnions
           lblIon(i - 1).Visible = True
           txtPercentage(i - 1).Visible = True
           j = Anions_Selected(i)
           lblIon(i - 1) = Anion(j).Name
           Percentages(i) = Resin.PresaturantPercentage(j)
           txtPercentage(i - 1) = Format$(Percentages(i), "0.00")
       Next i

       NumIonsSelected = NumSelectedAnions

    End If

       fraPresaturant.Height = lblIon(NumIonsSelected - 1).Top + lblIon(NumIonsSelected - 1).Height + 120
       cmdOK.Top = fraPresaturant.Top + fraPresaturant.Height + 180
       cmdCancel.Top = cmdOK.Top
       frmResinPresaturantConditions.Height = cmdOK.Top + cmdOK.Height + 540

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.Width - frmResinPresaturantConditions.Width) / 2, (Screen.Height - frmResinPresaturantConditions.Height) / 2
    End If
       
End Sub

Sub txtPercentage_GotFocus (index As Integer)

    If index = 0 Then Exit Sub

    Call TextGetFocus(txtPercentage(index), Temp_Text)
End Sub

Sub txtPercentage_KeyPress (index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Sub txtPercentage_LostFocus (index As Integer)
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

