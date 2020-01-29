VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frminfo 
   Caption         =   "Pearls Information"
   ClientHeight    =   3060
   ClientLeft      =   1695
   ClientTop       =   2025
   ClientWidth     =   7485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7485
   Begin VB.CommandButton acceptcmd 
      Caption         =   "Accept"
      Height          =   375
      Left            =   2820
      TabIndex        =   2
      Top             =   2580
      Width           =   1095
   End
   Begin VB.CommandButton cancelcmd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4095
      TabIndex        =   1
      Top             =   2565
      Width           =   1095
   End
   Begin VB.Frame rangepnl 
      Caption         =   "Temperature Range"
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.OptionButton Option2 
         Caption         =   "Legal Range"
         Height          =   255
         Left            =   5520
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Current Range"
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   7
         Top             =   1800
         Width           =   1515
      End
      Begin VB.TextBox maxtbx 
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox mintbx 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1920
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid rangegrd 
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2566
         _Version        =   65541
      End
      Begin VB.Label unitlbl 
         Caption         =   "Label1"
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   1980
         Width           =   435
      End
      Begin VB.Label tolbl 
         Caption         =   "to"
         Height          =   255
         Left            =   2580
         TabIndex        =   6
         Top             =   1980
         Width           =   315
      End
      Begin VB.Label selectlbl 
         Caption         =   "Range:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1980
         Width           =   675
      End
   End
End
Attribute VB_Name = "frminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' these store this information so we don't have
' to keep finding it
Dim hold_min(20) As Double
Dim hold_max(20) As Double
Private Sub acceptcmd_Click()

    FRMGraphSet!TXTMinT.Text = mintbx.Text
    FRMGraphSet!TXTMaxT.Text = maxtbx.Text
    FRMInfo.Hide
End Sub

Private Sub cancelcmd_Click()
FRMInfo.Hide

End Sub


Private Sub Form_Load()
    CenterForm Me
End Sub

Private Sub Option1_Click(Index As Integer)

        If Option1(0) = True Then
        ' put the range from the GraphSet form in
            mintbx.Text = FRMGraphSet!TXTMinT.Text
            maxtbx.Text = FRMGraphSet!TXTMaxT.Text
        End If
    
    
    
End Sub

Private Sub Option2_Click()
    Dim legalmin As Double
    Dim legalmax As Double
    Dim i As Integer
    
        If Option2 = True Then
            legalmin = 999999.99
            legalmax = -999999.99
            For i = 1 To rangegrd.Rows - 1
                rangegrd.Row = i
                rangegrd.Col = 0
                If rangegrd.Text = "X" Then
                    rangegrd.Col = 3
                    If Trim(rangegrd.Text) = "na" Then
                        GoTo next_I
                    End If
                    If CDbl(rangegrd.Text) < legalmin Then
                        legalmin = CDbl(rangegrd.Text)
                    End If
                    rangegrd.Col = 4
                    If CDbl(rangegrd.Text) > legalmax Then
                        legalmax = CDbl(rangegrd.Text)
                    End If
                End If
next_I:
            Next i
            ' if nothing is checked, don't give a legal
            ' min and max
            If legalmin <> 999999.99 Then
                mintbx.Text = legalmin
            Else
                mintbx.Text = " "
            End If
            If legalmax <> -999999.99 Then
                maxtbx.Text = legalmax
            Else
                maxtbx.Text = " "
            End If
    
        End If
        
        
    
End Sub


Private Sub rangegrd_Click()

    Dim legalmin As Double
    Dim legalmax As Double
    Dim RowSelected As Integer
    Dim ColSelected As Integer
    Dim i As Integer
       
    ' make sure the whole row is selected
    RowSelected = Me!rangegrd.Row
    ColSelected = Me!rangegrd.Col
    'SelStartCol = 1
    'SelEndCol = 4
    
    ' now handle either selecting or deselecting
    
    rangegrd.Col = 0
    If rangegrd.Text <> "X" Then
        rangegrd.Text = "X"
    Else
        rangegrd.Text = " "
    End If
    ' need to rewrite the text boxes, depending
    ' on which option is selected
    If Option2 = True Then
       ' figure out the legal range for the selected chems
        legalmin = 999999.99
        legalmax = -999999.99
        For i = 1 To rangegrd.Rows - 1
            rangegrd.Row = i
            rangegrd.Col = 0
            If rangegrd.Text = "X" Then
                
                rangegrd.Col = 3
                If Trim(rangegrd.Text) = "na" Then
                    GoTo next_I
                End If
                If CDbl(rangegrd.Text) < legalmin Then
                    legalmin = CDbl(rangegrd.Text)
                End If
                rangegrd.Col = 4
                If CDbl(rangegrd.Text) > legalmax Then
                    legalmax = CDbl(rangegrd.Text)
                End If
            End If
next_I:
        Next i
            ' if nothing is checked, don't give a legal
            ' min and max
        If legalmin <> 999999.99 Then
            mintbx.Text = legalmin
        Else
            mintbx.Text = " "
        End If
        If legalmax <> -999999.99 Then
            maxtbx.Text = legalmax
        Else
            maxtbx.Text = " "
        End If
    End If
     
    
End Sub

