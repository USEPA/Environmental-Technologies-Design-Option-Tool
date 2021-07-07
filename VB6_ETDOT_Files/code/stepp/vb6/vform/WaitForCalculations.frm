VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmWaitForCalculations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "StEPP"
   ClientHeight    =   2280
   ClientLeft      =   5115
   ClientTop       =   3840
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel Panel3D1 
      Height          =   1995
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   3519
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel Panel3D2 
         Height          =   735
         Left            =   180
         TabIndex        =   1
         Top             =   630
         Visible         =   0   'False
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Panel3D2"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmWaitForCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.Width - frmWaitForCalculations.Width) / 2, (Screen.Height - frmWaitForCalculations.Height) / 2
    End If

    Panel3D1.BackColor = &HC0C0C0
    Panel3D1.ForeColor = &H0&

    Panel3D1.FontSize = 13.8
    Panel3D1.Caption = "Performing Calculations" & Chr$(13) & Chr$(13) & "Please Wait"

End Sub


Private Sub SSPanel1_Click()

End Sub
