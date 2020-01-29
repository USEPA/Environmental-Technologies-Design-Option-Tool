VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmExtraDisclaimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disclaimer"
   ClientHeight    =   6795
   ClientLeft      =   3495
   ClientTop       =   2055
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdIAgree 
      Appearance      =   0  'Flat
      Caption         =   "I Agree"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   90
      TabIndex        =   3
      Top             =   6090
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8310
      TabIndex        =   2
      Top             =   6090
      Width           =   972
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   4935
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   8705
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
      Begin VB.Label lblDisclaimer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "lblDisclaimer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   8865
      End
   End
End
Attribute VB_Name = "frmExtraDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

  Unload Me
  End

End Sub

Private Sub cmdIAgree_Click()

  Unload Me

End Sub

Private Sub Form_Load()
Dim msg As String

  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2

  msg = "This demonstration version of StEPP is "
  msg = msg & "FOR PROPOSAL EVALUATION ONLY by the "
  msg = msg & "reviewers of:"
  msg = msg & Chr$(13) & Chr$(13)
  msg = msg & "    Vicksburg Consolidated Contracting Office"
  msg = msg & Chr$(13)
  msg = msg & "    U.S. Army Corps of Engineers"
  msg = msg & Chr$(13)
  msg = msg & "    Solicitation No.:  DACW39-98-R-0008"
  msg = msg & Chr$(13)
  msg = msg & "    RFP Title:  Environmental and Ecological Risk Assessment / Modeling"
  msg = msg & Chr$(13)
  msg = msg & Chr$(13)
  msg = msg & "This demonstration version expires on "
  msg = msg & "April 30, 1999 and is NOT for re-sale "
  msg = msg & "or use."
  lblDisclaimer.Caption = msg

End Sub


