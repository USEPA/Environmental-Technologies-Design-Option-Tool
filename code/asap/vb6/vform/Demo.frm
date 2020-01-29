VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demo Version"
   ClientHeight    =   3975
   ClientLeft      =   2730
   ClientTop       =   2340
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdButton1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6900
      TabIndex        =   1
      Top             =   3300
      Width           =   1305
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   3075
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8085
      _Version        =   65536
      _ExtentX        =   14261
      _ExtentY        =   5424
      _StockProps     =   15
      ForeColor       =   -2147483640
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
      Begin VB.Label lblDisclaimer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "lblDisclaimer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Width           =   7845
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Const frmDemo_decl_end = True


Sub frmDemo_GO()
  frmDemo.Show 1
End Sub


Private Sub cmdButton1_Click()
  Unload Me
  Exit Sub
End Sub
Private Sub cmdExit_Click()
  End
End Sub


Private Sub Form_Load()
  Call CenterOnScreen(Me)
  lblDisclaimer.Caption = _
      "This is a DEMONSTRATION version of the ASAP program. " & _
      "This demonstration version may only be used to run calculations " & _
      "with the default data. " & _
      "For the full version of this program, please contact " & _
      "Dr. David W. Hand (dwhand@mtu.edu or 906-487-2777). " & _
      "Additional information about this program is available at " & _
      "our web site (http://www.cpas.mtu.edu/etdot/)."
End Sub



