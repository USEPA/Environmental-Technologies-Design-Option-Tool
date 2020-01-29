VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmH2O2ExtinctionCoeffTable 
   Caption         =   "frmH2O2ExtinctionCoeffTable"
   ClientHeight    =   5925
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4980
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   4980
   Begin Threed.SSPanel SSPanel1 
      Height          =   3765
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   4035
      _Version        =   65536
      _ExtentX        =   7117
      _ExtentY        =   6641
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox Picture1 
         Height          =   3585
         Left            =   90
         Picture         =   "frmExtin.frx":0000
         ScaleHeight     =   3525
         ScaleWidth      =   3795
         TabIndex        =   3
         Top             =   60
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1860
      TabIndex        =   0
      Top             =   5460
      Width           =   1305
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1365
      Left            =   60
      TabIndex        =   2
      Top             =   3990
      Width           =   3765
      _Version        =   65536
      _ExtentX        =   6641
      _ExtentY        =   2408
      _StockProps     =   15
      Caption         =   "SSPanel2"
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox Picture2 
         Height          =   1215
         Left            =   60
         Picture         =   "frmExtin.frx":2C092
         ScaleHeight     =   1155
         ScaleWidth      =   3585
         TabIndex        =   4
         Top             =   60
         Width           =   3645
      End
   End
End
Attribute VB_Name = "frmH2O2ExtinctionCoeffTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmH2O2ExtinctionCoeffTable.left = frmPhotoChem.left
frmH2O2ExtinctionCoeffTable.top = frmPhotoChem.top + frmPhotoChem.height - frmH2O2ExtinctionCoeffTable.height
frmH2O2ExtinctionCoeffTable.Caption = "Molecular Extinction Coefficients for Hydrogen Peroxide"
SSPanel1.Caption = ""
SSPanel2.Caption = ""
SSPanel1.left = frmH2O2ExtinctionCoeffTable.width / 2 - SSPanel1.width / 2
SSPanel2.left = frmH2O2ExtinctionCoeffTable.width / 2 - SSPanel2.width / 2
cmdClose.left = frmH2O2ExtinctionCoeffTable.width / 2 - cmdClose.width / 2
End Sub


