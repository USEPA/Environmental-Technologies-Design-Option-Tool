VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmTechAssistance 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Technical Support Provided By:"
   ClientHeight    =   2970
   ClientLeft      =   2295
   ClientTop       =   4965
   ClientWidth     =   5325
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2970
   ScaleWidth      =   5325
   Begin Threed.SSPanel pnl_title 
      Height          =   2205
      Index           =   3
      Left            =   60
      TabIndex        =   2
      Top             =   240
      Width           =   5205
      _Version        =   65536
      _ExtentX        =   9181
      _ExtentY        =   3889
      _StockProps     =   15
      Caption         =   "pnl_title(3)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3900
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1365
   End
   Begin VB.PictureBox pnl_titleX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00400040&
      Height          =   735
      Index           =   5
      Left            =   1020
      ScaleHeight     =   705
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   3300
      Visible         =   0   'False
      Width           =   5200
   End
End
Attribute VB_Name = "frmTechAssistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim msg As String

Call CenterOnForm(Me, frmMainMenu)
'Move frmpfpsdm.Left + (frmpfpsdm.Width / 2) - (Me.Width / 2), frmpfpsdm.Top + (frmpfpsdm.Height / 2) - (Me.Height / 2)

msg = "David R. Hokanson" & Chr$(13)
msg = msg & "David W. Hand" & Chr$(13)
msg = msg & "John C. Crittenden" & Chr$(13)
msg = msg & "Tony N. Rogers" & Chr$(13)
msg = msg & "Fr" & Chr$(233) & "d" & Chr$(233) & "ric Gobin" & Chr$(13)
msg = msg & "Eric J. Oman"
pnl_title(3).Caption = msg
pnl_title(3).BackColor = &HC0C0C0
pnl_title(3).ForeColor = &H0&
'pnl_title(3).Caption = " Model and Software:      Fr" & Chr$(233) & "d" & Chr$(233) & "ric Gobin     Eric J. Oman" & Chr$(13) & " Development   :      Tony N. Rogers"

'pnl_title(5).Caption = " Programing Support:     Richard J. Hossli" & Chr$(13) & "                                   Jason E. Mclean"
'pnl_title(5).BackColor = &HC0C0C0
'pnl_title(5).ForeColor = &H0&

''''picmtu(0).Picture = LoadPicture(app.Path & "\mtu_logo.bmp")
End Sub

