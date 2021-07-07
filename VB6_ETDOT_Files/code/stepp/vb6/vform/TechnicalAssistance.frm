VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmTechnicalAssistance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Model & Software Developers"
   ClientHeight    =   2850
   ClientLeft      =   4125
   ClientTop       =   6660
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3540
      TabIndex        =   1
      Top             =   2400
      Width           =   1365
   End
   Begin Threed.SSPanel pnl_title 
      Height          =   2235
      Index           =   3
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   3942
      _StockProps     =   15
      Caption         =   "pnl_title(3)"
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
   End
End
Attribute VB_Name = "frmTechnicalAssistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Dim msg As String
If WindowState = 0 Then
  'don't attempt if screen Minimized or Maximized
  Move contam_prop_form.Left + (contam_prop_form.Width / 2) - (Me.Width / 2), contam_prop_form.Top + (contam_prop_form.Height / 2) - (Me.Height / 2)

End If

'msg = "Programming Support:    Richard J. Hossli" & Chr$(13) & "                                    Jason E. Mclean" & Chr$(13) & "                                    Eric J. Oman" & Chr$(13)
'pnl_title(5).Caption = msg & "                                    Thomas F. Budd" & Chr$(13) & "                                    Kristine L. Grove"
'pnl_title(5).BackColor = &HC0C0C0
'pnl_title(5).ForeColor = &H0&

msg = "David R. Hokanson" & Chr$(13)
msg = msg & "Tony N. Rogers" & Chr$(13)
msg = msg & "David W. Hand" & Chr$(13)
msg = msg & "John C. Crittenden" & Chr$(13)
msg = msg & "Fr" & Chr$(233) & "d" & Chr$(233) & "ric Gobin" & Chr$(13)
msg = msg & "Eric J. Oman"
pnl_title(3).Caption = msg

pnl_title(3).BackColor = &HC0C0C0
pnl_title(3).ForeColor = &H0&

''''picmtu(0).Picture = LoadPicture(app.Path & "\mtu_logo.bmp")
End Sub


