VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDimensionless 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dimensionless Groups"
   ClientHeight    =   5145
   ClientLeft      =   915
   ClientTop       =   810
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5700
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   20
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   5640
      Picture         =   "Dimensionless.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   9885
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   9885
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   4365
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   7699
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtDimless 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "txtDimless(6)"
         Top             =   3300
         Width           =   1200
      End
      Begin VB.TextBox txtDimless 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "txtDimless(5)"
         Top             =   2940
         Width           =   1200
      End
      Begin VB.TextBox txtDimless 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "txtDimless(4)"
         Top             =   2580
         Width           =   1200
      End
      Begin VB.TextBox txtDimless 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "txtDimless(3)"
         Top             =   2220
         Width           =   1200
      End
      Begin VB.TextBox txtDimless 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "txtDimless(2)"
         Top             =   1860
         Width           =   1200
      End
      Begin VB.TextBox txtDimless 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "txtDimless(1)"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.TextBox txtDimless 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "txtDimless(0)"
         Top             =   1140
         Width           =   1200
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   795
         Left            =   180
         TabIndex        =   8
         Top             =   150
         Width           =   5115
         _Version        =   65536
         _ExtentX        =   9022
         _ExtentY        =   1402
         _StockProps     =   14
         Caption         =   "Select Component:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboSelectCompo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   300
            Width           =   4875
         End
      End
      Begin Threed.SSCommand cmdDefs 
         Height          =   495
         Left            =   1680
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3720
         Visible         =   0   'False
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Definitions ..."
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
      Begin VB.Label lblDimless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Surface Solute Distribution Parameter, Dgs ="
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
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   3330
         Width           =   3855
      End
      Begin VB.Label lblDimless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pore Solute Distribution Parameter, Dgp ="
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
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   15
         Top             =   2970
         Width           =   3615
      End
      Begin VB.Label lblDimless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Surface Biot Number, Bis ="
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
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   2610
         Width           =   3615
      End
      Begin VB.Label lblDimless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pore Biot Number, Bip ="
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
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   13
         Top             =   2250
         Width           =   3615
      End
      Begin VB.Label lblDimless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pore Diffusivity Modulus, Edp ="
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
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   1890
         Width           =   3615
      End
      Begin VB.Label lblDimless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Surface Diffusivity Modulus, Eds ="
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
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   1530
         Width           =   3615
      End
      Begin VB.Label lblDimless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stanton number, St ="
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
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1170
         Width           =   3615
      End
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   495
      Left            =   3720
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Close"
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
Attribute VB_Name = "frmDimensionless"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub populate_cboSelectCompo()
Dim i As Integer
  cboSelectCompo.Clear
  For i = 1 To Number_Component
    cboSelectCompo.AddItem Trim$(Component(i).Name)
  Next i
End Sub


Sub Display_Component(Which As Integer)
Dim N As Integer
  If (Which < 1) Or (Which > Number_Component) Then
    Exit Sub
  End If
  N = Which
  Call AssignTextAndTag(txtDimless(0), NumberToMFBString(ST(N)))
  Call AssignTextAndTag(txtDimless(1), NumberToMFBString(Eds(N)))
  Call AssignTextAndTag(txtDimless(2), NumberToMFBString(Edp(N)))
  Call AssignTextAndTag(txtDimless(3), NumberToMFBString(Bip(N)))
  Call AssignTextAndTag(txtDimless(4), NumberToMFBString(Bis(N)))
  Call AssignTextAndTag(txtDimless(5), NumberToMFBString(Dgp(N)))
  Call AssignTextAndTag(txtDimless(6), NumberToMFBString(Dgs(N)))
  If (Which <= cboSelectCompo.ListCount) Then
    cboSelectCompo.ListIndex = Which - 1
  End If
End Sub


Private Sub cboSelectCompo_Click()
  Call Display_Component(cboSelectCompo.ListIndex + 1)
End Sub


Private Sub cmdDefs_Click()
  'Me.Hide
  'frmDimensionlessDefs.Show
  'Me.Show
End Sub
Private Sub cmdOK_Click()
  Unload Me
End Sub


Private Sub Command4_Click()
    Set Picture2.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture2.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus

End Sub

Private Sub Form_Load()
  'MISC INITS.
  Call CenterOnForm(Me, frmMain)
  Call populate_cboSelectCompo
  Call Display_Component(frmMain.cboSelectCompo.ListIndex + 1)
End Sub


Private Sub txtDimless_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtDimless(Index)
  Call Global_GotFocus(Ctl)
End Sub
Private Sub txtDimless_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_ReadOnlyKeyPress(KeyAscii)
End Sub
Private Sub txtDimless_LostFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtDimless(Index)
  Call Global_LostFocus(Ctl)
End Sub


