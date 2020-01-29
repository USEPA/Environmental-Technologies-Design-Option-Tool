VERSION 5.00
Begin VB.Form frmH2O2ExtinctionCoeffTable 
   Caption         =   "frmH2O2ExtinctionCoeffTable"
   ClientHeight    =   5550
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
   ScaleHeight     =   5550
   ScaleWidth      =   4980
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   5100
      Width           =   1305
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   60
      ScaleHeight     =   1125
      ScaleWidth      =   4365
      TabIndex        =   1
      Top             =   3810
      Width           =   4425
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   90
      ScaleHeight     =   3585
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   90
      Width           =   4395
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
Picture1.AutoSize = True
Picture1.Picture = LoadPicture("bitmaps\table.gif")
Picture2.AutoSize = True
Picture2.Picture = LoadPicture("bitmaps\ref.gif")
Picture1.left = frmH2O2ExtinctionCoeffTable.width / 2 - Picture1.width / 2
Picture2.left = frmH2O2ExtinctionCoeffTable.width / 2 - Picture2.width / 2
cmdClose.left = frmH2O2ExtinctionCoeffTable.width / 2 - cmdClose.width / 2
End Sub


