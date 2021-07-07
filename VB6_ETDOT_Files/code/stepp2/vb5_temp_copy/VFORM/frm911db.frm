VERSION 5.00
Begin VB.Form frm911DBInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6390
   ClientLeft      =   900
   ClientTop       =   1545
   ClientWidth     =   7830
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
   Icon            =   "frm911db.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6390
   ScaleWidth      =   7830
   Begin VB.TextBox TXTMaxTemp 
      Height          =   285
      Left            =   2355
      TabIndex        =   22
      Text            =   "TXTMaxTemp"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TXTMinTemp 
      Height          =   285
      Left            =   1515
      TabIndex        =   21
      Text            =   "TXTMinTemp"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TXT801Code 
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "TXT801Code"
      Top             =   1620
      Width           =   1575
   End
   Begin VB.TextBox TXTPressure 
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "TXTPressure"
      Top             =   810
      Width           =   1575
   End
   Begin VB.TextBox TXTTemperature 
      Height          =   285
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "TXTTemperature"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox TXTRating 
      Height          =   285
      Left            =   4995
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "TXTRating"
      Top             =   1170
      Width           =   1575
   End
   Begin VB.TextBox TXTValue 
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "TXTValue"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox TXTProperty 
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "TXTProperty"
      Top             =   480
      Width           =   5175
   End
   Begin VB.TextBox TXTChemName 
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "TXTChemName"
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox TXTCAS 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "TXTCAS"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CMDClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   5850
      Width           =   1455
   End
   Begin VB.TextBox TXTComment 
      Height          =   1380
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Text            =   "frm911db.frx":030A
      Top             =   3945
      Width           =   7575
   End
   Begin VB.TextBox TXTCitations 
      Height          =   1290
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Text            =   "frm911db.frx":0315
      Top             =   2340
      Width           =   7575
   End
   Begin VB.Label PressUnits 
      Caption         =   "PressUnits"
      Height          =   285
      Left            =   6705
      TabIndex        =   26
      Top             =   855
      Width           =   960
   End
   Begin VB.Label TempUnits 
      Caption         =   "TempUnits"
      Height          =   285
      Left            =   3015
      TabIndex        =   25
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label ValUnits 
      Caption         =   "ValUnits"
      Height          =   285
      Left            =   3030
      TabIndex        =   24
      Top             =   855
      Width           =   780
   End
   Begin VB.Label LBLTempRange 
      Alignment       =   1  'Right Justify
      Caption         =   "Temp Range"
      Height          =   255
      Left            =   135
      TabIndex        =   23
      Top             =   1215
      Width           =   1155
   End
   Begin VB.Label LBLPressure 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pressure"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4275
      TabIndex        =   6
      Top             =   810
      Width           =   735
   End
   Begin VB.Label LBLTemperature 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temperature"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   9
      Top             =   1215
      Width           =   1095
   End
   Begin VB.Label LBLRating 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rating"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label LBLValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Value"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   810
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label LBLCitations 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Literature Citations"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2070
      Width           =   7575
   End
   Begin VB.Label LBL801Code 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Project 801 Code Format"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label LBLComment 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comment"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   3
      Top             =   3705
      Width           =   7575
   End
   Begin VB.Label LBLProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Property"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label LBLChemName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Chemical Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LBLCAS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CAS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   150
      Width           =   375
   End
End
Attribute VB_Name = "frm911DBInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CMDClose_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    CenterForm Me

End Sub

Private Sub Label1_Click()

End Sub

Private Sub TXTCitations_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub


Private Sub TXTComment_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub


