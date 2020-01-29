VERSION 5.00
Begin VB.Form frmLoading 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plug-Flow Pore Diffusion Model for Ion Exchange"
   ClientHeight    =   2550
   ClientLeft      =   2250
   ClientTop       =   3375
   ClientWidth     =   5310
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   5310
   Begin VB.Label lblTotalBeds 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Number of Beds:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   2115
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Started This Bed:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   2115
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Calculations Began:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Currently Calculating Results for Bed Number 1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1620
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Initializing... Please wait."
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4452
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    top = Screen.height / 2 - height / 2
    left = Screen.width / 2 - width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmIonExchangeMain.Enabled = True
  frmIonExchangeMain.Show
End Sub

