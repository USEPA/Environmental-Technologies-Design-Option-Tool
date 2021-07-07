VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDimensionlessDefs 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dimensionless Groups : Definitions"
   ClientHeight    =   4170
   ClientLeft      =   3510
   ClientTop       =   3225
   ClientWidth     =   10020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   10020
   Begin Threed.SSCommand cmdOK 
      Height          =   765
      Left            =   8400
      TabIndex        =   1
      Top             =   90
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2672
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "&OK"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   60
      Picture         =   "DimensionlessDefs.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   9885
      TabIndex        =   0
      Top             =   60
      Width           =   9885
   End
End
Attribute VB_Name = "frmDimensionlessDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
  Unload Me
End Sub
