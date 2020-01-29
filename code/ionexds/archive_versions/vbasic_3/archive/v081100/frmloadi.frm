VERSION 2.00
Begin Form frmLoading 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Plug-Flow Pore Diffusion Model for Ion Exchange"
   ClientHeight    =   2550
   ClientLeft      =   2250
   ClientTop       =   3375
   ClientWidth     =   5310
   ControlBox      =   0   'False
   Height          =   2955
   Left            =   2190
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5310
   Top             =   3030
   Width           =   5430
   Begin Label lblTotalBeds 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Number of Beds:"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin Label lblTime 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   2115
   End
   Begin Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Started This Bed:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin Label lblTime 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   2115
   End
   Begin Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Calculations Began:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Currently Calculating Results for Bed Number 1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1620
      Visible         =   0   'False
      Width           =   4455
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Initializing... Please wait."
      Height          =   252
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4452
   End
End
Option Explicit

Sub Form_Load ()
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
End Sub

Sub Form_Unload (Cancel As Integer)
  frmIonExchangeMain.Enabled = True
  frmIonExchangeMain.Show
End Sub

