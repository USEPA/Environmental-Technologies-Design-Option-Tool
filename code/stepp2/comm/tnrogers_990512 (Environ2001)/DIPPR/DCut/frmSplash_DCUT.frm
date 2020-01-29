VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash_DCUT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   0
      Picture         =   "frmSplash_DCUT.frx":000C
      ScaleHeight     =   5865
      ScaleWidth      =   7530
      TabIndex        =   0
      Top             =   0
      Width           =   7530
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Picture1_click()
    Unload Me
End Sub

