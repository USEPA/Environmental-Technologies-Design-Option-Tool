VERSION 2.00
Begin Form frmFirst 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   1155
   ClientTop       =   2880
   ClientWidth     =   9450
   ForeColor       =   &H00C0C0C0&
   Height          =   6480
   Left            =   1095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9450
   Top             =   2535
   Width           =   9570
   Begin SSPanel panelCopyright 
      Caption         =   "panelCopyright"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   5040
      TabIndex        =   10
      Top             =   1800
      Width           =   3975
   End
   Begin SSPanel panelAuthors 
      Caption         =   "panelAuthors"
      FloodColor      =   &H0000FFFF&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   5040
      TabIndex        =   2
      Top             =   3360
      Width           =   3975
   End
   Begin SSPanel Label1 
      FloodColor      =   &H0000FFFF&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3795
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
      Begin PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   1440
         Picture         =   FRMFIRST.FRX:0000
         ScaleHeight     =   1695
         ScaleWidth      =   1935
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
      Begin PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   720
         Picture         =   FRMFIRST.FRX:E63A
         ScaleHeight     =   735
         ScaleWidth      =   3375
         TabIndex        =   5
         Top             =   3000
         Width           =   3375
      End
      Begin Label lblCenCITT_MTU 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "lblCenCITT_MTU"
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   4455
      End
   End
   Begin SSPanel Panel3D1 
      FloodColor      =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   8955
      Begin Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TM"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   7570
         TabIndex        =   9
         Top             =   450
         Width           =   375
      End
      Begin Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ion Exchange Design Software (IonExDesignS    )"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   15
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   878
         TabIndex        =   8
         Top             =   480
         Width           =   7200
      End
   End
   Begin SSCommand cmdContinue 
      Caption         =   "&Continue"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5460
      Width           =   1335
   End
   Begin SSCommand cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   5460
      Width           =   1455
   End
End
Option Explicit

Sub cmdContinue_Click ()
  Load frmIonExchangeMain
End Sub

Sub cmdContinue_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub cmdExit_Click ()
  End
End Sub

Sub cmdExit_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)
End Sub

Sub Form_Load ()

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
    panelCopyright.Caption = "Version 1.0" & Chr$(13) & Chr$(13) & "Copyright 1995 - 2002"
    panelAuthors.Caption = "David R. Hokanson" & Chr$(13) & Chr$(13) & "David W. Hand" & Chr$(13) & Chr$(13) & "John C. Crittenden"
    lblCenCITT_MTU = "National Center for" & Chr$(13) & "Clean Industrial and Treatment Technologies" & Chr$(13) & Chr$(13) & "Michigan Technological University" & Chr$(13) & "Houghton, Michigan"
    Label1.BackColor = &HC0C0C0
    panelAuthors.BackColor = &HC0C0C0
    Panel3D1.BackColor = &HC0C0C0

End Sub

Sub Key_Pressed_On_Control (Ascii_Code As Integer)
  Select Case Ascii_Code
    Case 67, 99 'C,c
      cmdContinue_Click
    Case 88, 120'X,x
      cmdExit_Click
  End Select
End Sub

