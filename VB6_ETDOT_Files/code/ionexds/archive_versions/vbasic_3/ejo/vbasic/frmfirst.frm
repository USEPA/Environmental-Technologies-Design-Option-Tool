VERSION 2.00
Begin Form frmFirst 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   330
   ClientTop       =   1035
   ClientWidth     =   9330
   ForeColor       =   &H00C0C0C0&
   Height          =   6480
   Left            =   270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9330
   Top             =   690
   Width           =   9450
   Begin SSPanel Label1 
      Caption         =   "Panel3D1"
      FloodColor      =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00808080&
      Height          =   1515
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   9075
   End
   Begin SSPanel Panel3D1 
      Caption         =   "Simulation Program for Ion Exchange"
      FloodColor      =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF00FF&
      Height          =   1575
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   8955
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
   Begin Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Demonstration Version 0.65 - May 5, 1995"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   1800
      Width           =   8955
   End
   Begin Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "David R. Hokanson, David W. Hand, John C. Crittenden"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   9075
   End
   Begin Label lblCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright Michigan Technological University - 1995"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   9075
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
    Panel3D1.Caption = "SPIEs" & Chr$(13) & "Simulation Program for Ion Exchange"
    Label1.Caption = "Department of Civil and Environmental Engineering" & Chr$(13) & "Michigan Technological University"
    Label2 = "David R. Hokanson, David W. Hand, and John C. Crittenden"
    Label1.BackColor = &HC0C0C0
    Label2.BackColor = &HC0C0C0
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

