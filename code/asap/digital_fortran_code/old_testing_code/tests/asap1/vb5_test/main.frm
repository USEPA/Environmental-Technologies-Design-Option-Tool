VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   4785
   ClientTop       =   3930
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   5145
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   555
      Left            =   1110
      TabIndex        =   0
      Top             =   360
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Const Form1_declarations_end = True


Private Sub Command1_Click()
Dim DG As Double
Dim TEMP As Double
Dim PRES As Double
  TEMP = 298.15
  PRES = 1#
  Call AIRDENS(DG, TEMP, PRES)
  MsgBox "DG = " & Trim$(Str$(DG))
End Sub

Private Sub Form_Load()
  ChDir "X:\etdot10\code\asap\vb5_forcode\tests\asap1\vb5_test"
  ChDrive "X:\etdot10\code\asap\vb5_forcode\tests\asap1\vb5_test"
End Sub
