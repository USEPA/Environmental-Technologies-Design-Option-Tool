VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   4035
   ClientTop       =   4515
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
'Dim DG As Double
'Dim TEMP As Double
'Dim PRES As Double
'  TEMP = 298.15
'  PRES = 1#
'  Call AIRDENS(DG, TEMP, PRES)
'  MsgBox "DG = " & Trim$(Str$(DG))

Dim A1 As Double
Dim A2 As Double
  A1 = 1.2
  Call Go_TESTSUB1(A1, A2)
  MsgBox "DG = " & Trim$(Str$(A2))

'SUBROUTINE AIRDENS(DG, TEMP, PRES)
Dim DG As Double
Dim TEMP As Double
Dim PRES As Double
  DG = 0#     'kg/m^3
  TEMP = 298.15   'degK
  PRES = 1#     'atm
  Call Go_Hokanson_AIRDENS(DG, TEMP, PRES)
  MsgBox "DG(Hokanson) = " & Trim$(Str$(DG))

Dim CS As Double
Dim VQ As Double
Dim HC As Double
Dim CI As Double
Dim CE As Double
  VQ = 1#
  HC = 1#
  CI = 100#
  CE = 5#
  'Call GETCSPT(CS, VQ, HC, CI, CE)
  Call Go_Hokanson_GETCSPT(CS, VQ, HC, CI, CE)
  MsgBox "CS(Hokanson) = " & Trim$(Str$(CS))


'------------------------------------------------------------
  HC = 1#
  CI = 100#
  CE = 5#
  'Call GETCSPT(CS, VQ, HC, CI, CE)
  Call Go_asapptad_test4_GETCSPT(CS, VQ, HC, CI, CE)
  MsgBox "CS(asapptad_test4) = " & Trim$(Str$(CS))

  HC = 1#
  CI = 100#
  CE = 5#
  'Call GETCSPT(CS, VQ, HC, CI, CE)
  Call Go_asapptad_test3_GETCSPT(CS, VQ, HC, CI, CE)
  MsgBox "CS(asapptad_test3) = " & Trim$(Str$(CS))



'------------------------------------------------------------

'SUBROUTINE AIRDENS(DG, TEMP, PRES)
'Dim DG As Double
'Dim TEMP As Double
'Dim PRES As Double
  DG = 0#     'kg/m^3
  TEMP = 298.15   'degK
  PRES = 1#     'atm
  Call Go_AIRDENS(DG, TEMP, PRES)
  MsgBox "DG = " & Trim$(Str$(DG))



'Dim CS As Double
'Dim VQ As Double
'Dim HC As Double
'Dim CI As Double
'Dim CE As Double
  VQ = 1#
  HC = 1#
  CI = 100#
  CE = 5#
  'Call GETCSPT(CS, VQ, HC, CI, CE)
  Call Go_GETCSPT(CS, VQ, HC, CI, CE)
  MsgBox "CS = " & Trim$(Str$(CS))
 
'          CS = (1# / (VQ * HC)) * (CI - CE)

End Sub

Private Sub Form_Load()
Dim ThisPath As String
Dim SearchFor As String
  ThisPath = App.Path
  SearchFor = "\\cen-server\srcsafe"
  ''''ThisPath = "X:\etdot10\code\asap\vb5_forcode\tests\asap1\vb5_test2"
  If (Left$(UCase$(ThisPath), Len(SearchFor)) = UCase$(SearchFor)) Then
    ThisPath = Right$(ThisPath, Len(ThisPath) - Len(SearchFor))
    ThisPath = "x:" & ThisPath
  End If
  ''''MsgBox ThisPath
  ChDir ThisPath
  ChDrive ThisPath
  MsgBox ThisPath
End Sub
