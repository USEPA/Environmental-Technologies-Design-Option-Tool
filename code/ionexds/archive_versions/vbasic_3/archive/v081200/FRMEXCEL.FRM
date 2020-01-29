VERSION 2.00
Begin Form frmExcel 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Transfer results to Excel"
   ClientHeight    =   2955
   ClientLeft      =   1650
   ClientTop       =   1890
   ClientWidth     =   5190
   Height          =   3360
   Left            =   1590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5190
   Top             =   1545
   Width           =   5310
   Begin SSFrame fraLang 
      Caption         =   "Select the language:"
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   2175
      Begin SSOption optFrench 
         Caption         =   "&French"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin SSOption optEnglish 
         Caption         =   "&English"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
   End
   Begin CommonDialog CMDialog1 
      Left            =   600
      Top             =   120
   End
   Begin TextBox txtTemp 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin SSFrame frmExcel 
      Caption         =   "Select the version:"
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   2055
      Begin SSOption optExcel 
         Caption         =   "Excel &5.0"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin SSOption optExcel 
         Caption         =   "Excel &4.0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin SSCommand cmdTransfer 
      Caption         =   "&Transfer"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin SSCommand cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
End
Option Explicit

Sub cmdCancel_Click ()
  Unload Me
End Sub

Sub cmdCancel_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub cmdTransfer_Click ()
Dim TaskID As Integer, I As Integer, J As Integer
Dim Version_Excel_Number  As Integer
Dim Filename_Excel As String, Temp  As String, Row As String
Dim f As Integer

   Excel_4 = OptExcel(0)
   If OptFrench Then
    Row = "L"
   Else
    If OptEnglish Then
     Row = "R"
    End If
   End If
   If Excel_4 Then
    Version_Excel_Number = 4
   Else
    Version_Excel_Number = 5
   End If
On Error GoTo File_Error
   CMDialog1.Filename = ""
   CMDialog1.CancelError = True
   CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.xls)|*.xls"
   CMDialog1.FilterIndex = 2
   CMDialog1.DialogTitle = "Save data points in Excel " & Version_Excel_Number & ".0"
'------Begin Modification Hokanson: 12-Aug2000
   CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
   CMDialog1.Action = 2

   f = FileNameIsValid(Filename_Excel, CMDialog1)
   If Not (f) Then Exit Sub

Begin_Execution:
 On Error GoTo Error_DDE_Excel
  txtTemp.LinkTimeout = 600
  txtTemp.LinkTopic = "Excel|System"
 

'Open DDE --------------------------------
  Dim Rep As Integer
  txtTemp.LinkMode = 2
  'txtTemp.LinkPoke
  txtTemp.LinkExecute "[NEW()]"
  If Dir(Filename_Excel) <> "" Then
   Rep = IDYES
   If Rep = IDYES Then
    Kill Filename_Excel
    txtTemp.LinkExecute "[SAVE.AS(" & Chr$(34) & Filename_Excel & Chr$(34) & ",1," & Chr$(34) & Chr$(34) & ",FALSE," & Chr$(34) & Chr$(34) & ",FALSE)]"
   Else
    Exit Sub
   End If
  Else
   txtTemp.LinkExecute "[SAVE.AS(" & Chr$(34) & Filename_Excel & Chr$(34) & ",1," & Chr$(34) & Chr$(34) & ",FALSE," & Chr$(34) & Chr$(34) & ",FALSE)]"
  End If
  txtTemp.LinkMode = 0
  Filename_Excel = File_Get_Rid_Of_Path(Filename_Excel)
  If Excel_4 Then
   txtTemp.LinkTopic = "Excel|" & Filename_Excel
  Else
   txtTemp.LinkTopic = "Excel|[" & Filename_Excel & "]Sheet1"
  End If
  txtTemp.LinkMode = 2

'---------------------- PFPDM  -------------------------------------
    Temp = "PFPDM results for "
    For I = 1 To Results.NComponent
     Temp = Temp & Trim$(Ion(Component_Index_PFPDM(I)).Name) & ", "
    Next I

    txtTemp = Temp
    txtTemp.LinkItem = Row & "1C1"
    txtTemp.LinkPoke

    txtTemp = "Time"
    txtTemp.LinkItem = Row & "2C1"
    txtTemp.LinkPoke

     If frmBreak!cboTime.ListIndex = 0 Then       'min
        txtTemp = "min"
     ElseIf frmBreak!cboTime.ListIndex = 1 Then   's
        txtTemp = "s"
     ElseIf frmBreak!cboTime.ListIndex = 2 Then   'hr
        txtTemp = "hr"
     ElseIf frmBreak!cboTime.ListIndex = 3 Then   'd
        txtTemp = "d"
     End If

    txtTemp.LinkItem = Row & "3C1"
    txtTemp.LinkPoke
    txtTemp = "BVF"
    txtTemp.LinkItem = Row & "2C2"
    txtTemp.LinkPoke
    txtTemp = "Usage Rate"
    txtTemp.LinkItem = Row & "2C3"
    txtTemp.LinkPoke
    txtTemp = "m3/kg of Resin"
    txtTemp.LinkItem = Row & "3C3"
    txtTemp.LinkPoke
    For I = 1 To Results.NComponent
      txtTemp = Trim$(Ion(Component_Index_PFPDM(I)).Name)
      txtTemp.LinkItem = Row & "2C" & Format$(I + 3, "0")
      txtTemp.LinkPoke
      txtTemp = "C/Ct"
      txtTemp.LinkItem = Row & "3C" & Format$(I + 3, "0")
      txtTemp.LinkPoke
    Next I

    For I = 1 To Results.NPoints

     If frmBreak!cboTime.ListIndex = 0 Then       'min
        txtTemp = Results.T(I)
     ElseIf frmBreak!cboTime.ListIndex = 1 Then   's
        txtTemp = Results.T(I) * 60#
     ElseIf frmBreak!cboTime.ListIndex = 2 Then   'hr
        txtTemp = Results.T(I) / 60#
     ElseIf frmBreak!cboTime.ListIndex = 3 Then   'd
        txtTemp = Results.T(I) / 60# / 24#
     End If

      txtTemp.LinkItem = Row & Format$(I + 3, "0") & "C1"
      txtTemp.LinkPoke

     txtTemp = Results.T(I) * 60# * Bed.Flowrate.Value / Bed.Length / Pi / (Bed.Diameter / 2) ^ 2
     txtTemp.LinkItem = Row & Format$(I + 3, "0") & "C2"
     txtTemp.LinkPoke

     txtTemp = Results.T(I) * 60# * Bed.Flowrate.Value / Bed.Weight
     txtTemp.LinkItem = Row & Format$(I + 3, "0") & "C3"
     txtTemp.LinkPoke
     For J = 1 To Results.NComponent
      txtTemp = Results.CP(J, I)
      txtTemp.LinkItem = Row & Format$(I + 3, "0") & "C" & Format$(J + 3, "0")
      txtTemp.LinkPoke
     Next J
    Next I

  txtTemp.LinkMode = 0
'Close DDE -------------------------------
  Unload Me
  Exit Sub

'-------------------------------------------------------------------
File_Error:
  If Err = 32755 Then
  Else
    MsgBox "Unknown error.", 48, Application_Name
  End If
  Resume Exit_DDE
Error_DDE_Excel:
  If Err = 282 Then
    If Not (Load_Excel()) Then
      MsgBox "Excel does not seem to be installed on this system.", MB_ICONEXCLAMATION, Application_Name
      Unload Me
      Exit Sub
    Else
      Resume Begin_Execution
    End If
  Else
    MsgBox "Excel is not responding properly.", MB_ICONEXCLAMATION, Application_Name
    Unload Me
    Exit Sub
  End If
The_End_Here:
  Resume Exit_DDE
Exit_DDE:
End Sub

Sub cmdTransfer_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub Form_Load ()
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
'    Me.HelpContextID = Hlp_Transfer_results
    OptExcel(1) = True
    OptEnglish = True
End Sub

Sub Key_Pressed_On_Control (Ascii_Code As Integer)
  Select Case Ascii_Code
    Case 67, 99 'C,c
      cmdCancel_Click
    Case 69, 101 'E,e
      OptEnglish = True
    Case 70, 102 'F,f
      OptFrench = True
    Case 84, 116'T,t
      cmdTransfer_Click
    Case 52 '4
      OptExcel(0) = True
    Case 53 '5
      OptExcel(1) = True
  End Select
End Sub

Function Load_Excel () As Integer
Dim TaskID As Integer
   On Error GoTo No_Excel
    TaskID = Shell("excel", 1)
    Load_Excel = True
    Exit Function
No_Excel:
  Load_Excel = False
  Resume Exit_Load_Excel
Exit_Load_Excel:
End Function

Sub optEnglish_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub optExcel_KeyPress (Index As Integer, KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub optFrench_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub txtTemp_LinkError (ErrNum As Integer)
Dim Msg As String
  Const OUT_OF_MEMORY = 11, WRONG_FORMAT = 1, TOO_MANY_DESTINATIONS = 7
  Const UPDATE_FAILED = 8

  Select Case ErrNum
   Case OUT_OF_MEMORY
    Msg = "Not Enough Memory to perform DDE"
   Case UPDATE_FAILED
    Msg = "Could not update data via DDE"
   Case TOO_MANY_DESTINATIONS
    Msg = "DDE SOURCE can not handle this many destinations"
   Case Else
    Msg = "unexpected DDE Error:" & ErrNum
  End Select
  If ErrNum <> WRONG_FORMAT Then
    MsgBox Msg, MB_ICONEXCLAMATION, "DDE Failure"
  End If
End Sub

