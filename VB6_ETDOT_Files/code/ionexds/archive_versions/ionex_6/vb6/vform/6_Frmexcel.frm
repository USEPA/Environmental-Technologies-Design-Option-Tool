VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExcel 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfer results to Excel"
   ClientHeight    =   2955
   ClientLeft      =   1650
   ClientTop       =   1890
   ClientWidth     =   5190
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2955
   ScaleWidth      =   5190
   Begin Threed.SSCommand cmdCancel 
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Cancel"
   End
   Begin Threed.SSFrame fraLang 
      Height          =   1695
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   2990
      _StockProps     =   14
      Caption         =   "Select the language:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optEnglish 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&English"
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
      Begin Threed.SSOption optFrench 
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&French"
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
   End
   Begin Threed.SSFrame fraExcel 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   2990
      _StockProps     =   14
      Caption         =   "Select the version:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optexcel 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Excel &4.0"
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
      Begin Threed.SSOption optexcel 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Excel &5.0"
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
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   2400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtTemp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Threed.SSCommand cmdTransfer 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Transfer"
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub cmdTransfer_Click()
Dim TaskID As Integer, i As Integer, j As Integer
Dim Version_Excel_Number  As Integer
Dim Filename_Excel As String, temp  As String, Row As String
Dim F As Integer

   Excel_4 = optexcel(0)
   If optFrench Then
    Row = "L"
   Else
    If optEnglish Then
     Row = "R"
    End If
   End If
   If Excel_4 Then
    Version_Excel_Number = 4
   Else
    Version_Excel_Number = 5
   End If
On Error GoTo File_Error
   CMDialog1.filename = ""
   CMDialog1.CancelError = True
   CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.xls)|*.xls"
   CMDialog1.FilterIndex = 2
   CMDialog1.DialogTitle = "Save data points in Excel " & Version_Excel_Number & ".0"
'------Begin Modification Hokanson: 12-Aug2000
   CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
   CMDialog1.Action = 2

   F = FileNameIsValid(Filename_Excel, CMDialog1)
   If Not (F) Then Exit Sub

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
    temp = "PFPDM results for "
    For i = 1 To Results.NComponent
     temp = temp & Trim$(Ion(Component_Index_PFPDM(i)).Name) & ", "
    Next i

    txtTemp = temp
    txtTemp.LinkItem = Row & "1C1"
    txtTemp.LinkPoke

    txtTemp = "Time"
    txtTemp.LinkItem = Row & "2C1"
    txtTemp.LinkPoke

     If frmbreak!cboTime.ListIndex = 0 Then       'min
        txtTemp = "min"
     ElseIf frmbreak!cboTime.ListIndex = 1 Then   's
        txtTemp = "s"
     ElseIf frmbreak!cboTime.ListIndex = 2 Then   'hr
        txtTemp = "hr"
     ElseIf frmbreak!cboTime.ListIndex = 3 Then   'd
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
    For i = 1 To Results.NComponent
      txtTemp = Trim$(Ion(Component_Index_PFPDM(i)).Name)
      txtTemp.LinkItem = Row & "2C" & Format$(i + 3, "0")
      txtTemp.LinkPoke
      txtTemp = "C/Ct"
      txtTemp.LinkItem = Row & "3C" & Format$(i + 3, "0")
      txtTemp.LinkPoke
    Next i

    For i = 1 To Results.NPoints

     If frmbreak!cboTime.ListIndex = 0 Then       'min
        txtTemp = Results.T(i)
     ElseIf frmbreak!cboTime.ListIndex = 1 Then   's
        txtTemp = Results.T(i) * 60#
     ElseIf frmbreak!cboTime.ListIndex = 2 Then   'hr
        txtTemp = Results.T(i) / 60#
     ElseIf frmbreak!cboTime.ListIndex = 3 Then   'd
        txtTemp = Results.T(i) / 60# / 24#
     End If

      txtTemp.LinkItem = Row & Format$(i + 3, "0") & "C1"
      txtTemp.LinkPoke

     txtTemp = Results.T(i) * 60# * NowProj.Bed.Flowrate.Value / NowProj.Bed.length / Pi / (NowProj.Bed.Diameter / 2) ^ 2
     txtTemp.LinkItem = Row & Format$(i + 3, "0") & "C2"
     txtTemp.LinkPoke

     txtTemp = Results.T(i) * 60# * NowProj.Bed.Flowrate.Value / NowProj.Bed.Weight
     txtTemp.LinkItem = Row & Format$(i + 3, "0") & "C3"
     txtTemp.LinkPoke
     For j = 1 To Results.NComponent
      txtTemp = Results.CP(j, i)
      txtTemp.LinkItem = Row & Format$(i + 3, "0") & "C" & Format$(j + 3, "0")
      txtTemp.LinkPoke
     Next j
    Next i

  txtTemp.LinkMode = 0
'Close DDE -------------------------------
  Unload Me
  Exit Sub

'-------------------------------------------------------------------
File_Error:
  If Err = 32755 Then
  Else
    MsgBox "Unknown error.", 48, App.title
  End If
  Resume Exit_DDE
Error_DDE_Excel:
  If Err = 282 Then
    If Not (Load_Excel()) Then
      MsgBox "Excel does not seem to be installed on this system.", MB_ICONEXCLAMATION, App.title
      Unload Me
      Exit Sub
    Else
      Resume Begin_Execution
    End If
  Else
    MsgBox "Excel is not responding properly.", MB_ICONEXCLAMATION, App.title
    Unload Me
    Exit Sub
  End If
The_End_Here:
  Resume Exit_DDE
Exit_DDE:
End Sub

Private Sub cmdTransfer_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub Form_Load()
    top = Screen.height / 2 - height / 2
    left = Screen.width / 2 - width / 2
'    Me.HelpContextID = Hlp_Transfer_results
    optexcel(1) = True
    optEnglish = True
End Sub

Private Sub Key_Pressed_On_Control(Ascii_Code As Integer)
  Select Case Ascii_Code
    Case 67, 99 'C,c
      cmdCancel_Click
    Case 69, 101 'E,e
      optEnglish = True
    Case 70, 102 'F,f
      optFrench = True
    Case 84, 116 'T,t
      cmdTransfer_Click
    Case 52 '4
      optexcel(0) = True
    Case 53 '5
      optexcel(1) = True
  End Select
End Sub

Private Function Load_Excel() As Integer
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

Private Sub optEnglish_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub optExcel_KeyPress(index As Integer, KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub optFrench_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub txtTemp_LinkError(ErrNum As Integer)
Dim msg As String
  Const OUT_OF_MEMORY = 11, WRONG_FORMAT = 1, TOO_MANY_DESTINATIONS = 7
  Const UPDATE_FAILED = 8

  Select Case ErrNum
   Case OUT_OF_MEMORY
    msg = "Not Enough Memory to perform DDE"
   Case UPDATE_FAILED
    msg = "Could not update data via DDE"
   Case TOO_MANY_DESTINATIONS
    msg = "DDE SOURCE can not handle this many destinations"
   Case Else
    msg = "unexpected DDE Error:" & ErrNum
  End Select
  If ErrNum <> WRONG_FORMAT Then
    MsgBox msg, MB_ICONEXCLAMATION, "DDE Failure"
  End If
End Sub

