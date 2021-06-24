VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serial Number Testing Window"
   ClientHeight    =   7755
   ClientLeft      =   1740
   ClientTop       =   1860
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame2 
      Height          =   2235
      Left            =   120
      TabIndex        =   20
      Top             =   5190
      Width           =   6165
      _Version        =   65536
      _ExtentX        =   10874
      _ExtentY        =   3942
      _StockProps     =   14
      Caption         =   "Test Serial Number Verification (function snumVerify()):"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdMiscTests 
         Caption         =   "Misc. Tests"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1740
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   990
         Width           =   1335
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   1410
         TabIndex        =   8
         Text            =   "txtData(7)"
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton cmdVerifySnum 
         Caption         =   "Verify Serial Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   990
         Width           =   2625
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   1
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "txtOutput(1)"
         Top             =   1620
         Width           =   2415
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Input:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   510
         TabIndex        =   23
         Top             =   405
         Width           =   825
      End
      Begin VB.Label lblDescOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "Output:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2550
         TabIndex        =   22
         Top             =   1665
         Width           =   825
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   270
      Width           =   6165
      _Version        =   65536
      _ExtentX        =   10874
      _ExtentY        =   8281
      _StockProps     =   14
      Caption         =   "Test Serial Number Creation (function snumGenerate()):"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtOutput 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   0
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "txtOutput(0)"
         Top             =   4140
         Width           =   4455
      End
      Begin VB.CommandButton cmdCreateSnum 
         Caption         =   "Create Serial Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1950
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3540
         Width           =   3915
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3450
         TabIndex        =   0
         Text            =   "txtData(0)"
         Top             =   420
         Width           =   2415
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3450
         TabIndex        =   1
         Text            =   "txtData(1)"
         Top             =   855
         Width           =   2415
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3450
         TabIndex        =   2
         Text            =   "txtData(2)"
         Top             =   1275
         Width           =   2415
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3450
         TabIndex        =   3
         Text            =   "txtData(3)"
         Top             =   1695
         Width           =   2415
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3450
         TabIndex        =   4
         Text            =   "txtData(4)"
         Top             =   2115
         Width           =   2415
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   3450
         TabIndex        =   5
         Text            =   "txtData(5)"
         Top             =   2535
         Width           =   2415
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3450
         TabIndex        =   6
         Text            =   "txtData(6)"
         Top             =   2955
         Width           =   2415
      End
      Begin VB.Label lblDescOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "Output:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   510
         TabIndex        =   19
         Top             =   4185
         Width           =   825
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Comma-Delimited iModules:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   17
         Top             =   465
         Width           =   3075
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "iVersionType:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   16
         Top             =   900
         Width           =   3075
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "iExpires:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   15
         Top             =   1320
         Width           =   3075
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "iExpiresDay:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   300
         TabIndex        =   14
         Top             =   1740
         Width           =   3075
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "iExpiresMonth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   300
         TabIndex        =   13
         Top             =   2160
         Width           =   3075
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "iExpiresYear:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   300
         TabIndex        =   12
         Top             =   2580
         Width           =   3075
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "longInternalSnum:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   300
         TabIndex        =   11
         Top             =   3000
         Width           =   3075
      End
   End
   Begin VB.Menu mnuTests 
      Caption         =   "&Tests"
      Begin VB.Menu mnuTestsItem 
         Caption         =   "snumIsModulePurchased"
         Index           =   10
      End
      Begin VB.Menu mnuTestsItem 
         Caption         =   "snumGetExpiration*"
         Index           =   20
      End
      Begin VB.Menu mnuTestsItem 
         Caption         =   "-"
         Index           =   99
      End
      Begin VB.Menu mnuTestsItem 
         Caption         =   "snumCpasLicGenerate"
         Index           =   100
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Const Form1_declarations_end = True


Sub Test_snumCpasLicGenerate()
Dim spCpasDir As String * 300
Dim spWinDir As String * 300
Dim spNumber As String * 100
Dim spUserName As String * 100
Dim spUserCompany As String * 100
Dim RetVal As Integer
Dim Msg As String
  spCpasDir = "c:\etdot10" & Chr$(0)
  spWinDir = "c:\winnt"
  spNumber = Trim$(txtData(7).Text) & Chr$(0)
  spUserName = "Test User" & Chr$(0)
  spUserCompany = "Test Company" & Chr$(0)
  RetVal = snumCpasLicGenerate( _
    spCpasDir, _
    spWinDir, _
    spNumber, _
    spUserName, _
    spUserCompany)
  Msg = Trim$(Str$(RetVal))
  MsgBox "Result of snumCpasLicGenerate test = " & Msg, _
      vbInformation, App.Title
End Sub


Sub Test_snumIsModulePurchased()
Dim RetVal As Integer
Dim spNumber As String * 100
Dim i As Integer
Dim Msg As String
  spNumber = CStr(txtData(7).Text)
  '========== snumIsModulePurchased TESTS =====================================================
  'MsgBox "Performing snumIsModulePurchased tests.  ", _
      vbInformation, App.Title
  Msg = ""
  For i = 0 To 10
    'MsgBox "Performing snumIsModulePurchased tests.  " & _
        "About to call with iModule = " & Trim$(Str$(i)) & _
        ".", vbInformation, App.Title
    RetVal = snumIsModulePurchased(spNumber, CInt(i))
    If (Msg <> "") Then Msg = Msg & ","
    Msg = Msg & Trim$(Str$(RetVal))
  Next i
  MsgBox "Result of snumIsModulePurchased test = " & Msg, _
      vbInformation, App.Title
End Sub


Sub Test_snumGetExpiration_All()
Dim RetVal As Integer
Dim spNumber As String * 100
Dim i As Integer
Dim Msg As String
  spNumber = CStr(txtData(7).Text)
  '========== TESTS =====================================================
  'MsgBox "Performing snumGetExpiration* tests.", _
      vbInformation, App.Title
  Msg = ""
  RetVal = snumIsExpirationPresent(spNumber)
  If (RetVal = 1) Then
    Msg = Msg & "Expiration Date: Present." & vbCrLf
  Else
    Msg = Msg & "Expiration Date: None." & vbCrLf
  End If
  Msg = Msg & vbCrLf
  RetVal = snumGetExpirationDay(spNumber)
  Msg = Msg & "Day = " & Trim$(Str$(RetVal)) & vbCrLf
  RetVal = snumGetExpirationMonth(spNumber)
  Msg = Msg & "Month = " & Trim$(Str$(RetVal)) & vbCrLf
  RetVal = snumGetExpirationYear(spNumber)
  Msg = Msg & "Year = " & Trim$(Str$(RetVal)) & vbCrLf
  MsgBox Msg, vbInformation, App.Title
End Sub


'RETURNS:
'    TRUE = RETURNED SUCCESSFULLY.
'    FALSE = FAILED TO GENERATE THAT SNUM.
Function Call_snumGenerate( _
    out_spNumber As String, _
    in_iModules() As Integer, _
    in_iVersionType As Integer, _
    in_iExpires As Integer, _
    in_iExpiresDay As Integer, _
    in_iExpiresMonth As Integer, _
    in_iExpiresYear As Integer, _
    in_longInternalSnum As Long) _
    As Boolean
Dim spNumber As String * 100
Dim iModules(0 To 49) As Long
Dim iVersionType As Integer
Dim iExpires As Integer
Dim iExpiresDay As Integer
Dim iExpiresMonth As Integer
Dim iExpiresYear As Integer
Dim longInternalSnum As Long
Dim iCheck As Integer
Dim RetVal As Integer
Dim i As Integer
  ''CHANGE DIRECTORIES (IS THIS STILL REQUIRED?).
  'ChDir "X:\etdot10\license\snum\cpaslib_vbtest1\vb5"
  'ChDrive "X:\etdot10\license\snum\cpaslib_vbtest1\vb5"
  'COPY PARAMETERS INTO TRANSFER VARIABLES.
  For i = 0 To 49
    iModules(i) = CLng(in_iModules(i))
  Next i
  iVersionType = CInt(in_iVersionType)
  iExpires = CInt(in_iExpires)
  iExpiresDay = CInt(in_iExpiresDay)
  iExpiresMonth = CInt(in_iExpiresMonth)
  iExpiresYear = CInt(in_iExpiresYear)
  longInternalSnum = CLng(in_longInternalSnum)
  iCheck = 13892
  'MAKE THE DLL CALL.
  ''''Call ChangeDir_Main
  RetVal = snumGenerate( _
    spNumber, _
    iModules(0), _
    iVersionType, _
    iExpires, _
    iExpiresDay, _
    iExpiresMonth, _
    iExpiresYear, _
    longInternalSnum, _
    iCheck)
  'RETURN WHETHER THE CALL WORKED OR NOT.
  If (RetVal = 0) Then
    Call_snumGenerate = False
  Else
    Call_snumGenerate = True
    out_spNumber = ""
    For i = 1 To 100
      If (Mid$(spNumber, i, 1) = Chr$(0)) Then Exit For
      out_spNumber = out_spNumber & Mid$(spNumber, i, 1)
    Next i
  End If
End Function


Private Sub cmdCreateSnum_Click()
Dim spNumber As String
Dim iModules(0 To 49) As Integer
Dim iVersionType As Integer
Dim iExpires As Integer
Dim iExpiresDay As Integer
Dim iExpiresMonth As Integer
Dim iExpiresYear As Integer
Dim longInternalSnum As Long
Dim iCheck As Integer
Dim RetVal As Integer
Dim i As Integer
Dim NumArgs As Integer
Dim ThisArg As String
Dim iThisArg As Integer
  'TRANSFER CONTROL OBJECT CONTENTS INTO VARIABLES.
  For i = 0 To 49
    iModules(i) = 0
  Next i
  NumArgs = Parser_GetNumArgs(",", CStr(txtData(0).Text))
  For i = 1 To NumArgs
    Call Parser_GetArg(",", CStr(txtData(0).Text), i, ThisArg)
    iThisArg = CInt(Val(ThisArg))
    If (iThisArg >= 0) And (iThisArg <= 49) Then
      iModules(iThisArg) = 1
    End If
  Next i
  iVersionType = CInt(Val(txtData(1)))
  iExpires = CInt(Val(txtData(2)))
  iExpiresDay = CInt(Val(txtData(3)))
  iExpiresMonth = CInt(Val(txtData(4)))
  iExpiresYear = CInt(Val(txtData(5)))
  longInternalSnum = CLng(Val(txtData(6)))
  RetVal = Call_snumGenerate( _
    spNumber, _
    iModules(), _
    iVersionType, _
    iExpires, _
    iExpiresDay, _
    iExpiresMonth, _
    iExpiresYear, _
    longInternalSnum)
  ''''MsgBox "Returned from call = " & Trim$(Str$(retVal))
  If (RetVal = True) Then
    txtOutput(0).Text = spNumber
    txtData(7).Text = spNumber
  Else
    txtOutput(0).Text = "Invalid input parameters."
  End If
End Sub


Private Sub cmdMiscTests_Click()
  Me.PopupMenu mnuTests
End Sub

Private Sub cmdVerifySnum_Click()
Dim RetVal As Integer
Dim spNumber As String * 100
  spNumber = CStr(txtData(7).Text)
  RetVal = snumVerify(spNumber)
  txtOutput(1).Text = Trim$(Str$(RetVal))
End Sub


Private Sub Form_Load()
  txtData(0).Text = "0,1,2"
  txtData(1).Text = "1"
  txtData(2).Text = "0"
  txtData(3).Text = "0"
  txtData(4).Text = "0"
  txtData(5).Text = "0"
  txtData(6).Text = "1"
  txtOutput(0).Text = ""
  txtData(7).Text = ""
  txtOutput(1).Text = ""
End Sub


Private Sub mnuTestsItem_Click(Index As Integer)
  Select Case Index
    Case 10:
      Call Test_snumIsModulePurchased
    Case 20:
      Call Test_snumGetExpiration_All
    Case 100:
      Call Test_snumCpasLicGenerate
  End Select
End Sub
