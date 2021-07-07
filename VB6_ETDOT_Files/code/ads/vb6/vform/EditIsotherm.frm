VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEditIsotherm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Isotherm Database"
   ClientHeight    =   6630
   ClientLeft      =   1935
   ClientTop       =   1935
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9375
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9240
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   29
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6960
      TabIndex        =   28
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   6240
      Width           =   1455
   End
   Begin Threed.SSFrame fraOne 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   10610
      _StockProps     =   14
      Caption         =   "Select a Chemical:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstCompo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4320
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   4215
      End
      Begin Threed.SSOption optSort 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Sort by &Name"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Value           =   -1  'True
      End
      Begin Threed.SSOption optSort 
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   3
         Top             =   270
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "S&ort by CAS number"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   5520
         Width           =   4185
         _Version        =   65536
         _ExtentX        =   7382
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Select Chemic&al"
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
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   1
         Left            =   2250
         TabIndex        =   26
         Top             =   5130
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Find A&gain"
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
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   5130
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Find"
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
      Begin VB.Label lblEmpty_lstCompo 
         Alignment       =   2  'Center
         Caption         =   "No Chemicals Available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         TabIndex        =   24
         Top             =   510
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin Threed.SSFrame fraTwo 
      Height          =   6015
      Left            =   4710
      TabIndex        =   1
      Top             =   150
      Width           =   4515
      _Version        =   65536
      _ExtentX        =   7964
      _ExtentY        =   10610
      _StockProps     =   14
      Caption         =   "No Chemical Selected"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstRange 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Index           =   1
         Left            =   1860
         TabIndex        =   6
         Top             =   1230
         Width           =   2535
      End
      Begin VB.ListBox lstRange 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1230
         Width           =   1575
      End
      Begin VB.Label lblPhase 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPhase"
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
         Height          =   255
         Left            =   1740
         TabIndex        =   23
         Top             =   4545
         Width           =   2655
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(mg/g)*(L/mg)^(1/n)"
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
         Height          =   255
         Index           =   0
         Left            =   2460
         TabIndex        =   22
         Top             =   330
         Width           =   1800
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblValue(3)"
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
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Width           =   4275
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Source:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblValue(2)"
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
         Height          =   255
         Index           =   2
         Left            =   1740
         TabIndex        =   19
         Top             =   3945
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Carbon Type:"
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
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   3960
         Width           =   1515
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "pH Range:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concentration Range (mg/L):"
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
         Height          =   255
         Left            =   1860
         TabIndex        =   16
         Top             =   990
         Width           =   2535
      End
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1/n"
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
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   630
         Width           =   975
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblValue(1)"
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
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   14
         Top             =   615
         Width           =   975
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblValue(0)"
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
         Height          =   255
         Index           =   0
         Left            =   1380
         TabIndex        =   13
         Top             =   315
         Width           =   975
      End
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K"
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
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Phase:"
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
         Height          =   255
         Left            =   420
         TabIndex        =   11
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTemperature"
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
         Height          =   255
         Left            =   1740
         TabIndex        =   10
         Top             =   4245
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature (C):"
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
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   4260
         Width           =   1635
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Label lblComments 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblComments"
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   4275
      End
   End
   Begin VB.Menu mnuChemical 
      Caption         =   "C&hemical"
      Begin VB.Menu mnuChemicalItem 
         Caption         =   "&New"
         Index           =   1
      End
      Begin VB.Menu mnuChemicalItem 
         Caption         =   "&Edit Current"
         Index           =   2
      End
      Begin VB.Menu mnuChemicalItem 
         Caption         =   "&Delete Current"
         Index           =   3
      End
   End
   Begin VB.Menu mnuIsotherm 
      Caption         =   "&Isotherm"
      Begin VB.Menu mnuIsothermItem 
         Caption         =   "&New"
         Index           =   1
      End
      Begin VB.Menu mnuIsothermItem 
         Caption         =   "&Edit Current"
         Index           =   2
      End
      Begin VB.Menu mnuIsothermItem 
         Caption         =   "&Delete Current"
         Index           =   3
      End
      Begin VB.Menu mnuIsothermItem 
         Caption         =   "Delete &All"
         Index           =   4
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmEditIsotherm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FORM_MODE As Integer
Const FORM_MODE_EDIT_DATABASE = 2

Dim DB_Isotherm As Database
Dim Find_String As String

Dim HALT_LSTCOMPO As Boolean
Dim HALT_LSTRANGE As Boolean






Const frmEditIsotherm_declarations_end = True


Sub frmEditIsotherm_EditDatabase()
  On Error GoTo err_frmEditIsotherm_EditDatabase
  'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
  'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
  Set DB_Isotherm = _
      Ws1.OpenDatabase(fn_DB_Isotherm, True, False, _
      ";pwd=" & decrypt_string(Encrypted_User_Password))
  'Set DB_Isotherm = ws1.OpenDatabase(fn_DB_Isotherm)
  FORM_MODE = FORM_MODE_EDIT_DATABASE
  frmEditIsotherm.Show 1
  DB_Isotherm.Close
  Exit Sub
exit_err_frmEditIsotherm_EditDatabase:
  Exit Sub
err_frmEditIsotherm_EditDatabase:
  Call Show_Trapped_Error("frmEditIsotherm_EditDatabase")
  Resume exit_err_frmEditIsotherm_EditDatabase
End Sub
Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    mnuChemicalItem(1).Enabled = False
    mnuChemicalItem(2).Enabled = False
    mnuChemicalItem(3).Enabled = False
    mnuIsothermItem(1).Enabled = False
    mnuIsothermItem(2).Enabled = False
    mnuIsothermItem(3).Enabled = False
    mnuIsothermItem(4).Enabled = False
  End If
End Sub


Sub populate_lstCompo()
Dim Rs1 As Recordset
Dim Current_Criteria As String
Dim SAVE_CURRENT_POSITION As Long
Dim NEW_LISTINDEX As Integer
Dim This_ID As Long
Dim NumRecords As Long
Dim SortCode As String
Dim TempStr As String * 15
Dim Output_Line As String
Dim ThisChemicalName As String
Dim ThisChemicalCAS As String
  On Error GoTo err_populate_lstCompo
  'SAVE CURRENT POSITION.
  If (lstCompo.ListCount > 0) And (lstCompo.ListIndex >= 0) Then
    SAVE_CURRENT_POSITION = lstCompo.ItemData(lstCompo.ListIndex)
  Else
    SAVE_CURRENT_POSITION = -1
  End If
  'SET UP SEARCH CRITERIA.
  If (optSort(0).Value) Then SortCode = "Name, CAS"
  If (optSort(1).Value) Then SortCode = "CAS, Name"
  Current_Criteria = "select * from [Chemicals] " & _
      "order by " & SortCode
  'START SEARCH.
  Set Rs1 = _
      DB_Isotherm.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lstCompo
  'POPULATE LISTBOX.
  lstCompo.Clear
  If (NumRecords = 0) Then
    'NO RECORDS AVAILABLE.
    lstCompo.Visible = False
    lblEmpty_lstCompo.Move lstCompo.Left, lstCompo.Top
    lblEmpty_lstCompo.Visible = True
  Else
    'DISPLAY RECORDS.
    lstCompo.Visible = True
    lblEmpty_lstCompo.Visible = False
    NEW_LISTINDEX = 0
    Do Until Rs1.EOF
      This_ID = Database_Get_Long(Rs1, "Compo ID")
      ThisChemicalName = Database_Get_String(Rs1, "Name")
      ThisChemicalCAS = Database_Get_String(Rs1, "CAS")
      TempStr = ThisChemicalCAS
              'THIS STRING IS ENSURED TO BE 15 CHARACTERS LONG.
      Output_Line = _
          TempStr & _
          " " & _
          ThisChemicalName
      lstCompo.AddItem Output_Line
      lstCompo.ItemData(lstCompo.NewIndex) = This_ID
      If (SAVE_CURRENT_POSITION <> -1) Then
        If (SAVE_CURRENT_POSITION = This_ID) Then
          NEW_LISTINDEX = lstCompo.NewIndex
        End If
      End If
      Rs1.MoveNext
    Loop
    If (lstCompo.ListCount > 0) Then
      HALT_LSTCOMPO = True
      lstCompo.ListIndex = NEW_LISTINDEX
      HALT_LSTCOMPO = False
    End If
  End If
  'CLOSE DATABASE AND EXIT.
  Rs1.Close
  Exit Sub
exit_err_populate_lstCompo:
  Exit Sub
err_populate_lstCompo:
  Call Show_Trapped_Error("populate_lstCompo")
  Resume exit_err_populate_lstCompo
End Sub
Sub populate_lstRange(ThisCAS As String, ThisChemical As String)
Dim PHASE_CODE As Integer
Dim Rs1 As Recordset
Dim Current_Criteria As String
Dim SAVE_CURRENT_POSITION As Long
Dim This_ID As Long
Dim NEW_LISTINDEX As Long
Dim NumRecords As Long
Dim PhaseCode As String
Dim ThisCMin As Double
Dim ThisCMax As Double
Dim ThisPHMin As Double
Dim ThisPHMax As Double
Dim ThisDbl As Double
Dim ThisOutput As String
  On Error GoTo err_populate_lstRange
  ''GET PHASE CODE.
  'Select Case Bed.Phase
  '  Case 0: PhaseCode = "Liquid"
  '  Case 1: PhaseCode = "Gas"
  'End Select
  'SAVE CURRENT POSITION.
  If (lstRange(0).ListCount > 0) And (lstRange(0).ListIndex >= 0) Then
    SAVE_CURRENT_POSITION = lstRange(0).ItemData(lstRange(0).ListIndex)
  Else
    SAVE_CURRENT_POSITION = -1
  End If
  'SET UP SEARCH CRITERIA.
  If (Trim$(ThisCAS) = "0") Then ThisCAS = ""
  If (Trim$(ThisCAS) <> "") Then
    Current_Criteria = "select * from Isotherms" & _
        " where [Component Number] = " & Trim$(ThisCAS) & _
        " order by CarbonName, ID"
  Else
    Current_Criteria = "select * from Isotherms" & _
        " where Name = '" & Trim$(ThisChemical) & "'" & _
        " order by CarbonName, ID"
  End If
  'START SEARCH.
  Set Rs1 = _
      DB_Isotherm.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lstRange
  'POPULATE LISTBOX.
  lstRange(0).Clear
  lstRange(1).Clear
  If (NumRecords = 0) Then
    'NO RECORDS AVAILABLE.
    fraTwo.Caption = "No Isotherms Available."
  Else
    'DISPLAY RECORDS.
    fraTwo.Caption = Trim$(Str$(NumRecords)) & " " & _
        "Isotherm" & _
        IIf(NumRecords = 1, "", "s") & _
        " Available"
    NEW_LISTINDEX = -1
    Do Until Rs1.EOF
      This_ID = Database_Get_Long(Rs1, "ID")
      ThisCMin = Database_Get_Double(Rs1, "C min")
      ThisCMax = Database_Get_Double(Rs1, "C max")
      ThisPHMin = Database_Get_Double(Rs1, "pH min")
      ThisPHMax = Database_Get_Double(Rs1, "pH max")
      If (ThisPHMin = 0#) And (ThisPHMax = 0#) Then
        ThisOutput = "No pH Range"
      Else
        If (ThisPHMin = 0#) Or (ThisPHMax = 0#) Then
          If (ThisPHMin <> 0#) Then ThisDbl = ThisPHMin
          If (ThisPHMax <> 0#) Then ThisDbl = ThisPHMax
          ThisOutput = Format$(ThisDbl, "0.000")
        Else
          ThisOutput = Format$(ThisPHMin, "0.000") & " - " & _
              Format$(ThisPHMax, "0.000")
        End If
      End If
      lstRange(0).AddItem ThisOutput
      lstRange(0).ItemData(lstRange(0).NewIndex) = This_ID
      If (ThisCMin = 0#) And (ThisCMax = 0#) Then
        ThisOutput = "No Conc. Range"
      Else
        If (ThisCMin = 0#) Or (ThisCMax = 0#) Then
          If (ThisCMin <> 0#) Then ThisDbl = ThisCMin
          If (ThisCMax <> 0#) Then ThisDbl = ThisCMax
          ThisOutput = Format$(ThisDbl, "0.000")
        Else
          ThisOutput = Format$(ThisCMin, "0.000") & " - " & _
              Format$(ThisCMax, "0.000")
        End If
      End If
      lstRange(1).AddItem ThisOutput
      lstRange(1).ItemData(lstRange(1).NewIndex) = This_ID
      If (SAVE_CURRENT_POSITION <> -1) Then
        If (SAVE_CURRENT_POSITION = This_ID) Then
          NEW_LISTINDEX = lstRange(0).NewIndex
        End If
      End If
      Rs1.MoveNext
    Loop
    If (lstRange(0).ListCount > 0) And (NEW_LISTINDEX > -1) Then
      lstRange(0).ListIndex = NEW_LISTINDEX
      lstRange(1).ListIndex = NEW_LISTINDEX
    End If
  End If
  'CLOSE DATABASE AND EXIT.
  Rs1.Close
  Exit Sub
exit_err_populate_lstRange:
  Exit Sub
err_populate_lstRange:
  Call Show_Trapped_Error("populate_lstRange")
  Resume exit_err_populate_lstRange
End Sub
Sub populate_lblValue(This_ID As Long)
Dim Rs1 As Recordset
Dim Current_Criteria As String
Dim NumRecords As Long
  On Error GoTo err_populate_lblValue
  'SET UP SEARCH CRITERIA.
  Current_Criteria = "select * from Isotherms" & _
      " where ID = " & Trim$(Str$(This_ID)) & _
      " order by CarbonName"
  'START SEARCH.
  Set Rs1 = _
      DB_Isotherm.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lblValue
  'POPULATE LISTBOX.
  If (NumRecords = 0) Then
    'COULD NOT FIND THAT ISOTHERM (WEIRD PROBLEM).
  Else
    'DISPLAY RECORD.
    Call AssignCaptionAndTag(lblValue(0), Database_Get_Double(Rs1, "K"))
    Call AssignCaptionAndTag(lblValue(1), Database_Get_Double(Rs1, "1/n"))
    Call AssignCaptionAndTag(lblValue(2), Database_Get_String(Rs1, "CarbonName"))
    Call AssignCaptionAndTag(lblTemperature, Database_Get_Double(Rs1, "Tmin"))
    Call AssignCaptionAndTag(lblPhase, Database_Get_String(Rs1, "Phase"))
    Call AssignCaptionAndTag(lblValue(3), Database_Get_String(Rs1, "Source"))
    Call AssignCaptionAndTag(lblComments, Database_Get_String(Rs1, "Comments"))
  End If
  'CLOSE DATABASE AND EXIT.
  Rs1.Close
  Exit Sub
exit_err_populate_lblValue:
  Exit Sub
err_populate_lblValue:
  Call Show_Trapped_Error("populate_lblValue")
  Resume exit_err_populate_lblValue
End Sub


Sub Clear_lblValue()
  lblValue(0) = ""      'ISODB : K.
  lblValue(1) = ""      'ISODB : 1/N.
  lblValue(2) = ""      'ISODB : ADSORBENT TYPE.
  lblTemperature = ""   'ISODB : TEMP.
  lblPhase = ""         'ISODB : PHASE.
  lblValue(3) = ""      'ISODB : SOURCE.
  lblComments = ""      'ISODB : COMMENTS.
End Sub


'Returns:
'- TRUE = Succeeded
'- FALSE = Failed
Function Search_String( _
    J As Integer, _
    ShowErrorMessages As Integer) As Boolean
Dim i As Integer
Dim Res As Integer
  'If (fraIsothermDB.Visible) Then
  '  lstCompo.SetFocus
  'End If
  'For I = J + 1 To lstCompo.ListCount
  For i = J + 1 To lstCompo.ListCount - 1
    Res = InStr(1, lstCompo.List(i), Find_String, 1)
    If (Res > 0) Then
      'NOTE: BY HALTING lstCompo_Click(), THIS ALLOWS THE
      'COMPONENT TO BE SELECTED WITHOUT CLEARING THE ISOTHERM DB
      'VALUES OF K AND 1/N.
      'lstCompo.ListIndex = I
      HALT_LSTCOMPO = True
      Call Do_Select_Component(i)
      HALT_LSTCOMPO = False
      'If (fraIsothermDB.Visible) Then lstCompo.SetFocus
      Search_String = True
      Exit Function
    End If
  Next i
  For i = 0 To J
    Res = InStr(1, lstCompo.List(i), Find_String, 1)
    If (Res > 0) Then
      'NOTE: BY HALTING lstCompo_Click(), THIS ALLOWS THE
      'COMPONENT TO BE SELECTED WITHOUT CLEARING THE ISOTHERM DB
      'VALUES OF K AND 1/N.
      'lstCompo.ListIndex = I
      HALT_LSTCOMPO = True
      Call Do_Select_Component(i)
      HALT_LSTCOMPO = False
      'If (fraIsothermDB.Visible) Then lstCompo.SetFocus
      Search_String = True
      Exit Function
    End If
  Next i
  '----- If not found, show error message: -----
  If (ShowErrorMessages) Then
    Call Show_Error("String Not Found: " & Chr$(34) & _
        Trim$(Find_String) & Chr$(34))
  End If
  Search_String = False
End Function
'Returns:
'- TRUE = Succeeded
'- FALSE = Failed
Function Do_Search_For_Text(ShowErrorMessages As Integer) As Integer
Dim LIST_INDEX As Integer
  LIST_INDEX = lstCompo.ListIndex
  Do_Search_For_Text = Search_String(LIST_INDEX, ShowErrorMessages)
End Function


Private Sub cmdFind_Click(Index As Integer)
Dim NewName As String
Dim USER_HIT_CANCEL As Boolean
  Select Case Index
    Case 0:     'FIND.
      NewName = Find_String
      Do While (1 = 1)
        NewName = frmNewName.frmNewName_GetName( _
            "Search for String", _
            "Enter the string to find:", _
            NewName, _
            USER_HIT_CANCEL)
        If (USER_HIT_CANCEL) Then Exit Sub
        NewName = Trim$(NewName)
        If (NewName <> "") Then Exit Do
        Call Show_Error("You may only enter a non-blank search string.")
      Loop
      Find_String = NewName
      Call Do_Search_For_Text(True)
    Case 1:     'FIND AGAIN.
      Call Do_Search_For_Text(True)
  End Select
End Sub


Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
  'MISC INITS.
  Me.Height = 7395
  Me.Width = 9500
  Call CenterOnForm(Me, frmMain)
  optSort(0).Enabled = True
  optSort(1).Enabled = True
  Find_String = ""
  'RE-POPULATE CHEMICAL LIST.
  Call populate_lstCompo
  ' DEMO SETTINGS.
  Call LOCAL___Reset_DemoVersionDisablings
End Sub


Sub Do_Select_Component(WhichComp As Integer)
''''Dim THIS_ITEMDATA As Long
Dim ThisText As String
Dim ThisCAS As String
Dim ThisChemical As String
  If (WhichComp < 0) Or (lstCompo.ListCount <= 0) Then
    Exit Sub
  End If
  lstCompo.ListIndex = WhichComp
  ''''THIS_ITEMDATA = lstCompo.ItemData(lstCompo.ListIndex)
  'EXTRACT CAS NUMBER AND COMPONENT NAME.
  ThisText = lstCompo.List(WhichComp)
  ThisCAS = Trim$(Left$(ThisText, 15))
  ThisChemical = Trim$(Mid$(ThisText, 16, Len(ThisText) - 15))
  Call populate_lstRange(ThisCAS, ThisChemical)
End Sub

Private Sub lstCompo_Click()
  If (HALT_LSTCOMPO) Then Exit Sub
  HALT_LSTCOMPO = True
  Call Do_Select_Component(lstCompo.ListIndex)
  HALT_LSTCOMPO = False
  'INVALIDATE EXISTING ISOTHERM RECORD LINK (IF ANY).
  'Component(0).IsothermDB_OneOverN = -1#
  'Component(0).IsothermDB_K = -1#
  HALT_LSTRANGE = True
  lstRange(0).ListIndex = -1
  lstRange(1).ListIndex = -1
  HALT_LSTRANGE = False
  'CLEAR EXISTING RECORD DATA.
  Call Clear_lblValue
  'Call frmFreundlich_Refresh
End Sub
Private Sub lstCompo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ((Button And 2) = 2) Then
    Me.PopupMenu mnuChemical
  End If
End Sub


Private Sub lstRange_Click(Index As Integer)
  If (HALT_LSTRANGE) Then Exit Sub
  HALT_LSTRANGE = True
  'KEEP THE RANGE LISTBOXES IN SYNCH.
  Select Case Index
    Case 0: lstRange(1).ListIndex = lstRange(0).ListIndex
    Case 1: lstRange(0).ListIndex = lstRange(1).ListIndex
  End Select
  ''TRANSFER LINK TO COMPONENT(0) STRUCTURE.
  'Component(0).IsothermDB_Component_Name = Trim$(lstCompo.List(lstCompo.ListIndex))
  'Component(0).IsothermDB_Range_Num = lstRange(0).ListIndex
  'DISPLAY ISOTHERM RECORD.
  Call populate_lblValue(lstRange(0).ItemData(lstRange(0).ListIndex))
  ''THROW DIRTY FLAG.
  'Call frmFreundlich_DirtyStatus_Throw
  ''REFRESH WINDOW.
  'Call frmFreundlich_Refresh
  HALT_LSTRANGE = False
End Sub
Private Sub lstRange_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ((Button And 2) = 2) Then
    Me.PopupMenu mnuIsotherm
  End If
End Sub


Private Sub mnuChemicalItem_Click(Index As Integer)
Dim THIS_CHEMICAL_ID As Long
Dim DummyStr1 As String, DummyStr2 As String
Dim DummyBool1 As Boolean, DummyBool2 As Boolean
Dim USER_HIT_CANCEL As Boolean
Dim NewName As String
Dim NewCAS As String
Dim OldName As String
Dim OldCAS As String
'Dim USER_HIT_CANCEL As Boolean
Dim Current_Criteria As String
Dim Rs1 As Recordset
Dim msg As String
Dim RetVal As Integer
Dim i As Integer
Dim Select_Index As Integer
Dim NumRecords As Integer
Dim RecordCount_Chemicals As Integer
Dim RecordCount_Isotherms_CAS As Integer
Dim RecordCount_Isotherms_Name As Integer
  On Error GoTo err_mnuChemicalItem_Click
  If (Index = 2) Or (Index = 3) Then
    If (lstCompo.ListIndex < 0) Or (lstCompo.ListCount = 0) Then
      Call Show_Error("You must first select a chemical.")
      Exit Sub
    End If
    THIS_CHEMICAL_ID = lstCompo.ItemData(lstCompo.ListIndex)
  End If
  Select Case Index
    Case 1:     'new /////////////////////////////////////////////////////////////////////////////////////////////////
      NewName = "New Chemical"
      Do While (1 = 1)
        Call frmEditIsothermCAS.frmEditIsothermCAS_Run( _
            "Creating New Chemical", _
            "New CAS Number", _
            "New Chemical Name", _
            "*", _
            "*", _
            "*", _
            "*", _
            "&Save", _
            USER_HIT_CANCEL, _
            NewCAS, NewName, _
            DummyStr1, DummyStr2, _
            DummyBool1, DummyBool2)
        If (USER_HIT_CANCEL) Then Exit Sub
        NewName = Trim$(NewName)
        NewCAS = Trim$(NewCAS)
        If (NewName <> "") Then Exit Do
        Call Show_Error("The chemical name must be a non-blank string.")
      Loop
      'ADD THE NEW CHEMICAL RECORD.
      Current_Criteria = "select * from [Chemicals]"
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      Rs1.AddNew
      'THE FIELD [Compo ID] IS AUTOMATICALLY UPDATED.
      Rs1("Name") = NewName
      If (NewCAS <> "") Then
        Rs1("CAS") = NewCAS
      Else
        Rs1("CAS") = Null
      End If
      THIS_CHEMICAL_ID = Database_Get_Long(Rs1, "Compo ID")
      Rs1.Update
      Rs1.Close
      'REDISPLAY WINDOW.
      Call populate_lstCompo
      'SELECT THE NEW MANUFACTURER.
      Select_Index = 0
      For i = 0 To lstCompo.ListCount - 1
        If (lstCompo.ItemData(i) = THIS_CHEMICAL_ID) Then
          Select_Index = i
          Exit For
        End If
      Next i
      If (lstCompo.ListCount > 0) Then
        lstCompo.ListIndex = Select_Index
      End If
    Case 2:     'edit current /////////////////////////////////////////////////////////////////////////////////////////////////
      'START THE EDIT PROCESS.
      Current_Criteria = "select * from [Chemicals] where " & _
          "[Compo ID] = " & Trim$(Str$(THIS_CHEMICAL_ID))
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      Rs1.Edit
      'DO USER INPUT.
      OldName = Database_Get_String(Rs1, "Name")
      OldCAS = Database_Get_String(Rs1, "CAS")
      NewName = OldName
      NewCAS = OldCAS
      DummyBool1 = True
      DummyBool2 = True
      Do While (1 = 1)
        Call frmEditIsothermCAS.frmEditIsothermCAS_Run( _
            "Editing a Chemical", _
            "^Current CAS Number", _
            "^Current Chemical Name", _
            "New CAS Number", _
            "New Chemical Name", _
            "Modify all isotherms with the same CAS number", _
            "Modify all isotherms with the same chemical name", _
            "&Save", _
            USER_HIT_CANCEL, _
            OldCAS, OldName, _
            NewCAS, NewName, _
            DummyBool1, DummyBool2)
        If (USER_HIT_CANCEL) Then
          Rs1.CancelUpdate
          Exit Sub
        End If
        NewName = Trim$(NewName)
        NewCAS = Trim$(NewCAS)
        If (NewName = "") Then
          Call Show_Error("The chemical name must be a non-blank string.")
        Else
          If (OldCAS = "") And (NewCAS <> "") And (DummyBool1) Then
            Call Show_Error("You cannot automatically assign CAS numbers " & _
                "to all isotherm records currently without CAS numbers.")
          Else
            Exit Do
          End If
        End If
      Loop
      If (NewCAS <> "") Then
        NewCAS = Format$(CLng(Val(NewCAS)), "0")
      End If
      Rs1("Name") = NewName
      If (NewCAS <> "") Then
        Rs1("CAS") = NewCAS
      Else
        Rs1("CAS") = Null
      End If
      Rs1.Update
      Rs1.Close
      RecordCount_Chemicals = 1
      RecordCount_Isotherms_CAS = 0
      RecordCount_Isotherms_Name = 0
      'MODIFY ISOTHERMS WITH SAME CAS#/CHEMICAL NAME.
      If (DummyBool1) Then
        'MODIFY CAS NUMBER IN ISOTHERM RECORDS.
        If (OldCAS <> "") Then
          Current_Criteria = "select * from [Isotherms] where " & _
              "[Component Number] = " & OldCAS
        Else
          Current_Criteria = "select * from [Isotherms] where " & _
              "[Component Number] = Null"
        End If
        Set Rs1 = _
            DB_Isotherm.OpenRecordset(Current_Criteria)
        On Error Resume Next
        Rs1.MoveFirst
        Rs1.MoveLast
        Rs1.MoveFirst
        NumRecords = Rs1.RecordCount
        On Error GoTo err_mnuChemicalItem_Click
        If (NumRecords = 0) Then
          'DO NOTHING.
        Else
          Do Until Rs1.EOF
            Rs1.Edit
            If (NewCAS <> "") Then
              Rs1("Component Number") = NewCAS
            Else
              Rs1("Component Number") = Null
            End If
            Rs1.Update
            RecordCount_Isotherms_CAS = RecordCount_Isotherms_CAS + 1
            Rs1.MoveNext
          Loop
        End If
        Rs1.Close
      End If
      If (DummyBool2) Then
        'MODIFY CHEMICAL NAME IN ISOTHERM RECORDS.
        Current_Criteria = "select * from [Isotherms] where " & _
            "[Name] = " & Chr$(34) & OldName & Chr$(34)
        Set Rs1 = _
            DB_Isotherm.OpenRecordset(Current_Criteria)
        On Error Resume Next
        Rs1.MoveFirst
        Rs1.MoveLast
        Rs1.MoveFirst
        NumRecords = Rs1.RecordCount
        On Error GoTo err_mnuChemicalItem_Click
        If (NumRecords = 0) Then
          'DO NOTHING.
        Else
          Do Until Rs1.EOF
            Rs1.Edit
            Rs1("Name") = NewName
            Rs1.Update
            RecordCount_Isotherms_Name = RecordCount_Isotherms_Name + 1
            Rs1.MoveNext
          Loop
        End If
        Rs1.Close
      End If
      'REDISPLAY WINDOW.
      Call populate_lstCompo
      Call lstCompo_Click
      'DISPLAY SUMMARY.
      Call Show_Message("Modification Summary:" & vbCrLf & vbCrLf & _
          "Total Chemical Records Changed: " & _
          Trim$(Str$(RecordCount_Chemicals)) & _
          vbCrLf & _
          "CAS Number Modified For: " & _
          Trim$(Str$(RecordCount_Isotherms_CAS)) & " Isotherm Record" & _
          IIf(RecordCount_Isotherms_CAS = 1, "", "s") & _
          vbCrLf & _
          "Chemical Name Modified For: " & Trim$(Str$(RecordCount_Isotherms_Name)) & " Isotherm Record" & _
          IIf(RecordCount_Isotherms_Name = 1, "", "s"))
    Case 3:       'delete current /////////////////////////////////////////////////////////////////////////////////////////////////
      'START THE EDIT PROCESS.
      Current_Criteria = "select * from [Chemicals] where " & _
          "[Compo ID] = " & Trim$(Str$(THIS_CHEMICAL_ID))
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      Rs1.MoveFirst
      'DO USER INPUT.
      NewName = Database_Get_String(Rs1, "Name")
      NewCAS = Database_Get_String(Rs1, "CAS")
      DummyBool1 = True
      DummyBool2 = True
      Do While (1 = 1)
        Call frmEditIsothermCAS.frmEditIsothermCAS_Run( _
            "Deleting a Chemical", _
            "Delete CAS Number", _
            "Delete Chemical Name", _
            "*", _
            "*", _
            "Delete all isotherms with the same CAS number", _
            "Delete all isotherms with the same chemical name", _
            "&Delete", _
            USER_HIT_CANCEL, _
            NewCAS, NewName, _
            DummyStr1, DummyStr2, _
            DummyBool1, DummyBool2)
        If (USER_HIT_CANCEL) Then
          Exit Sub
        End If
        NewName = Trim$(NewName)
        NewCAS = Trim$(NewCAS)
        If (NewName = "") Then
          Call Show_Error("The chemical name must be a non-blank string.")
        Else
          If (NewCAS = "") And (DummyBool1) Then
            Call Show_Error("You cannot automatically delete " & _
                "all isotherm records currently without CAS numbers.")
          Else
            Exit Do
          End If
        End If
      Loop
      If (NewCAS <> "") Then
        NewCAS = Format$(CLng(Val(NewCAS)), "0")
      End If
      If (NewCAS <> "") Then
        Current_Criteria = "select * from [Chemicals] where " & _
            "[Name] = " & Chr$(34) & NewName & Chr$(34) & _
            " and [CAS] = " & NewCAS
      Else
        Current_Criteria = "select * from [Chemicals] where " & _
            "[Name] = " & Chr$(34) & NewName & Chr$(34)
      End If
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      On Error Resume Next
      Rs1.MoveFirst
      Rs1.MoveLast
      Rs1.MoveFirst
      NumRecords = Rs1.RecordCount
      On Error GoTo err_mnuChemicalItem_Click
      RecordCount_Chemicals = 0
      If (NumRecords = 0) Then
        'DO NOTHING.
      Else
        Do Until Rs1.EOF
          Rs1.Delete
          RecordCount_Chemicals = RecordCount_Chemicals + 1
          Rs1.MoveNext
        Loop
      End If
      Rs1.Close
      RecordCount_Isotherms_CAS = 0
      RecordCount_Isotherms_Name = 0
      'DELETE ISOTHERMS WITH SAME CAS#/CHEMICAL NAME.
      If (DummyBool1) Then
        'DELETE ISOTHERM RECORDS WITH THIS CAS NUMBER.
        If (NewCAS <> "") Then
          Current_Criteria = "select * from [Isotherms] where " & _
              "[Component Number] = " & NewCAS
        Else
          Current_Criteria = "select * from [Isotherms] where " & _
              "[Component Number] = Null"
        End If
        Set Rs1 = _
            DB_Isotherm.OpenRecordset(Current_Criteria)
        On Error Resume Next
        Rs1.MoveFirst
        Rs1.MoveLast
        Rs1.MoveFirst
        NumRecords = Rs1.RecordCount
        On Error GoTo err_mnuChemicalItem_Click
        If (NumRecords = 0) Then
          'DO NOTHING.
        Else
          Do Until Rs1.EOF
            Rs1.Delete
            RecordCount_Isotherms_CAS = RecordCount_Isotherms_CAS + 1
            Rs1.MoveNext
          Loop
        End If
        Rs1.Close
      End If
      If (DummyBool2) Then
        'DELETE ISOTHERM RECORDS WITH THIS CHEMICAL NAME.
        Current_Criteria = "select * from [Isotherms] where " & _
            "[Name] = " & Chr$(34) & NewName & Chr$(34)
        Set Rs1 = _
            DB_Isotherm.OpenRecordset(Current_Criteria)
        On Error Resume Next
        Rs1.MoveFirst
        Rs1.MoveLast
        Rs1.MoveFirst
        NumRecords = Rs1.RecordCount
        On Error GoTo err_mnuChemicalItem_Click
        If (NumRecords = 0) Then
          'DO NOTHING.
        Else
          Do Until Rs1.EOF
            Rs1.Delete
            RecordCount_Isotherms_Name = RecordCount_Isotherms_Name + 1
            Rs1.MoveNext
          Loop
        End If
        Rs1.Close
      End If
      'REDISPLAY WINDOW.
      Call populate_lstCompo
      'DISPLAY SUMMARY.
      Call Show_Message("Modification Summary:" & vbCrLf & vbCrLf & _
          "Total Chemical Records Deleted: " & _
          Trim$(Str$(RecordCount_Chemicals)) & _
          vbCrLf & _
          "Isotherm Records Deleted with Matching CAS Number: " & _
          Trim$(Str$(RecordCount_Isotherms_CAS)) & " Isotherm Record" & _
          IIf(RecordCount_Isotherms_CAS = 1, "", "s") & _
          vbCrLf & _
          "Isotherm Records Deleted with Matching Chemical Name: " & Trim$(Str$(RecordCount_Isotherms_Name)) & " Isotherm Record" & _
          IIf(RecordCount_Isotherms_Name = 1, "", "s"))
  End Select
  Exit Sub
exit_err_mnuChemicalItem_Click:
  Exit Sub
err_mnuChemicalItem_Click:
  Call Show_Trapped_Error("mnuChemicalItem_Click")
  Resume exit_err_mnuChemicalItem_Click

End Sub
Private Sub mnuExit_Click()
  Unload Me
  Exit Sub
End Sub
Private Sub mnuIsothermItem_Click(Index As Integer)
Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_SAVE As Boolean
Dim USER_HIT_SAVEASNEW As Boolean
Dim THIS_CHEM_ID As Long
Dim THIS_ISOTHERM_ID As Long
Dim Current_Criteria As String
Dim Rs1 As Recordset
Dim NumRecords As Long
Dim DEFAULT_PHASE_IS_LIQUID As Boolean
Const INDEX_NEW = 1
Const INDEX_EDIT = 2
Const INDEX_DELETE = 3
Const INDEX_DELETE_ALL = 4
Dim NewName As String
Dim msg As String
Dim RetVal As Integer
Dim Select_Index As Integer
Dim i As Integer
Dim ThisName As String
Dim ThisCAS As String
Dim DEFAULT_CHEMICALNAME As String
Dim DEFAULT_CHEMICALCAS As String
Dim IsothermCount As Integer
Dim RecordCount_Deleted As Integer
  On Error GoTo err_mnuIsothermItem_Click
  If (lstCompo.ListIndex < 0) Or (lstCompo.ListCount = 0) Then
    Call Show_Error("You must first select a chemical.")
    Exit Sub
  End If
  THIS_CHEM_ID = lstCompo.ItemData(lstCompo.ListIndex)
  'GET CHEMICAL NAME AND CAS FOR THIS CHEMICAL.
  Current_Criteria = "select * from [Chemicals] " & _
        "where [Compo ID]=" & _
        Trim$(Str$(THIS_CHEM_ID)) & " " & _
        "order by [Name]"
  Set Rs1 = _
      DB_Isotherm.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_mnuIsothermItem_Click
  If (NumRecords = 0) Then
    'NO RECORD(S) AVAILABLE; WEIRD PROBLEM.
    'EXIT SUBROUTINE.
    Exit Sub
  End If
  ThisName = Database_Get_String(Rs1, "Name")
  ThisCAS = Database_Get_String(Rs1, "CAS")
  If (Index = INDEX_EDIT) Or (Index = INDEX_DELETE) Then
    If (lstRange(0).ListIndex < 0) Or (lstRange(0).ListCount = 0) Then
      Call Show_Error("You must first select an isotherm.")
      Exit Sub
    End If
    THIS_ISOTHERM_ID = lstRange(0).ItemData(lstRange(0).ListIndex)
    'POSITION TO CURRENT ISOTHERM RECORD.
    Current_Criteria = "select * from [Isotherms] " & _
        "where [ID]=" & _
        Trim$(Str$(THIS_ISOTHERM_ID)) & " " & _
        "order by [Name]"
    Set Rs1 = _
        DB_Isotherm.OpenRecordset(Current_Criteria)
    On Error Resume Next
    Rs1.MoveFirst
    Rs1.MoveLast
    Rs1.MoveFirst
    NumRecords = Rs1.RecordCount
    On Error GoTo err_mnuIsothermItem_Click
    If (NumRecords = 0) Then
      'NO RECORD(S) AVAILABLE; WEIRD PROBLEM.
      'EXIT SUBROUTINE.
      Exit Sub
    End If
  End If
  Select Case Index
    Case INDEX_NEW:     '//////// NEW ISOTHERM. /////////////////////////////////////////////////////////////////////////////////////
      'DETERMINE DEFAULT PHASE.
      If (Bed.Phase = 0) Then
        DEFAULT_PHASE_IS_LIQUID = True
      Else
        DEFAULT_PHASE_IS_LIQUID = False
      End If
      'ALLOW USER TO ADD NEW RECORD.
      DEFAULT_CHEMICALNAME = ThisName
      DEFAULT_CHEMICALCAS = ThisCAS
      Call frmEditIsothermData.frmEditIsothermData_AddNew( _
          DEFAULT_PHASE_IS_LIQUID, _
          DEFAULT_CHEMICALNAME, _
          DEFAULT_CHEMICALCAS, _
          USER_HIT_CANCEL, _
          USER_HIT_SAVE)
      If (USER_HIT_CANCEL) Then Exit Sub
      'ADD THE NEW ISOTHERM RECORD.
      Current_Criteria = "select * from [Isotherms]"
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      Rs1.AddNew
      'THE FIELD [ID] IS AUTOMATICALLY UPDATED.
      Rs1("Name") = frmEditIsothermData_Record.Name
      Rs1("K") = frmEditIsothermData_Record.k
      Rs1("1/n") = frmEditIsothermData_Record.OneOverN
      Rs1("C min") = frmEditIsothermData_Record.Cmin
      Rs1("C max") = frmEditIsothermData_Record.Cmax
      Rs1("pH min") = frmEditIsothermData_Record.pHmin
      Rs1("pH max") = frmEditIsothermData_Record.pHmax
      Rs1("Source") = frmEditIsothermData_Record.Source
      Rs1("CarbonName") = frmEditIsothermData_Record.CarbonName
      Rs1("Tmin") = frmEditIsothermData_Record.Tmin
      If (Trim$(frmEditIsothermData_Record.CAS) <> "") Then
        Rs1("Component Number") = CDbl(Val(frmEditIsothermData_Record.CAS))
      Else
        Rs1("Component Number") = Null
      End If
      If (frmEditIsothermData_Record.PhaseIsLiquid) Then
        Rs1("Phase") = "Liquid"
      Else
        Rs1("Phase") = "Gas"
      End If
      Rs1("Comments") = frmEditIsothermData_Record.Comments
      THIS_ISOTHERM_ID = Database_Get_Long(Rs1, "ID")
      Rs1.Update
      'CLOSE THE DATABASE.
      Rs1.Close
      'UPDATE WINDOW.
      Call lstCompo_Click
      'SELECT THE NEW ISOTHERM.
      Select_Index = 0
      For i = 0 To lstRange(0).ListCount - 1
        If (lstRange(0).ItemData(i) = THIS_ISOTHERM_ID) Then
          Select_Index = i
          Exit For
        End If
      Next i
      If (lstRange(0).ListCount > 0) Then
        lstRange(0).ListIndex = Select_Index
      End If
    Case INDEX_EDIT:     '//////// EDIT ISOTHERM. //////////////////////////////////////////////////////////////////////////////////////////////
      'TRANSFER DATABASE RECORD FIELDS TO LOCAL MEMORY.
      If (Database_Get_String(Rs1, "Phase") = "Liquid") Then
        frmEditIsothermData_Record.PhaseIsLiquid = True
      Else
        frmEditIsothermData_Record.PhaseIsLiquid = False
      End If
      frmEditIsothermData_Record.Name = Database_Get_String(Rs1, "Name")
      frmEditIsothermData_Record.k = Database_Get_Double(Rs1, "K")
      frmEditIsothermData_Record.OneOverN = Database_Get_Double(Rs1, "1/n")
      frmEditIsothermData_Record.Cmin = Database_Get_Double(Rs1, "C min")
      frmEditIsothermData_Record.Cmax = Database_Get_Double(Rs1, "C max")
      frmEditIsothermData_Record.pHmin = Database_Get_Double(Rs1, "pH min")
      frmEditIsothermData_Record.pHmax = Database_Get_Double(Rs1, "pH max")
      frmEditIsothermData_Record.Source = Database_Get_String(Rs1, "Source")
      frmEditIsothermData_Record.CarbonName = Database_Get_String(Rs1, "CarbonName")
      frmEditIsothermData_Record.Tmin = Database_Get_Double(Rs1, "Tmin")
      frmEditIsothermData_Record.CAS = Database_Get_String(Rs1, "Component Number")
      frmEditIsothermData_Record.Comments = Database_Get_String(Rs1, "Comments")
      'ALLOW USER TO EDIT THIS RECORD.
      Call frmEditIsothermData.frmEditIsothermData_Edit( _
          USER_HIT_CANCEL, _
          USER_HIT_SAVE, _
          USER_HIT_SAVEASNEW)
      If (USER_HIT_CANCEL) Then Exit Sub
      'SAVE THE EDITED ISOTHERM RECORD.
      Current_Criteria = "select * from [Isotherms] " & _
          "where [ID]=" & Trim$(Str$(THIS_ISOTHERM_ID))
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      If (USER_HIT_SAVE) Then
        'MODIFY EXISTING RECORD.
        Rs1.Edit
        'KEEP ORIGINAL [ID] FIELD INTACT.
      End If
      If (USER_HIT_SAVEASNEW) Then
        'GENERATE NEW RECORD.
        Rs1.AddNew
        'THE FIELD [ID] IS AUTOMATICALLY CREATED DURING THE .Update COMMAND.
      End If
      Rs1("Name") = frmEditIsothermData_Record.Name
      Rs1("K") = frmEditIsothermData_Record.k
      Rs1("1/n") = frmEditIsothermData_Record.OneOverN
      Rs1("C min") = frmEditIsothermData_Record.Cmin
      Rs1("C max") = frmEditIsothermData_Record.Cmax
      Rs1("pH min") = frmEditIsothermData_Record.pHmin
      Rs1("pH max") = frmEditIsothermData_Record.pHmax
      Rs1("Source") = frmEditIsothermData_Record.Source
      Rs1("CarbonName") = frmEditIsothermData_Record.CarbonName
      Rs1("Tmin") = frmEditIsothermData_Record.Tmin
      If (Trim$(frmEditIsothermData_Record.CAS) <> "") Then
        Rs1("Component Number") = CDbl(Val(frmEditIsothermData_Record.CAS))
      Else
        Rs1("Component Number") = Null
      End If
      If (frmEditIsothermData_Record.PhaseIsLiquid) Then
        Rs1("Phase") = "Liquid"
      Else
        Rs1("Phase") = "Gas"
      End If
      Rs1("Comments") = frmEditIsothermData_Record.Comments
      Rs1.Update
      'CLOSE THE DATABASE.
      Rs1.Close
      'UPDATE WINDOW.
      Call lstCompo_Click
    Case INDEX_DELETE:          '//////// DELETE ISOTHERM. /////////////////////////////////////////////////////////////////////////
      msg = "Isotherm Selected for Deletion:" & vbCrLf & _
          vbCrLf & _
          "    K = " & NumberToMFBString(Database_Get_Double(Rs1, "K")) & vbCrLf & _
          "    1/n = " & NumberToMFBString(Database_Get_Double(Rs1, "1/n")) & vbCrLf & _
          "    Carbon Name = " & Database_Get_String(Rs1, "CarbonName") & vbCrLf & _
          vbCrLf & _
          "Do you really want to delete this isotherm ?"
      RetVal = MsgBox(msg, vbCritical + vbYesNo, AppName_For_Display_Long)
      If RetVal = vbNo Then Exit Sub
      'PERFORM DELETION.
      Current_Criteria = "select * from [Isotherms] where " & _
          "[ID] = " & Trim$(Str$(THIS_ISOTHERM_ID))
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      Rs1.Delete
      'CLOSE THE DATABASE.
      Rs1.Close
      'UPDATE WINDOW.
      Call lstCompo_Click
    Case INDEX_DELETE_ALL:      '//////// DELETE ALL ISOTHERMS. /////////////////////////////////////////////////////////////////////////
      'DETERMINE ISOTHERM COUNT.
      If (ThisCAS <> "") Then
        Current_Criteria = "select * from [Isotherms] " & _
            "where [Name]=" & Chr$(34) & Trim$(ThisName) & Chr$(34) & _
            " and [Component Number]=" & Trim$(ThisCAS)
      Else
        Current_Criteria = "select * from [Isotherms] " & _
            "where [Name]=" & Chr$(34) & Trim$(ThisName) & Chr$(34) & _
            " and [Component Number]=Null"
      End If
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      On Error Resume Next
      Rs1.MoveFirst
      Rs1.MoveLast
      Rs1.MoveFirst
      NumRecords = Rs1.RecordCount
      On Error GoTo err_mnuIsothermItem_Click
      If (NumRecords = 0) Then
        'NO RECORD(S) AVAILABLE.
        Call Show_Error("There are no isotherms to delete for this chemical and CAS number.")
        Exit Sub
      Else
        IsothermCount = NumRecords
      End If
      'CHECK WITH USER: "ARE YOU SURE?"
      msg = "Isotherms Selected for Deletion:" & vbCrLf & _
          vbCrLf & _
          "    Chemical Name = " & ThisName & vbCrLf & _
          "    CAS = " & ThisCAS & vbCrLf & _
          vbCrLf & _
          "    Total = " & Trim$(Str$(IsothermCount)) & " Isotherm Record" & _
          IIf(IsothermCount = 1, "", "s") & vbCrLf & vbCrLf & _
          "Do you really want to delete these isotherms ?"
      RetVal = MsgBox(msg, vbCritical + vbYesNo, AppName_For_Display_Long)
      If RetVal = vbNo Then Exit Sub
      'PERFORM DELETION.
      'USE CRITERIA "Current_Criteria" FROM ABOVE.
      Set Rs1 = _
          DB_Isotherm.OpenRecordset(Current_Criteria)
      RecordCount_Deleted = 0
      Do Until Rs1.EOF
        Rs1.Delete
        RecordCount_Deleted = RecordCount_Deleted + 1
        Rs1.MoveNext
      Loop
      'CLOSE THE DATABASE.
      Rs1.Close
      'UPDATE WINDOW.
      Call lstCompo_Click
      'DISPLAY SUMMARY.
      Call Show_Message("Modification Summary:" & vbCrLf & vbCrLf & _
          "Total Isotherm Records Deleted: " & _
          Trim$(Str$(RecordCount_Deleted)))
  End Select
  Exit Sub
exit_err_mnuIsothermItem_Click:
  Exit Sub
err_mnuIsothermItem_Click:
  Call Show_Trapped_Error("mnuIsothermItem_Click")
  Resume exit_err_mnuIsothermItem_Click
End Sub


Private Sub optSort_Click(Index As Integer, Value As Integer)
  Call populate_lstCompo
End Sub



