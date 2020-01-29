VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmVarConcentrations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Influent/Effluent Concentrations"
   ClientHeight    =   6450
   ClientLeft      =   1935
   ClientTop       =   3120
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8040
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   7560
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   25
      Top             =   -360
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
      Left            =   6240
      TabIndex        =   24
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   120
      Width           =   1455
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   495
      Left            =   1410
      TabIndex        =   0
      Top             =   60
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   -2147483640
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
   Begin Threed.SSCommand cmdCancel 
      Height          =   495
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Cancel"
      ForeColor       =   -2147483640
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
   Begin VCIF1Lib.F1Book Sheet1 
      Height          =   4275
      Left            =   60
      OleObjectBlob   =   "VarConcentrations.frx":0000
      TabIndex        =   22
      Top             =   1920
      Width           =   6345
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1005
      Left            =   6720
      TabIndex        =   23
      Top             =   2100
      Visible         =   0   'False
      Width           =   2805
      _Version        =   65536
      _ExtentX        =   4948
      _ExtentY        =   1773
      _StockProps     =   14
      Caption         =   "Invisible"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   330
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "A="
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
      Left            =   90
      TabIndex        =   21
      Top             =   615
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   690
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "B="
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
      Left            =   90
      TabIndex        =   19
      Top             =   855
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "C="
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
      Left            =   90
      TabIndex        =   18
      Top             =   1095
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "D="
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
      Left            =   90
      TabIndex        =   17
      Top             =   1335
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "E="
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
      Index           =   4
      Left            =   90
      TabIndex        =   16
      Top             =   1575
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "F="
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
      Index           =   5
      Left            =   4050
      TabIndex        =   15
      Top             =   615
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "G="
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
      Index           =   6
      Left            =   4050
      TabIndex        =   14
      Top             =   855
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "H="
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
      Index           =   7
      Left            =   4050
      TabIndex        =   13
      Top             =   1095
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "I="
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
      Index           =   8
      Left            =   4050
      TabIndex        =   12
      Top             =   1335
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "J="
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
      Index           =   9
      Left            =   4050
      TabIndex        =   11
      Top             =   1575
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   690
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   690
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   690
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   690
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   4650
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   4650
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   4650
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Index           =   8
      Left            =   4650
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Index           =   9
      Left            =   4650
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open ..."
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As ..."
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   190
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&1 Old File #1"
         Index           =   191
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&2 Old File #2"
         Index           =   192
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&3 Old File #3"
         Index           =   193
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&4 Old File #4"
         Index           =   194
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "Cu&t"
         Index           =   0
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Copy"
         Index           =   1
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Paste"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmVarConcentrations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim Shifting As Integer, X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
Dim TempStr As String, saveas  As Integer, Filename_Concentration As String
Dim Temp_Array() As String

Dim UserWantsCancel As Integer



Const frmVarConcentrations_declarations_end = True


Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdOK.Enabled = False
  End If
End Sub


Private Sub ClearGrid()
Dim i As Integer
Dim sserror As Integer
  Screen.MousePointer = 11
  For i = 1 To Sheet1.MaxCol
    Sheet1.ClearRange 1, i, 400, i, F1ClearAll
    'sserror = SSDeleteRange(Sheet1.ss, 1, i, 500, i, 3)
  Next i
  'sserror = SSDeleteTable(sheet1.SS)
  'sserror = SSDeleteRange(sheet1.SS, 1, 1, 500, 500, 1)
  Screen.MousePointer = 0
End Sub


Private Function CountConc(i As Integer, npoints As Integer) As Integer
On Error GoTo Error_In_CountConc
  npoints = 0
  Sheet1.Col = i
  Sheet1.Row = 1
  Do Until Sheet1.Text = "" Or Sheet1.Row = Number_Max_Influent_Points
    npoints = npoints + 1
    Sheet1.Row = Sheet1.Row + 1
  Loop
  If Sheet1.Text <> "" Then npoints = npoints + 1
  CountConc = True
  Exit Function
Error_In_CountConc:
  CountConc = False
  Call Show_Error("Invalid data.")
  Resume Exit_CountConc
Exit_CountConc:
End Function


Private Function Load_Concentrations(OverrideFilename As String) As Boolean
Dim f As Integer, npoints As Integer, i As Integer, J  As Integer
ReDim T(Number_Max_Influent_Points) As Double, C(Number_Compo_Max, Number_Max_Influent_Points) As Double
  Load_Concentrations = False
   On Error GoTo Error_In_Reading:
  If (OverrideFilename = "") Then
    CMDialog1.CancelError = True
    CMDialog1.DialogTitle = "Load Concentrations"
    CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel 4.0 (*.xls)|*.xls"
    CMDialog1.FilterIndex = 2
    CMDialog1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNPathMustExist
    CMDialog1.Action = 1
    If CMDialog1.Filename = "" Then
      Exit Function
    End If
    Filename_Concentration = CMDialog1.Filename
  Else
    Filename_Concentration = OverrideFilename
  End If
  ''''mnuFileItem(2).Enabled = True
  If Right(Filename_Concentration, 3) = "XLS" Then
    'OPEN EXCEL FORMAT.
    Sheet1.ReadFile = Filename_Concentration
  Else
    'OPEN TEXT FORMAT.
    f = FreeFile
    Open Filename_Concentration For Input As f
    Input #f, npoints
    For i = 1 To npoints
      Select Case Number_Component
        Case 1
          Input #f, T(i), C(1, i)
        Case 2
          Input #f, T(i), C(1, i), C(2, i)
        Case 3
          Input #f, T(i), C(1, i), C(2, i), C(3, i)
        Case 4
          Input #f, T(i), C(1, i), C(2, i), C(3, i), C(4, i)
        Case 5
          Input #f, T(i), C(1, i), C(2, i), C(3, i), C(4, i), C(5, i)
        Case 6
          Input #f, T(i), C(1, i), C(2, i), C(3, i), C(4, i), C(5, i), C(6, i)
        Case 7
          Input #f, T(i), C(1, i), C(2, i), C(3, i), C(4, i), C(5, i), C(6, i), C(7, i)
        Case 8
          Input #f, T(i), C(1, i), C(2, i), C(3, i), C(4, i), C(5, i), C(6, i), C(7, i), C(8, i)
        Case 9
          Input #f, T(i), C(1, i), C(2, i), C(3, i), C(4, i), C(5, i), C(6, i), C(7, i), C(8, i), C(9, i)
        Case 10
          Input #f, T(i), C(1, i), C(2, i), C(3, i), C(4, i), C(5, i), C(6, i), C(7, i), C(8, i), C(9, i), C(10, i)
      End Select
    Next i
    Close (f)
    'Sheet1.Row = 1
    'Sheet1.Col = 1
    For i = 1 To npoints
      Sheet1.EntryRC(i, 1) = T(i)
      'Sheet1.Text = T(i)
      'Sheet1.Row = Sheet1.Row + 1
    Next i
    'Sheet1.Col = 1
    For J = 1 To Number_Component
      'Sheet1.Col = Sheet1.Col + 1
      'Sheet1.Row = 1
      For i = 1 To npoints
        Sheet1.EntryRC(i, J + 1) = C(J, i)
        'Sheet1.Text = C(J, i)
        'Sheet1.Row = Sheet1.Row + 1
      Next i
    Next J
  End If
  Load_Concentrations = True
  Exit Function
Error_In_Reading:
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("Load_Concentrations")
  End If
  Close #f
  Resume Exit_Load_Points
Exit_Load_Points:
End Function
Private Function SaveConcentrations() As Integer
Dim f As Integer, npoints As Integer, i As Integer, J As Integer
Dim Stemp As String, temp As String, Error_Code As Integer
Dim PreviousFilename_Concentration As String, temporaryname As String
  On Error GoTo Error_In_SaveConcentrations
  If Not (CountConc(1, npoints)) Then
    SaveConcentrations = False
    Call Show_Error("Invalid data.  No data has been saved.")
    Exit Function
  End If
  If (Trim$(Filename_Concentration) <> "") And Not (saveas) Then GoTo Save_File
  PreviousFilename_Concentration = Filename_Concentration
  CMDialog1.CancelError = True
  CMDialog1.DialogTitle = "Save Concentrations"
  CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel 4.0 (*.xls)|*.xls"
  CMDialog1.FilterIndex = 2
  CMDialog1.flags = _
      cdlOFNOverwritePrompt + _
      cdlOFNPathMustExist
  CMDialog1.Action = 2
  temporaryname = CMDialog1.Filename
  Filename_Concentration = temporaryname
  'If IsValidPath(temporaryname, "C:") And CMDialog1.Filename <> "" Then
  '  temporaryname = Mid$(temporaryname, 1, Len(temporaryname) - 1)
  '  Filename_Concentration = temporaryname
  'Else
  '  Filename_Concentration = PreviousFilename_Concentration
  '  CMDialog1.Filename = ""
  '  MsgBox "No data has been saved.", 64, AppName_For_Display_long
  '  Exit Function
  'End If
Save_File:
  If Right(Filename_Concentration, 4) = ".XLS" Then
    'EXCEL FORMAT.
    Sheet1.Write Filename_Concentration, F1FileExcel4
    'Sheet1.WriteFile = Filename_Concentration
  Else
    'TEXT FORMAT.
    mnuFileItem(2).Enabled = True
    f = FreeFile
    'Sheet1.Col = 1
    'Sheet1.Row = 1
    Open Filename_Concentration For Output As f
    Print #f, Format$(npoints, "0")
    For i = 1 To npoints
      Stemp = Format$(CDbl(Sheet1.EntryRC(i, 1)), "0.0000E+00")
      For J = 1 To Number_Component
        'Sheet1.Col = Sheet1.Col + 1
        Stemp = Stemp & "," & Format$(CDbl(Sheet1.EntryRC(i, J + 1)), "0.0000E+00")
      Next J
      Print #f, Stemp
      'Sheet1.Row = Sheet1.Row + 1
      'Sheet1.Col = 1
    Next i
    Close (f)
    SaveConcentrations = True
  End If
  Exit Function
Error_In_SaveConcentrations:
  SaveConcentrations = False
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("SaveConcentrations")
  End If
  If Err = 13 Then
    Call Show_Error("The data entered are not valid data.")
  End If
  Close #f
  Resume Exit_Save_Points
Exit_Save_Points:
End Function


Private Sub cmdCancel_Click()
  frmConcentrations_cancelled = True
  Unload Me
End Sub
Private Sub cmdOK_Click()
Dim i As Integer, response As Integer
ReDim ndata(24) As Integer
Dim DFlag As Integer, f As Integer
Dim J As Integer, No_Var_Influent As Integer
  No_Var_Influent = False
  If Not (CountConc(1, frmConcentrations_NumPoints)) Then
    Sheet1.SetFocus
    Exit Sub
  End If
  Sheet1.Row = 1
  For i = 1 To frmConcentrations_NumConcs
    Sheet1.Col = i
    If Sheet1.Text = "" Then No_Var_Influent = True
  Next i
  If (No_Var_Influent) Then
    response = MsgBox( _
        "There is no data for the first row." & vbCrLf & _
        "It will be assumed that there is no concentration data.", _
        vbExclamation + vbOKCancel, AppName_For_Display_Long)
    Select Case response
      Case vbOK
        GoTo NoInfluent_Conc
      Case vbCancel
        Sheet1.SetFocus
        Exit Sub
    End Select
  End If
  For J = 1 To frmConcentrations_NumConcs + 1
    Sheet1.Col = J
    ndata(J) = 0
    Sheet1.Row = 1
    Do Until ((Sheet1.Text = "") Or (Sheet1.Row >= Number_Max_Influent_Points))
      Sheet1.Row = Sheet1.Row + 1
      ndata(J) = ndata(J) + 1
    Loop
  Next J
  DFlag = False
  For i = 1 To frmConcentrations_NumConcs + 1
    For J = i + 1 To frmConcentrations_NumConcs + 1
      If (ndata(i) <> ndata(J)) Then DFlag = True
    Next J
  Next i
  If (DFlag) Then
    response = MsgBox( _
        "There is not the same number of data in each column." & _
        vbCrLf & _
        "It will be assumed that there is no concentration data.", _
        vbExclamation + vbOKCancel, AppName_For_Display_Short)
    Select Case response
      Case vbOK
      Case vbCancel
        Sheet1.SetFocus
        Exit Sub
    End Select
  End If
  'Store times
  On Error GoTo Time_Error
  Sheet1.Col = 1
  For i = 1 To frmConcentrations_NumPoints
    Sheet1.Row = i
    frmConcentrations_Times(i) = CDbl(Sheet1.Text) * 24# * 60#    'To convert from days to minutes
    If (i > 1) Then
      If (frmConcentrations_TimeOrderImportant) Then
        If (frmConcentrations_Times(i) <= frmConcentrations_Times(i - 1)) Then GoTo Time_Error2
      End If
    End If
  Next i
  'Store concentrations
  On Error GoTo Conc_Error
  For J = 2 To frmConcentrations_NumConcs + 1
    Sheet1.Col = J
    For i = 1 To frmConcentrations_NumPoints
      Sheet1.Row = i
      frmConcentrations_Concs(J - 1, i) = CDbl(Sheet1.Text)
    Next i
  Next J
  Unload Me
Exit_This_OK:
  frmConcentrations_cancelled = False
  Exit Sub
Time_Error:
  Call Show_Error( _
      "At least one value in time input (Row #" & Format$(i, "0") & _
      ") is not a real number." & vbCrLf & _
      "Change this cell (currently `" & _
      Sheet1.EntryRC(i, 1) & "`) to a number.")
  Resume Exit_This_OK:
Time_Error2:
  Call Show_Error( _
      "Time in row #" & Format$(i, "0") & _
      " is less than time in row #" & Format$(i - 1, "0") & "." & _
      vbCrLf & "Change your times to be in chronological order.")
  Sheet1.SetFocus
  Exit Sub
Conc_Error:
  Call Show_Error("At least one value of concentration (Row# " & _
      Format$(i, "0") & ", Col#" & Format$(J, "0") & _
      ") is not a real number." & vbCrLf & _
      "Change this cell (currently `" & Sheet1.EntryRC(i, J) & _
      "`) to a number.")
  Resume Exit_This_OK:
  Sheet1.SetFocus
NoInfluent_Conc:
  Number_Influent_Points = 0
  Unload Me
End Sub


Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Activate()
  Sheet1.SetFocus
End Sub
Private Sub Form_Load()
Dim i As Integer, J As Integer, TB As String, CB As String
Dim temp As String, LF  As String, C As Integer, SetWidth As Integer
  '-- Startup watch-for-cancel timer
  frmConcentrations_cancelled = True
  UserWantsCancel = False
  ''''Timer1.Enabled = True
  Me.Caption = frmConcentrations_caption
  '-- Initialize last-few-files list for this form
  'xaxaxaNC
  ''Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADSIM, LASTFEW_ADSIM_FRMCONCE_FRM)
  'Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADXDESIGNS, LASTFEW_ADXDESIGNS_FRMCONCE_FRM)
  '
  'POPULATE LAST-FEW-FILES LIST.
  Call OldFileList_Populate( _
      2, _
      Me.mnuFileItem(190), _
      Me.mnuFileItem(191), _
      Me.mnuFileItem(192), _
      Me.mnuFileItem(193), _
      Me.mnuFileItem(194))
  ''Me.HelpContextID = Hlp_Influent_Concentrations
  'mnuEditItem(0) = True
  'mnuEditItem(1) = True
  ''mnuEditItem(2) = False       '--- DON'T DO THIS!
  TB = Chr$(9)
  CB = Chr$(13)
  'set maxcols
  Sheet1.MaxCol = Number_Component + 1
  Sheet1.MaxRow = 400
  'set col headers
  Sheet1.ColText(1) = "Time (" & frmConcentrations_Tunits & ")"
  Sheet1.ColWidth(1) = 12 * 256
  'sserror = SSSetColText(Sheet1.ss, 1, "Time (" & frmConcentrations_Tunits & ")")
  'sserror = SSSetColWidth(Sheet1.ss, 1, 1, (12 * 256), False)
  For i = 1 To Number_Component
    Label1(i - 1).Visible = True
    Label2(i - 1).Visible = True
    Label2(i - 1).Caption = Trim$(Component(i).Name) & " (" & frmConcentrations_Cunits & ")"
    Sheet1.ColText(1 + i) = Chr$((Asc("A")) + i - 1)
    'sserror = SSSetColText(Sheet1.ss, 1 + i, Chr$((Asc("A")) + i - 1))
    ''sserror = SSSetColWidth(sheet1.SS, 1 + i, 1 + i, ((Len(LTrim(RTrim(temp))) + 3) * 256), False)
  Next i
  ' size for amount of chemicals
  If (Number_Component < 5) Then
    Sheet1.Top = 975 + (Number_Component * 255)
  Else
    Sheet1.Top = 975 + (5 * 255)
  End If
  'For I = 0 To Number_Component + 1
  ' grid1.FixedAlignment(I) = 2
  'Next I
  SetWidth = Screen.TwipsPerPixelX * 19
  If (frmConcentrations_NumPoints > 0) Then
    For i = 1 To frmConcentrations_NumPoints
      'CONVERT FROM minutes TO days.
      Sheet1.EntryRC(i, 1) = frmConcentrations_Times(i) / 60# / 24#
      'Sheet1.Row = i
      'Sheet1.Col = 1
      'Sheet1.number = frmConcentrations_Times(i) / 60# / 24#       'Convert form min. to days
      For J = 2 To Number_Component + 1
        Sheet1.EntryRC(i, J) = frmConcentrations_Concs(J - 1, i)
        'Sheet1.Col = J
        'Sheet1.number = frmConcentrations_Concs(J - 1, i)
      Next J
    Next i
  End If
  ''''Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmConcentrations.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmConcentrations.Height / 2)
  Call CenterOnForm(Me, frmMain)
  '
  ' DEMO SETTINGS.
  '
  Call LOCAL___Reset_DemoVersionDisablings
End Sub
Private Sub Form_Resize()
Dim XXX As Long
Dim USE_MARGIN As Long
  If (Me.WindowState = 1) Then
    'CANNOT RESIZE WHEN MINIMIZED; EXIT OUTTA HERE.
    Exit Sub
  End If
  USE_MARGIN = Sheet1.Left
  XXX = Me.ScaleWidth - 2 * USE_MARGIN
  If (XXX < 1000) Then XXX = 1000
  Sheet1.Width = XXX
  XXX = Me.ScaleHeight - Sheet1.Top + USE_MARGIN
  If (XXX < 1000) Then XXX = 1000
  Sheet1.Height = XXX
    'If WindowState <> 0 Then Exit Sub
    'If (Grid1.Left + Grid1.Width) > (cdmEdit(3).Left + cdmEdit(3).Width) Then
    '  Width = Grid1.Left + Grid1.Width + 20 * Screen.TwipsPerPixelX
    'Else
    '  Width = cdmEdit(3).Left + cdmEdit(3).Width + 20 * Screen.TwipsPerPixelX
    'End If
    'If Height > (Grid1.Top + 90 * Screen.TwipsPerPixelY + cmdCancel.Height) Then
    ' Grid1.Height = Height - Grid1.Top - 90 * Screen.TwipsPerPixelY - cmdCancel.Height
    ' cmdCancel.Top = Grid1.Top + Grid1.Height + 15 * Screen.TwipsPerPixelY
    ' cmdOK.Top = Grid1.Top + Grid1.Height + 15 * Screen.TwipsPerPixelY
    'End If
    'Top = Screen.Height / 2 - Height / 2
    'Left = Screen.Width / 2 - Width / 2
'Sheet1.SetFocus
End Sub


Private Sub mnuEditItem_Click(Index As Integer)
  Select Case Index
    Case 0:   'CUT.
      On Error Resume Next
      Sheet1.EditCut
    Case 1:   'COPY.
      On Error Resume Next
      Sheet1.EditCopy
    Case 2:   'PASTE.
      On Error Resume Next
      Sheet1.EditPaste
  End Select
'
'    Case 0  'cut
'      sserror = SSEditCut(Sheet1.ss)
'           mnuEditItem(2) = True
'           Sheet1.SetFocus
'           If sserror <> 0 Then
'              'oops
'           End If
'     Case 1  'copy
'           sserror = SSEditCopy(Sheet1.ss)
'           mnuEditItem(2) = True
'           If sserror <> 0 Then
'              'oops
'           End If
'
'     Case 2 'paste
'           If (CutString()) Then
'           Else
'             MsgBox "Impossible to paste data from the clipboard.", 64, AppName_For_Display_long
'           End If
'           'sserror = SSEditpastevalues(sheet1.SS)
'           'sheet1.SetFocus
'           'If sserror <> 0 Then
'           '   'oops
'           'End If
End Sub


Private Sub mnuFileItem_Click(Index As Integer)
Dim i As Integer, J As Integer, f As Integer
Dim response As Integer
Dim fn_new As String
  Select Case Index
    Case 0  'new
       'save changes?
       response = MsgBox("Do you wish to save the changes?", _
          vbExclamation + vbYesNoCancel, AppName_For_Display_Long)
       If response = vbCancel Then Exit Sub
       If response = vbYes Then
         saveas = False
         f = SaveConcentrations()
       End If
       '-- clear grid
       Call ClearGrid

       '   screen.MousePointer = 11
       '   For i = 1 To sheet1.maxcol
       '         sheet1.Col = i
       '      For j = 1 To sheet1.MaxRow
       '          sheet1.Row = j
       '          sheet1.Text = ""
       '      Next
       '   Next
       '
       '   screen.MousePointer = 0
    Case 1  'open
        'save changes?
       response = MsgBox("Do you wish to save the changes?", _
          vbExclamation + vbYesNoCancel, AppName_For_Display_Long)
       If response = vbCancel Then Exit Sub
       If response = vbYes Then
         saveas = False
         f = SaveConcentrations()
       End If
      Call Load_Concentrations("")
      ''''cd_HomeDir
      'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
      Call OldFileList_Promote( _
          Filename_Concentration, _
          2, _
          Me.mnuFileItem(190), _
          Me.mnuFileItem(191), _
          Me.mnuFileItem(192), _
          Me.mnuFileItem(193), _
          Me.mnuFileItem(194))
      ''''Call LastFewFiles_MoveFilenameToTop(Filename_Concentration)
    Case 2  'save
      saveas = False
      f = SaveConcentrations()
    Case 3 'saveas
      saveas = True
      f = SaveConcentrations()
      'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
      Call OldFileList_Promote( _
          Filename_Concentration, _
          2, _
          Me.mnuFileItem(190), _
          Me.mnuFileItem(191), _
          Me.mnuFileItem(192), _
          Me.mnuFileItem(193), _
          Me.mnuFileItem(194))
      ''''Call LastFewFiles_MoveFilenameToTop(Filename_Concentration)
  End Select

  '---- Handle last-few-files stuff
  If ((Index >= 191) And (Index <= 194)) Then
    '---- save first?
    response = MsgBox("Do you wish to save the changes?", _
       vbExclamation + vbYesNoCancel, AppName_For_Display_Long)
    If response = vbCancel Then Exit Sub
    If response = vbYes Then
      saveas = False
      f = SaveConcentrations()
    End If
    '---- clear grid
    Call ClearGrid
    '---- start open
    fn_new = mnuFileItem(Index).Caption
    fn_new = Right$(fn_new, Len(fn_new) - 5)
    If (Load_Concentrations(fn_new) = False) Then
      'DO NOTHING -- FILE NOT LOADED.
    Else
      mnuFileItem(2).Enabled = True
      'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
      Call OldFileList_Promote( _
          Filename_Concentration, _
          2, _
          Me.mnuFileItem(190), _
          Me.mnuFileItem(191), _
          Me.mnuFileItem(192), _
          Me.mnuFileItem(193), _
          Me.mnuFileItem(194))
      ''''Call LastFewFiles_MoveFilenameToTop(Filename_Concentration)
    End If
  End If
End Sub


'Private Function PasteString(StringToTransfer As String, Row As Integer, Col As Integer) As Integer
'On Error GoTo Error_In_PasteString
'  Sheet1.Row = Row
'  Sheet1.Col = Col
'  Sheet1.Text = StringToTransfer
'  PasteString = True
'  Exit Function
'Error_In_PasteString:
'  PasteString = False
'  Resume Exit_PasteString
'Exit_PasteString:
'End Function
'Private Function CutString() As Integer
'Dim ClipString As String, Length As Integer
'Dim CurrentPosition As Integer, PreviousPosition As Integer, Character As String * 1
'Dim StringToTransfer As String, Row As Integer, Col As Integer
'On Error GoTo Error_In_CutString
'  ClipString = Clipboard.GetText()
'  Length = Len(ClipString)
'  If Length > 0 Then
'    PreviousPosition = 1
'    CurrentPosition = 1
'    Col = 1
'    Row = 1
'    While PreviousPosition <= Length
'      Character = Mid$(ClipString, CurrentPosition, 1)
'      Select Case Asc(Character)
'        Case 10
'          CurrentPosition = CurrentPosition + 1
'          PreviousPosition = CurrentPosition
'        Case 13, 9
'          StringToTransfer = Mid$(ClipString, PreviousPosition, CurrentPosition - PreviousPosition)
'          If Not (PasteString(StringToTransfer, Row, Col)) Then
'            MsgBox "Error while pasting data.", 64, AppName_For_Display_long
'          End If
'          Col = Col Mod (Number_Component + 1) + 1
'          If Col = 1 Then
'           Row = Row + 1
'           If Row > Number_Max_Influent_Points Then GoTo Too_Many_Points
'          End If
'          CurrentPosition = CurrentPosition + 1
'          PreviousPosition = CurrentPosition
'        Case Else
'          CurrentPosition = CurrentPosition + 1
'          Character = Mid$(ClipString, CurrentPosition, 1)
'      End Select
'    Wend
'  Else
'  End If
'  CutString = True
'  Exit Function
'Too_Many_Points:
'  CutString = True
'  MsgBox "Too much data was selected. Only the first " & Format$(Number_Max_Influent_Points, "0") & " points were pasted.", 64, AppName_For_Display_long
'  GoTo Exit_CutString
'Error_In_CutString:
'  CutString = False
'  Resume Exit_CutString
'Exit_CutString:
'End Function


