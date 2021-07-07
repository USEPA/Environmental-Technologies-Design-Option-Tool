VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmbreak 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Results for the Plug-Flow Pore Diffusion Model"
   ClientHeight    =   6690
   ClientLeft      =   135
   ClientTop       =   675
   ClientWidth     =   9390
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6690
   ScaleWidth      =   9390
   Begin Threed.SSPanel SSPanel1 
      Height          =   3975
      Left            =   120
      TabIndex        =   41
      Top             =   1920
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   7011
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picBreak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3705
         ScaleWidth      =   7185
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   120
         Width           =   7215
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   495
      Left            =   7920
      TabIndex        =   29
      Top             =   240
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Exit"
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   8760
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSFrame frame3D4 
      Height          =   735
      Left            =   7680
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   "Display time in:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboTime 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
   End
   Begin Threed.SSFrame frame3D3 
      Height          =   735
      Left            =   7680
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   "Grid Style:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboGrid 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
   End
   Begin Threed.SSFrame frame3D2 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   "C/Co as a function of:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optType 
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   37
         Top             =   240
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&BVT"
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
      Begin Threed.SSOption optType 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   38
         Top             =   240
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Time"
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
      Begin Threed.SSOption optType 
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   39
         Top             =   240
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Treatment C&apacity"
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
   Begin Threed.SSFrame frame3D1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _Version        =   65536
      _ExtentX        =   13573
      _ExtentY        =   2990
      _StockProps     =   14
      Caption         =   "Results for: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboCompo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   4815
      End
      Begin Threed.SSCommand cmdTreat 
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Treatment Objectives"
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Time (days)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BVT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tr. Capacity"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5% of influent conc."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50% of influent conc."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "95% of influent conc."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   5040
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   6240
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   6240
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   6240
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C (mg/L)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   6240
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   2640
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   3840
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   5040
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   6240
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   495
      Left            =   7800
      TabIndex        =   30
      Top             =   4200
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Excel..."
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   495
      Left            =   7800
      TabIndex        =   31
      Top             =   4680
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Save &Curves"
   End
   Begin Threed.SSCommand cmdSelect 
      Height          =   495
      Left            =   7800
      TabIndex        =   32
      Top             =   5160
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Select Printer"
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   495
      Left            =   7800
      TabIndex        =   34
      Top             =   5640
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Print"
   End
   Begin Threed.SSCommand cmdFile 
      Height          =   495
      Left            =   7800
      TabIndex        =   35
      Top             =   6120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Print to &File"
   End
End
Attribute VB_Name = "frmbreak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1

'Dim Treatment_Objective(NumSelectedComponents_PFPDM) As Throughput
'Dim Flag_TO(NumSelectedComponents_PFPDM) As Integer
Dim Filename_Input As String

Const TAB_1 = 10
Const TAB_2 = 20
Const TAB_3 = 30
Const TAB_SAVE_INTERVAL = 10

Private Sub cboCompo_Click()

  Exit Sub

End Sub

Private Sub cboGrid_Click()
'   ThisGraph.GridStyle = cboGrid.ListIndex
 '  ThisGraph.DrawMode = 2
End Sub

Private Sub cboTime_Click()
    Dim i As Integer
    Dim FoundTrue As Integer

    FoundTrue = False
    For i = 0 To 2
        If optType(i).Value = True Then
           FoundTrue = True
           optType(i).Value = False
           optType(i).Value = True
           Exit For
        End If
    Next i

    If Not FoundTrue Then
       optType(0).Value = True
    End If

    TimeUnitsOnGraphs = cboTime.ListIndex

End Sub

Private Sub cmdExcel_Click()
  frmExcel.Show 1
End Sub

Private Sub cmdExcel_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub
Private Sub cmdFile_Click()
  '*'* put some code in here
End Sub
Private Sub cmdFile_KeyPress(KeyAscii As Integer)
  Call Key_Pressed_On_Control(KeyAscii)
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub cmdPrint_Click()
 
Dim Error_Code As Integer, temp As String
Dim i As Integer, H  As Single, W As Single
Dim Eq1 As String, j As Integer, MTZ As String, F As Double

On Error GoTo Print_Error

'---- Print the graph ------------------------
    For i = 1 To Number_Component
     ThisGraph.ThisPoint = i
     ThisGraph.PatternData = i - 1
    Next i

    H = ThisGraph.height
    W = ThisGraph.width

    ThisGraph.visible = False 'Hide it before printing

    If Printer.width < Printer.height Then
      ThisGraph.height = CSng(Printer.height / 2#)
      ThisGraph.width = Printer.width
    Else
      ThisGraph.height = Printer.height
      ThisGraph.width = Printer.width
    End If

    ThisGraph.PrintStyle = 2
    ThisGraph.DrawMode = 5

    ThisGraph.height = H
    ThisGraph.width = W

    ThisGraph.visible = True

    ThisGraph.PrintStyle = 2
    ThisGraph.DrawMode = 2

    Call PrintIonExchange

    Exit Sub

Print_Error:
  Error_Code = Err
  temp = "Error " & Format$(Error_Code, "0") & ": " & Error$(Error_Code)
  MsgBox "An error uccured while printing." & Chr$(13) & temp, MB_ICONEXCLAMATION, App.title
  Resume Exit_Print
Exit_Print:

End Sub

Private Sub cmdPrint_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub cmdSave_Click()
Dim F As Integer, i As Integer, j As Integer, temp As String
Dim Filename_PFS As String
Dim TimeToDisplay As Double
Dim ValueToDisplay As Double
 
On Error GoTo Save_Results_PF_Error

   CMDialog1.filename = ""
   CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.iex)|*.iex"
   CMDialog1.FilterIndex = 2
   CMDialog1.DialogTitle = "Save curves from PFPDM"
'------Begin Modification Hokanson: 12-Aug2000
   CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
   CMDialog1.Action = 2

   F = FileNameIsValid(Filename_PFS, CMDialog1)
   If Not (F) Then Exit Sub

   'Save, T, BVF, Usage rate, C/C0
    F = FreeFile
    Open Filename_PFS For Output As F
    Print #F, "Results file:  PFPDM for Windows - Version " & Format$(NVersion, "0.00")
    Print #F,
    Print #F,

     If cboTime.ListIndex = 0 Then       'min
        Print #F, "Time-min"; Tab(TAB_1); "BVF"; Tab(TAB_2); "UR-m3/kg";
     ElseIf cboTime.ListIndex = 1 Then   's
        Print #F, "Time-s"; Tab(TAB_1); "BVF"; Tab(TAB_2); "UR-m3/kg";
     ElseIf cboTime.ListIndex = 2 Then   'hr
        Print #F, "Time-hr"; Tab(TAB_1); "BVF"; Tab(TAB_2); "UR-m3/kg";
     ElseIf cboTime.ListIndex = 3 Then   'd
        Print #F, "Time-d"; Tab(TAB_1); "BVF"; Tab(TAB_2); "UR-m3/kg";
     End If

    For i = 1 To Results.NComponent
     Print #F, Tab(TAB_3 + (i - 1) * TAB_SAVE_INTERVAL); Trim$(Results.Component(i).Name);
    Next i
    
    Print #F,
    Print #F,

    For i = 1 To Results.NPoints
     If cboTime.ListIndex = 0 Then       'min
        TimeToDisplay = Results.T(i)
     ElseIf cboTime.ListIndex = 1 Then   's
        TimeToDisplay = Results.T(i) * 60#
     ElseIf cboTime.ListIndex = 2 Then   'hr
        TimeToDisplay = Results.T(i) / 60#
     ElseIf cboTime.ListIndex = 3 Then   'd
        TimeToDisplay = Results.T(i) / 60# / 24#
     End If

     Print #F, Format$(TimeToDisplay, GetTheFormat(TimeToDisplay));
     ValueToDisplay = Results.T(i) * 60 * Results.Bed.Flowrate.Value / Results.Bed.length / Pi / (Results.Bed.Diameter / 2) ^ 2
     Print #F, Tab(TAB_1); Format$(ValueToDisplay, GetTheFormat(ValueToDisplay));
     ValueToDisplay = Results.T(i) * 60 * Results.Bed.Flowrate.Value / Results.Bed.Weight
     Print #F, Tab(TAB_2); Format$(ValueToDisplay, GetTheFormat(ValueToDisplay));

     For j = 1 To Results.NComponent
       ValueToDisplay = Results.CP(j, i)
       Print #F, Tab(TAB_3 + (j - 1) * TAB_SAVE_INTERVAL); Format$(ValueToDisplay, GetTheFormat(ValueToDisplay));
     Next j
     Print #F,
    
    Next i
    Close F
    CMDialog1.filename = ""
 Exit Sub

Save_Results_PF_Error:
  If Err = 32755 Then
  Else
    MsgBox "Error occurred trying to save results, please retry.", MB_ICONEXCLAMATION, App.title
  End If
  Resume Exit_Save_Results_PF
Exit_Save_Results_PF:

End Sub

Private Sub cmdSave_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Private Sub cmdSelect_Click()
Dim Error_Code As Integer, temp As String
On Error GoTo Select_Print_Error
  CMDialog1.flags = PD_PRINTSETUP
'------Begin Modification Hokanson: 12-Aug2000
  CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
  CMDialog1.Action = 5
  Exit Sub
Select_Print_Error:
  If Err = 32755 Then
  Else
     Error_Code = Err
     temp = "Error " & Format$(Error_Code, "0") & ": " & Error$(Error_Code)
     MsgBox "An error occured while selecting the printer." & Chr$(13) & temp, MB_ICONEXCLAMATION, App.title
  End If
  Resume Exit_Select_Print
Exit_Select_Print:
End Sub

Private Sub cmdSelect_KeyPress(KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)
End Sub

Private Sub Draw_PFPSDM()
Dim i As Integer, j As Integer
Dim Data_Max As Double, factor As Double, Bottom_Title As String
ReDim X_Values(Number_Points_Max) As Double

'Copy the results
  If optType(0) Then  'Time
     If cboTime.ListIndex = 0 Then       'min
        factor = 1#
        Bottom_Title = "Time (min)"
     ElseIf cboTime.ListIndex = 1 Then   's
        factor = 1# * 60#
        Bottom_Title = "Time (s)"
     ElseIf cboTime.ListIndex = 2 Then   'hr
        factor = 1# / 60#
        Bottom_Title = "Time (hr)"
     ElseIf cboTime.ListIndex = 3 Then   'd
        factor = 1# / 60# / 24#
        Bottom_Title = "Time (d)"
     End If
    
  Else
    If optType(1) Then   'BVF         mn * (mn/s) * (m3/s) / m / (m2) -> dimensionless
      factor = 60# * Results.Bed.Flowrate.Value / Results.Bed.length / Pi / (Results.Bed.Diameter / 2#) ^ 2
      Bottom_Title = "Bed Volumes Treated"
    Else   'Treatment Capacity
      factor = 60# * Results.Bed.Flowrate.Value / Results.Bed.Weight   'mn * (s/mn) * (m3/s) / (kg) -> m3/kg
      Bottom_Title = "m3 treated per kg of resin"
    End If
  End If
  'Results.T(I,1) time is in mn
  'Results.T(I,2) is BVF
  For i = 1 To Number_Points_Max
    X_Values(i) = Results.T(i) * factor
  Next i

     
    'Define Graph
 '   ThisGraph.NumSets = Results.NComponent
 '   ThisGraph.GraphType = 6
  '  ThisGraph.GraphStyle = 4

'    For j = 1 To ThisGraph.NumSets
'     ThisGraph.ThisSet = j
'     ThisGraph.NumPoints = Results.NPoints
'    Next j

'    ThisGraph.AutoInc = 0

'    For j = 1 To ThisGraph.NumSets
'       ThisGraph.ThisSet = j
'       For i = 1 To ThisGraph.NumPoints
'         ThisGraph.ThisPoint = i
'         If Results.CP(j, i) < 0 Then
'           ThisGraph.GraphData = 0#
'         Else
'           ThisGraph.GraphData = Results.CP(j, i)
'         End If
'         ThisGraph.ThisPoint = i
'         ThisGraph.LabelText = ""
'         ThisGraph.ThisPoint = i
'         ThisGraph.XPosData = X_Values(i)
'       Next i
'       ThisGraph.ThisPoint = j
'       ThisGraph.LegendText = Trim$(Results.Component(j).Name)
'       ThisGraph.ThisPoint = j
'       ThisGraph.PatternData = j - 1
'    Next j
'
'    ThisGraph.PatternedLines = 0
'    Data_Max = 0
'    For j = 1 To ThisGraph.NumSets
'      ThisGraph.ThisSet = j
'    For i = 1 To ThisGraph.NumPoints
'     ThisGraph.ThisPoint = i
'       If ThisGraph.GraphData > Data_Max Then
'         Data_Max = ThisGraph.GraphData
'        End If
'       Next i
'    Next j
'    ThisGraph.YAxisMax = (Int(Data_Max * 10# + 1)) / 10#
'    ThisGraph.YAxisTicks = 4
'    ThisGraph.GridStyle = 0
'
'    ThisGraph.YAxisStyle = 2
'    ThisGraph.YAxisMin = 0#
'    ThisGraph.BottomTitle = Bottom_Title
'
''    ThisGraph.LeftTitle = "C/Co"
'    ThisGraph.LeftTitle = "C/Ct"
'    ThisGraph.DrawMode = 2

End Sub

Private Sub Form_Load()
Dim j As Integer, i As Integer
   'Set Window
    top = Screen.height / 2 - height / 2
    left = Screen.width / 2 - width / 2
'      Me.HelpContextID = Hlp_Results_for
    Screen.MousePointer = 11
    cboGrid.AddItem "None"
    cboGrid.AddItem "Horizontal"
    cboGrid.AddItem "Vertical"
    cboGrid.AddItem "Both"
    cboGrid.ListIndex = 0

    For i = 1 To Results.NComponent
       cboCompo.AddItem Trim$(Results.Component(i).Name)
'       Treatment_Objective(I) = Results.ThroughPut_05(I)
'       If Treatment_Objective(I).c <> -1 Then
'         Flag_TO(I) = True
'       Else
'         Flag_TO(I) = False
'       End If
    Next i

    Screen.MousePointer = 0
'    cboCompo.ListIndex = 0

    cboTime.AddItem "min"
    cboTime.AddItem "s"
    cboTime.AddItem "hr"
    cboTime.AddItem "d"
    cboTime.ListIndex = TimeUnitsOnGraphs
    
    Set ThisGraph = New GraphControl
    Set ThisGraph.handle_ctlPicture = picBreak
    Call ThisGraph.CreateGraph("", "", "")


End Sub

Private Sub Key_Pressed_On_Control(Ascii_Code As Integer)
  Select Case Ascii_Code
    Case 64, 97 'A,a
      optType(2) = True
    Case 66, 98 'B,a
      optType(1) = True
    Case 67, 99 'C,c
      cmdSave_Click
    Case 69, 101 'E,e
      cmdExcel_Click
    Case 70, 102 'F,f
      cmdFile_Click
    Case 80, 112 'P,p
      cmdPrint_Click
    Case 83, 115 'S,s
     cmdSelect_Click
    Case 84, 116 'T,t
     optType(0) = True
    Case 88, 120 'X,x
     cmdOK_Click
  End Select
End Sub


Private Sub optType_Click(index As Integer, Value As Integer)
  Call Draw_PFPSDM
End Sub

Private Sub optType_KeyPress(index As Integer, KeyAscii As Integer)
  Call Key_Pressed_On_Control(KeyAscii)
End Sub


'Private Sub PlotResults(proj As Project_Type)
'
'Dim data_x() As Double
'Dim data_y() As Double
'Dim num_rows As Integer
'Dim i As Integer
'
'  '
'  ' REMOVE ALL EXISTING GRAPH DATA.
'  '
'  Call ThisGraph.DeleteAllSeries
'  '
'  ' ADD THE FIRST SERIES.
'  '
'  Select Case proj.plottype
'
'  Case 0
'      num_rows = proj.Predicted_count
'      ReDim data_x(1 To num_rows)
'      ReDim data_y(1 To num_rows)
'
'      For i = 1 To num_rows
'        data_x(i) = proj.Predicted(i).Predicted_Theta: _
'          data_y(i) = proj.Predicted(i).Predicted_E
'      Next i
'
'      Call ThisGraph.AddSeriesData( _
'          "Series Whatever", CLng(num_rows), data_x, data_y, _
'          0, 1#, QBColor(9))
'
'  Case 1
'      num_rows = proj.PredictedDispClosed_count
'      ReDim data_x(1 To num_rows)
'      ReDim data_y(1 To num_rows)
'
'      For i = 1 To num_rows
'        data_x(i) = proj.DispClosed(i).PredictedDispClosed_Theta: _
'          data_y(i) = proj.DispClosed(i).PredictedDispClosed_E
'      Next i
'
'      Call ThisGraph.AddSeriesData( _
'          "Series Whatever", CLng(num_rows), data_x, data_y, _
'          0, 1#, QBColor(9))
'
'  Case 2
'      num_rows = proj.PredictedDispOpen_count
'      ReDim data_x(1 To num_rows)
'      ReDim data_y(1 To num_rows)
'
'      For i = 1 To num_rows
'        data_x(i) = proj.DispOpen(i).PredictedDispOpen_Theta: _
'          data_y(i) = proj.DispOpen(i).PredictedDispOpen_E
'      Next i
'
'      Call ThisGraph.AddSeriesData( _
'          "Series Whatever", CLng(num_rows), data_x, data_y, _
'          0, 1#, QBColor(9))
'
'  End Select
'
'   num_rows = proj.Experimental_count
'    ReDim data_x(1 To num_rows)
'    ReDim data_y(1 To num_rows)
'
'    For i = 1 To num_rows
'      data_x(i) = proj.Experimental(i).Experimental_Theta: _
'        data_y(i) = proj.Experimental(i).Experimental_E
'    Next i
'
'    Call ThisGraph.AddSeriesData( _
'    "Series Whatever", CLng(num_rows), data_x, data_y, _
'    1, 1#, QBColor(12))
'
'
'End Sub


