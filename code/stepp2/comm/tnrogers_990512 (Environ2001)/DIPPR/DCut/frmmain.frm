VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form FRMMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "DCUT"
   ClientHeight    =   4320
   ClientLeft      =   1485
   ClientTop       =   1020
   ClientWidth     =   7695
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   29.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4320
   ScaleWidth      =   7695
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin ComctlLib.ProgressBar SSPanel1 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Begin Extraction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.PictureBox Frame3D1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   7155
      TabIndex        =   2
      Top             =   1920
      Width           =   7215
      Begin VB.OptionButton Option3 
         Caption         =   "DIPPR911 Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2925
      End
      Begin VB.OptionButton Option2 
         Caption         =   "DIPPR801 Access Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2925
      End
      Begin VB.OptionButton Option1 
         Caption         =   "DIPPR801 Text Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2925
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   7080
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   3360
         Y1              =   1680
         Y2              =   0
      End
      Begin VB.Label nonelbl 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Data Source Path..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   120
         Width           =   2835
      End
      Begin VB.Label none2lbl 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label destlbl 
         Caption         =   "Data Destination Path..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   4
         Top             =   960
         Width           =   2835
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1800
         Width           =   3015
      End
   End
   Begin VB.Label Panel3D1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                            DCUT                             Data Conversion UTility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   7215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MNUAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FRMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelectedOption As Integer

Private Sub Command1_Click()
   
    Dim i As Integer
    Dim Response As Integer
    Dim SourcePath As String    ' dest is always dbmaster
    
    On Error GoTo CheckCancel:
    
    'Reset gauge
     'FRMMain!SSPanel1.Value = 0
    
    'Set dialog box properties
    'CommonDialog1.DialogTitle = "Database Path"
    'CommonDialog1.Filter = "All (*.*)|*.*|Text (*.TXT)|*.TXT|"
    'CommonDialog1.FilterIndex = 0
    'CommonDialog1.InitDir = "c:\"
    'CommonDialog1.ShowOpen
    'DBTXTFilePath = CommonDialog1.filename
    
    'Set the source path
    
    'Get only the source file path
    'I = Len(DBTXTFilePath)
   ' Do While I >= 1
      ' I = I - 1
      ' If Mid(DBTXTFilePath, I, 1) = "\" Then
         ' DBTXTFilePath = Mid(DBTXTFilePath, 1, I)
        '  FRMMain!Panel3D7 = DBTXTFilePath
        '  GoTo DestPath:
         ' Exit Sub
       'End If
   ' Loop
'make sure we've got the paths
If DestPath = NULLPATH Or SourcePath = NULLPATH Then
    MsgBox ("Make sure that there is a set path for the source and the destination")
    Exit Sub
End If

    'Set dialog box properties
    'CommonDialog1.DialogTitle = "Database Destination Path"
    'CommonDialog1.Filter = "All (*.*)|*.*|Text (*.TXT)|*.TXT|"
    'CommonDialog1.FilterIndex = 0
    'CommonDialog1.InitDir = "c:\"
    'CommonDialog1.ShowOpen
    'DBDestPath = CommonDialog1.filename
    
    'Get only the file path
    'I = Len(DBDestPath)
    'Do While I >= 1
    '   I = I - 1
    '   If Mid(DBDestPath, I, 1) = "\" Then
    '      DBDestPath = Mid(DBDestPath, 1, I)
    '      FRMMain!Panel3D9 = DBDestPath
    '      GoTo DoubleCheck:
    '      Exit Sub
    '   End If
    'Loop
        
    On Error GoTo 0
    
    Call FindTemplate

CheckCancel:
    
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub


Private Sub destlbl_Click()
    On Error Resume Next

    Dim i As Integer

    CommonDialog1.Filter = "Access database {*.mdb} |*.mdb"
    CommonDialog1.filename = "Master.mdb"

    CommonDialog1.CancelError = True
    
    On Error GoTo Cncl
    CommonDialog1.ShowOpen
    DestPath = CommonDialog1.filename
    none2lbl.Caption = DestPath

    For i = Len(DestPath) To 1 Step -1
        If Mid(DestPath, i, 1) = "\" Then
            DestPath = Mid(DestPath, 1, i - 1)
            Exit For
        End If
    Next

Cncl:
    Exit Sub
End Sub

Private Sub Form_Load()
   Panel3D1.Caption = "DCUT" & Chr$(13) & "Data Conversion UTility"
   nonelbl.Caption = "NONE"
   none2lbl.Caption = "NONE"
'   Panel3D9.Caption = ""
   Label1.Caption = ""
  
   Left = 0
   Top = 0

   Option1.Value = True
   SelectedOption = 1
   
   Me.Top = (Screen.Height - Me.Height) / 2
   Me.Left = (Screen.Width - Me.Width) / 2
   ' set paths to NULLpath
   SourcePath = NULLPATH
   DestPath = NULLPATH
   
End Sub

Private Sub Label4_Click()
    Dim i As Integer
    
    On Error Resume Next
    
    If Option1.Value = True Then
        CommonDialog1.Filter = "Data {*.dat} |*.dat"
    ElseIf Option2.Value = True Then
        CommonDialog1.Filter = "Access database {*.mdb} |*.mdb"
        CommonDialog1.filename = "Dippr801db.mdb"
    Else
        CommonDialog1.Filter = "Paradox database {*.db} |*.db"
        CommonDialog1.filename = "valuesf.db"
    End If
    
    On Error GoTo Cncl
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    SourcePath = CommonDialog1.filename
    nonelbl.Caption = SourcePath
        
    For i = Len(SourcePath) To 1 Step -1
        If Mid(SourcePath, i, 1) = "\" Then
            SourcePath = Mid(SourcePath, 1, i - 1)
            Exit For
        End If
    Next
    
Cncl:
    Exit Sub
End Sub

Private Sub MNUAbout_Click()
       
   frmSplash.Show 1
    
End Sub



Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub Option1_Click()
    If SelectedOption = 1 Then
        Exit Sub
    Else
        SelectedOption = 1
        SourcePath = NULLPATH
        nonelbl.Caption = "NONE"
    End If
End Sub

Private Sub Option2_Click()
    If SelectedOption = 2 Then
        Exit Sub
    Else
        SelectedOption = 2
        SourcePath = NULLPATH
        nonelbl.Caption = "NONE"
    End If

End Sub

Private Sub Option3_Click()
    If SelectedOption = 3 Then
        Exit Sub
    Else
        SelectedOption = 3
        SourcePath = NULLPATH
        nonelbl.Caption = "NONE"
    End If

End Sub
