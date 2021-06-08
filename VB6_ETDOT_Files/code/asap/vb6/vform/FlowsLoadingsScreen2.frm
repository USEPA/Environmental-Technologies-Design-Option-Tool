VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFlowsLoadingsScreen2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flows and Loadings"
   ClientHeight    =   4320
   ClientLeft      =   4245
   ClientTop       =   2865
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   810
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3570
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2610
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3570
      Width           =   1215
   End
   Begin Threed.SSFrame fraFlowsLoadings 
      Height          =   3135
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7429
      _ExtentY        =   5524
      _StockProps     =   14
      Caption         =   "Pick Flow and Loading Parameters to Specify:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optFlowsLoadings 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1350
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Water Loading Rate and Air Loading Rate"
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
      End
      Begin Threed.SSOption optFlowsLoadings 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   870
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Water Flow Rate And Air to Water Ratio"
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
      End
      Begin Threed.SSOption optFlowsLoadings 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   390
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Water Flow Rate and Air Flow Rate"
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
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4200
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   $"FlowsLoadingsScreen2.frx":0000
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
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   2070
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmFlowsLoadingsScreen2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmFlowsLoadingsScreen2.Hide
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim CurrentUserOption As Integer

    For i = 0 To 2
        If optFlowsLoadings(i).Value = True Then
           CurrentUserOption = i
        End If
    Next i

    If CurrentUserOption = UsersFlowsLoadingsOption Then
       cmdCancel_Click
       Exit Sub
    End If

    Call SetUpFlowsLoadingsTextBoxes(CurrentUserOption)

    UsersFlowsLoadingsOption = CurrentUserOption

    frmFlowsLoadingsScreen2.Hide
End Sub

Private Sub Form_Activate()
  Call CenterThisForm(Me)
End Sub

Private Sub Form_Load()
  Call CenterThisForm(Me)
End Sub


