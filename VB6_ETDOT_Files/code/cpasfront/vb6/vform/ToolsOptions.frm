VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmToolsOptions 
   Caption         =   "Global Options"
   ClientHeight    =   3855
   ClientLeft      =   4020
   ClientTop       =   3555
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5490
   Begin TabDlg.SSTab SSTab1 
      Height          =   3195
      Left            =   210
      TabIndex        =   2
      Top             =   90
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5636
      _Version        =   327680
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Display Options"
      TabPicture(0)   =   "ToolsOptions.frx":0000
      Tab(0).ControlCount=   7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDesc(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDesc(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkDisplayUninstalledApplications"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkMinimizeOnApplicationExecution"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkShowDescriptionText"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboView"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboArrange"
      Tab(0).Control(6).Enabled=   0   'False
      TabCaption(1)   =   "Unused"
      TabPicture(1)   =   "ToolsOptions.frx":001C
      Tab(1).ControlCount=   0
      Tab(1).ControlEnabled=   0   'False
      Begin VB.ComboBox cboArrange 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2460
         Width           =   2000
      End
      Begin VB.ComboBox cboView 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1980
         Width           =   2000
      End
      Begin Threed.SSCheck chkShowDescriptionText 
         Height          =   285
         Left            =   420
         TabIndex        =   7
         Top             =   630
         Width           =   3825
         _Version        =   65536
         _ExtentX        =   6747
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Show Descriptive Popup Window?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkMinimizeOnApplicationExecution 
         Height          =   285
         Left            =   420
         TabIndex        =   8
         Top             =   990
         Width           =   3825
         _Version        =   65536
         _ExtentX        =   6747
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Minimize On Application Execution?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkDisplayUninstalledApplications 
         Height          =   285
         Left            =   420
         TabIndex        =   9
         Top             =   1350
         Width           =   3825
         _Version        =   65536
         _ExtentX        =   6747
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Display Un-installed Applications?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Icon Arrange Method:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   6
         Top             =   2520
         Width           =   1995
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Icon View Method:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   5
         Top             =   2040
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
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
      Index           =   1
      Left            =   4050
      TabIndex        =   1
      Top             =   3390
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&OK"
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
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   3390
      Width           =   1245
   End
End
Attribute VB_Name = "frmToolsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempProj As ProjectType

Dim USER_HIT_CANCEL As Boolean




Const frmToolsOptions_declarations_end = 0


Sub frmToolsOptions_DoEdit( _
    out_USER_HIT_CANCEL As Boolean)
  TempProj = NowProj
  frmToolsOptions.Show 1
  out_USER_HIT_CANCEL = USER_HIT_CANCEL
  If (Not USER_HIT_CANCEL) Then
    NowProj = TempProj
  End If
End Sub


Sub Refresh_frmToolsOptions()
Dim new_tag As Integer
Dim i As Integer

  '----------- DISPLAY OPTIONS -------------------------
  'CHECK BOXES.
  chkMinimizeOnApplicationExecution.Tag = TempProj.MinimizeOnApplicationExecution
  chkMinimizeOnApplicationExecution.Value = TempProj.MinimizeOnApplicationExecution
  chkShowDescriptionText.Tag = TempProj.ShowDescriptionText
  chkShowDescriptionText.Value = TempProj.ShowDescriptionText
  chkDisplayUninstalledApplications.Tag = TempProj.DisplayUninstalledApplications
  chkDisplayUninstalledApplications.Value = TempProj.DisplayUninstalledApplications
  'SCROLL BOXES.
  new_tag = 0
  For i = 0 To cboView.ListCount - 1
    If (cboView.ItemData(i) = TempProj.lvGroups_View) Then
      new_tag = i
      Exit For
    End If
  Next i
  cboView.ListIndex = new_tag
  cboView.Tag = new_tag
  new_tag = 0
  For i = 0 To cboArrange.ListCount - 1
    If (cboArrange.ItemData(i) = TempProj.lvGroups_Arrange) Then
      new_tag = i
      Exit For
    End If
  Next i
  cboArrange.ListIndex = new_tag
  cboArrange.Tag = new_tag

End Sub


Private Sub cboArrange_Click()
  If (Val(cboArrange.Tag) <> cboArrange.ListIndex) Then
    TempProj.lvGroups_Arrange = cboArrange.ItemData(cboArrange.ListIndex)
    Call Refresh_frmToolsOptions
  End If
End Sub
Private Sub cboView_Click()
  If (Val(cboView.Tag) <> cboView.ListIndex) Then
    TempProj.lvGroups_View = cboView.ItemData(cboView.ListIndex)
    Call Refresh_frmToolsOptions
  End If
End Sub


Private Sub chkMinimizeOnApplicationExecution_Click(Value As Integer)
Dim ctl As Control
Set ctl = chkMinimizeOnApplicationExecution
  If (Value <> CBool(ctl.Tag)) Then
    TempProj.MinimizeOnApplicationExecution = Value
    Call Refresh_frmToolsOptions
  End If
End Sub
Private Sub chkShowDescriptionText_Click(Value As Integer)
Dim ctl As Control
Set ctl = chkShowDescriptionText
  If (Value <> CBool(ctl.Tag)) Then
    TempProj.ShowDescriptionText = Value
    Call Refresh_frmToolsOptions
  End If
End Sub
Private Sub chkDisplayUninstalledApplications_Click(Value As Integer)
Dim ctl As Control
Set ctl = chkDisplayUninstalledApplications
  If (Value <> CBool(ctl.Tag)) Then
    TempProj.DisplayUninstalledApplications = Value
    Call Refresh_frmToolsOptions
  End If
End Sub


Private Sub cmdExit_Click(Index As Integer)
  Select Case Index
    Case 0:     'OK.
      USER_HIT_CANCEL = False
      Unload Me
      Exit Sub
    Case 1:     'CANCEL.
      USER_HIT_CANCEL = True
      Unload Me
      Exit Sub
  End Select
End Sub


Sub populate_cboView()
  cboView.Clear
  cboView.AddItem "Icon"
  cboView.ItemData(cboView.NewIndex) = 0
  'cboView.AddItem "SmallIcon"
  'cboView.ItemData(cboView.NewIndex) = 1
  cboView.AddItem "List"
  cboView.ItemData(cboView.NewIndex) = 2
  cboView.AddItem "Report"
  cboView.ItemData(cboView.NewIndex) = 3
End Sub
Sub populate_cboArrange()
  cboArrange.Clear
  'cboArrange.AddItem "None"
  'cboArrange.ItemData(cboArrange.NewIndex) = 0
  cboArrange.AddItem "Left"
  cboArrange.ItemData(cboArrange.NewIndex) = 1
  cboArrange.AddItem "Top"
  cboArrange.ItemData(cboArrange.NewIndex) = 2
End Sub


Private Sub Form_Load()
  Call CenterOnForm(Me, frmMain)
  Call populate_cboView
  Call populate_cboArrange
  Call Refresh_frmToolsOptions
End Sub


Private Sub SSCheck1_Click(Value As Integer)
End Sub

Private Sub SSTab1_DblClick()
  Call Refresh_frmToolsOptions
End Sub
