VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEditIcon 
   Caption         =   "Icon Properties"
   ClientHeight    =   7005
   ClientLeft      =   1200
   ClientTop       =   930
   ClientWidth     =   9060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9060
   Begin VB.TextBox txtDataStr 
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
      Left            =   1785
      MaxLength       =   100
      TabIndex        =   1
      Text            =   "txtDataStr(1)"
      Top             =   570
      Width           =   6700
   End
   Begin VB.TextBox txtDataStr 
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
      Height          =   1230
      Index           =   2
      Left            =   1785
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "EditIcon.frx":0000
      Top             =   1020
      Width           =   6700
   End
   Begin VB.TextBox txtDataStr 
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
      Left            =   1770
      MaxLength       =   40
      TabIndex        =   0
      Text            =   "txtDataStr(0)"
      Top             =   120
      Width           =   6700
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
      Left            =   2490
      TabIndex        =   5
      Top             =   6480
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
      Left            =   1200
      TabIndex        =   4
      Top             =   6480
      Width           =   1245
   End
   Begin Threed.SSFrame ssframe_IconImage 
      Height          =   3795
      Left            =   270
      TabIndex        =   7
      Top             =   2430
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   6694
      _StockProps     =   14
      Caption         =   "Icon Image"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstIconFiles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   150
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1170
         Width           =   3225
      End
      Begin VB.PictureBox picHolder 
         Height          =   795
         Left            =   2430
         ScaleHeight     =   735
         ScaleWidth      =   855
         TabIndex        =   11
         Top             =   240
         Width           =   915
         Begin VB.PictureBox pic_fn_IconImage 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   90
            ScaleHeight     =   555
            ScaleWidth      =   675
            TabIndex        =   12
            Top             =   90
            Width           =   675
         End
      End
      Begin VB.Label lblDescSample 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Preview:"
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
         Index           =   2
         Left            =   1530
         TabIndex        =   8
         Top             =   450
         Width           =   810
      End
   End
   Begin Threed.SSFrame ssframe_ApplicationLink 
      Height          =   4425
      Left            =   3840
      TabIndex        =   13
      Top             =   2430
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   7805
      _StockProps     =   14
      Caption         =   "Application Link"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin TabDlg.SSTab SSTab1 
         Height          =   3975
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   7011
         _Version        =   327680
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Target"
         TabPicture(0)   =   "EditIcon.frx":0010
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblNote(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtDataStr(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtTranslation"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Start-In Directory"
         TabPicture(1)   =   "EditIcon.frx":002C
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblNote(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "txtDataStr(4)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "txtTranslation_Dir"
         Tab(1).Control(2).Enabled=   0   'False
         Begin VB.TextBox txtTranslation_Dir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   -74880
            Locked          =   -1  'True
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "EditIcon.frx":0048
            Top             =   2520
            Width           =   4425
         End
         Begin VB.TextBox txtDataStr 
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
            Height          =   750
            Index           =   4
            Left            =   -74880
            MaxLength       =   1000
            TabIndex        =   18
            Text            =   "txtDataStr(4)"
            Top             =   1710
            Width           =   4425
         End
         Begin VB.TextBox txtTranslation 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "EditIcon.frx":006F
            Top             =   2520
            Width           =   4425
         End
         Begin VB.TextBox txtDataStr 
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
            Height          =   750
            Index           =   3
            Left            =   120
            MaxLength       =   1000
            TabIndex        =   15
            Text            =   "txtDataStr(3)"
            Top             =   1710
            Width           =   4425
         End
         Begin VB.Label lblNote 
            BackStyle       =   0  'Transparent
            Caption         =   $"EditIcon.frx":0096
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   1
            Left            =   -74880
            TabIndex        =   20
            Top             =   420
            Width           =   4365
         End
         Begin VB.Label lblNote 
            BackStyle       =   0  'Transparent
            Caption         =   $"EditIcon.frx":0122
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   420
            Width           =   4365
         End
      End
   End
   Begin VB.Label lblDescStr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Longer Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   10
      Top             =   615
      Width           =   1500
   End
   Begin VB.Label lblDescStr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   9
      Top             =   1065
      Width           =   1500
   End
   Begin VB.Label lblDescStr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   6
      Top             =   165
      Width           =   1500
   End
End
Attribute VB_Name = "frmEditIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempProj As ProjectType
Dim TempGroup As GroupType
Dim TempIcon As IconType

Dim EDIT_TAB As Integer
Dim EDIT_GROUP As Integer
Dim EDIT_ICON As Integer
Dim ORIGINAL_NAME As String
Dim USER_HIT_CANCEL As Boolean




Const frmEditIcon_declarations_end = 0


Sub frmEditIcon_DoEdit( _
    in_EDIT_TAB As Integer, _
    in_EDIT_GROUP As Integer, _
    in_EDIT_ICON As Integer, _
    out_USER_HIT_CANCEL As Boolean)
  EDIT_TAB = in_EDIT_TAB
  EDIT_GROUP = in_EDIT_GROUP
  EDIT_ICON = in_EDIT_ICON
  TempProj = NowProj
  TempGroup = TempProj.Tabs(EDIT_TAB).Groups(EDIT_GROUP)
  TempIcon = TempProj.Tabs(EDIT_TAB).Groups(EDIT_GROUP).Icons(EDIT_ICON)
  ORIGINAL_NAME = TempIcon.Name
  frmEditIcon.Show 1
  out_USER_HIT_CANCEL = USER_HIT_CANCEL
  If (Not USER_HIT_CANCEL) Then
    TempProj.Tabs(EDIT_TAB).Groups(EDIT_GROUP).Icons(EDIT_ICON) = TempIcon
    NowProj = TempProj
  End If
End Sub


Sub populate_lstIconFiles()
Dim fn_this As String
Dim fspec As String
  lstIconFiles.Clear
  fspec = fpath_Icons & "\*.bmp"
  fn_this = Dir(fspec)
  Do While (fn_this <> "")
    lstIconFiles.AddItem LCase$(fn_this)
    fn_this = Dir
  Loop
  fspec = fpath_Icons & "\*.ico"
  fn_this = Dir(fspec)
  Do While (fn_this <> "")
    lstIconFiles.AddItem LCase$(fn_this)
    fn_this = Dir
  Loop
  



End Sub


Sub Refresh_frmEditIcon()
Dim fn_This_Image As String
Dim i As Integer
Dim new_tag As Integer
  'UPDATE TEXT VALUES TO WINDOW.
  Call AssignTextAndTag(txtDataStr(0), Trim$(TempIcon.Name))
  Call AssignTextAndTag(txtDataStr(1), Trim$(TempIcon.LongName))
  Call AssignTextAndTag(txtDataStr(2), Trim$(TempIcon.DescriptionText))
  Call AssignTextAndTag(txtDataStr(3), Trim$(TempIcon.fn_ApplicationLink))
  Call AssignTextAndTag(txtDataStr(4), Trim$(TempIcon.fn_ApplicationLink_Dir))
  'DISPLAY ICON IMAGE.
  fn_This_Image = fpath_Icons & "\" & TempIcon.fn_IconImage
  Set pic_fn_IconImage = LoadPicture(fn_This_Image)
  'SELECT CURRENT ICON FROM LIST.
  new_tag = 0
  For i = 0 To lstIconFiles.ListCount - 1
    If (Trim$(UCase$(lstIconFiles.List(i))) = Trim$(UCase$(TempIcon.fn_IconImage))) Then
      new_tag = i
      Exit For
    End If
  Next i
  lstIconFiles.ListIndex = new_tag

  ''DISPLAY COLORS.
  'sspanel_GroupBackgroundColor.BackColor = TempGroup.GroupBackgroundColor.Color
  'sspanel_GroupForegroundColor.BackColor = TempGroup.GroupForegroundColor.Color
  ''DISPLAY FONT SAMPLES.
  'Call FontInfo_SetFontOnControl(txt_GroupTitleFont, TempGroup.GroupTitleFont)
  'Call FontInfo_SetFontOnControl(txt_GroupIconFont, TempGroup.GroupIconFont)
  ''DISPLAY NUMBERS.
  'Call AssignTextAndTag_WithRange(txtData(0), TempGroup.Pos.Left, 100#, Screen.Width)
  'Call AssignTextAndTag_WithRange(txtData(1), TempGroup.Pos.Top, 100#, Screen.Height)
  'Call AssignTextAndTag_WithRange(txtData(2), TempGroup.Pos.Width, 100#, Screen.Width)
  'Call AssignTextAndTag_WithRange(txtData(3), TempGroup.Pos.Height, 100#, Screen.Height)
End Sub


Private Sub cmdExit_Click(Index As Integer)
  Select Case Index
    Case 0:     'OK.
      If (Trim$(UCase$(ORIGINAL_NAME)) <> Trim$(UCase$(TempIcon.Name))) Then
        If (Icon_IsNameExist(TempProj, TempGroup, TempIcon.Name)) Then
          Call Show_Error("That name already exists.  Choose another name.")
          Exit Sub
        End If
      End If
      USER_HIT_CANCEL = False
      Unload Me
      Exit Sub
    Case 1:     'CANCEL.
      USER_HIT_CANCEL = True
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub Form_Load()
  Call CenterOnForm(Me, frmMain)
  Call populate_lstIconFiles
  Call Refresh_frmEditIcon
End Sub


Private Sub SSFrame1_Click()

End Sub


Private Sub Label1_Click()

End Sub

Private Sub lstIconFiles_Click()
  TempIcon.fn_IconImage = _
      Trim$(UCase$(lstIconFiles.List(lstIconFiles.ListIndex)))
  Call Refresh_frmEditIcon
End Sub


Private Sub txtDataStr_Change(Index As Integer)
  If (Index = 3) Then
    txtTranslation.Text = _
        "Currently, this translates as:" & vbCrLf & _
        String_PrepareForApplicationLaunch(txtDataStr(3).Text)
  End If
  If (Index = 4) Then
    txtTranslation_Dir.Text = _
        "Currently, this translates as:" & vbCrLf & _
        String_PrepareForApplicationLaunch(txtDataStr(4).Text)
  End If
End Sub


Private Sub txtDataStr_GotFocus(Index As Integer)
Dim ctl As Control
Set ctl = txtDataStr(Index)
  Call Global_GotFocus(ctl)
End Sub
Private Sub txtDataStr_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtDataStr_LostFocus(Index As Integer)
Dim txtctl As Control
Set txtctl = txtDataStr(Index)
Dim ok_to_save As Integer
Dim refresh_type As Integer
  ok_to_save = False
  If (txtctl.Text <> txtctl.Tag) Then
    ok_to_save = True
  End If
  If (ok_to_save) Then
    'DATA LOOKS OKAY, LET'S GO AHEAD AND SAVE IT.
    refresh_type = 1
    Select Case Index
      Case 0: TempIcon.Name = Trim$(txtctl.Text)
      Case 1: TempIcon.LongName = Trim$(txtctl.Text)
      Case 2: TempIcon.DescriptionText = Trim$(txtctl.Text)
      Case 3: TempIcon.fn_ApplicationLink = Trim$(txtctl.Text)
      Case 4: TempIcon.fn_ApplicationLink_Dir = Trim$(txtctl.Text)
    End Select
    Call AssignTextAndTag(txtctl, txtctl.Text)
    'THROW DIRTY FLAG AND REFRESH WINDOW.
    'Call frmBed_DirtyFlag_Throw
    Call Refresh_frmEditIcon
  End If
  Call Global_LostFocus(txtctl)
End Sub


