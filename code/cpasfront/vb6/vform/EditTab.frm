VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmEditTab 
   Caption         =   "Tab Properties"
   ClientHeight    =   6195
   ClientLeft      =   615
   ClientTop       =   3285
   ClientWidth     =   8790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8790
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
      TabIndex        =   2
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
      Left            =   7230
      TabIndex        =   1
      Top             =   5730
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
      Left            =   5940
      TabIndex        =   0
      Top             =   5730
      Width           =   1245
   End
   Begin Threed.SSFrame ssframe_IconImage 
      Height          =   3945
      Left            =   240
      TabIndex        =   4
      Top             =   1650
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   6959
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
      Begin VB.ListBox lstBackgroundFiles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   3615
      End
      Begin VB.PictureBox picHolder 
         Height          =   3345
         Left            =   4050
         ScaleHeight     =   3285
         ScaleWidth      =   3915
         TabIndex        =   6
         Top             =   360
         Width           =   3975
         Begin VB.Image img_fn_BackgroundImage 
            Height          =   3165
            Left            =   60
            Stretch         =   -1  'True
            Top             =   60
            Width           =   3795
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
         Left            =   4050
         TabIndex        =   5
         Top             =   120
         Width           =   810
      End
   End
   Begin Threed.SSFrame ssframe_TabBackgroundColor 
      Height          =   915
      Left            =   4770
      TabIndex        =   8
      Top             =   600
      Width           =   3675
      _Version        =   65536
      _ExtentX        =   6482
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Tab Background Color"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmd_TabBackgroundColor 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   1305
      End
      Begin Threed.SSPanel sspanel_TabBackgroundColor 
         Height          =   465
         Left            =   2460
         TabIndex        =   10
         Top             =   330
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDescSample 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sample:"
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
         Left            =   1560
         TabIndex        =   11
         Top             =   420
         Width           =   810
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3900
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Label lblDescStr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tab Name:"
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
      TabIndex        =   3
      Top             =   165
      Width           =   1500
   End
End
Attribute VB_Name = "frmEditTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempProj As ProjectType
Dim TempTab As TabType

Dim EDIT_TAB As Integer
Dim ORIGINAL_NAME As String
Dim USER_HIT_CANCEL As Boolean




Const frmEditTab_declarations_end = 0


Sub frmEditTab_DoEdit( _
    in_EDIT_TAB As Integer, _
    out_USER_HIT_CANCEL As Boolean)
  EDIT_TAB = in_EDIT_TAB
  TempProj = NowProj
  TempTab = TempProj.Tabs(EDIT_TAB)
  ORIGINAL_NAME = TempTab.Name
  frmEditTab.Show 1
  out_USER_HIT_CANCEL = USER_HIT_CANCEL
  If (Not USER_HIT_CANCEL) Then
    TempProj.Tabs(EDIT_TAB) = TempTab
    NowProj = TempProj
  End If
End Sub


Sub populate_lstBackgroundFiles()
Dim fn_this As String
Dim fspec As String
  lstBackgroundFiles.Clear
  lstBackgroundFiles.AddItem "( None )"
  fspec = fpath_Backgrounds & "\*.bmp"
  fn_this = Dir(fspec)
  Do While (fn_this <> "")
    lstBackgroundFiles.AddItem LCase$(fn_this)
    fn_this = Dir
  Loop
  fspec = fpath_Backgrounds & "\*.jpg"
  fn_this = Dir(fspec)
  Do While (fn_this <> "")
    lstBackgroundFiles.AddItem LCase$(fn_this)
    fn_this = Dir
  Loop
  fspec = fpath_Backgrounds & "\*.gif"
  fn_this = Dir(fspec)
  Do While (fn_this <> "")
    lstBackgroundFiles.AddItem LCase$(fn_this)
    fn_this = Dir
  Loop
End Sub


Sub Refresh_frmEditTab()
Dim fn_This_Image As String
Dim i As Integer
Dim new_tag As Integer
  'UPDATE TEXT VALUES TO WINDOW.
  Call AssignTextAndTag(txtDataStr(0), Trim$(TempTab.Name))
  'DISPLAY ICON IMAGE.
  fn_This_Image = fpath_Backgrounds & "\" & TempTab.fn_BackgroundImage
  If (Trim$(TempTab.fn_BackgroundImage) = "") Then
    Set img_fn_BackgroundImage = LoadPicture("")
  Else
    Set img_fn_BackgroundImage = LoadPicture(fn_This_Image)
  End If
  'SELECT CURRENT ICON FROM LIST.
  new_tag = 0
  For i = 0 To lstBackgroundFiles.ListCount - 1
    If (Trim$(UCase$(lstBackgroundFiles.List(i))) = _
        Trim$(UCase$(TempTab.fn_BackgroundImage))) Then
      new_tag = i
      Exit For
    End If
  Next i
  lstBackgroundFiles.ListIndex = new_tag
  'DISPLAY COLORS.
  sspanel_TabBackgroundColor.BackColor = TempTab.TabBackgroundColor.Color

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


Private Sub cmd_TabBackgroundColor_Click()
  CommonDialog1.Color = TempTab.TabBackgroundColor.Color
  CommonDialog1.ShowColor
  TempTab.TabBackgroundColor.Color = CommonDialog1.Color
  Call Refresh_frmEditTab
End Sub


Private Sub cmdExit_Click(Index As Integer)
  Select Case Index
    Case 0:     'OK.
      If (Trim$(UCase$(ORIGINAL_NAME)) <> Trim$(UCase$(TempTab.Name))) Then
        If (Tab_IsNameExist(TempProj, TempTab.Name)) Then
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
  Call populate_lstBackgroundFiles
  If (Trim$(UCase$(TempTab.Name)) = _
      Trim$(UCase$("CPAS Main Tools"))) Then
    txtDataStr(0).Locked = True
  Else
    txtDataStr(0).Locked = False
  End If
  Call Refresh_frmEditTab
End Sub


Private Sub lstBackgroundFiles_Click()
  If (lstBackgroundFiles.ListIndex = 0) Then
    'NO IMAGE.
    TempTab.fn_BackgroundImage = ""
  Else
    'STORE IMAGE FILENAME.
    TempTab.fn_BackgroundImage = _
        Trim$(UCase$(lstBackgroundFiles.List(lstBackgroundFiles.ListIndex)))
  End If
  Call Refresh_frmEditTab
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
      Case 0: TempTab.Name = Trim$(txtctl.Text)
    End Select
    Call AssignTextAndTag(txtctl, txtctl.Text)
    'THROW DIRTY FLAG AND REFRESH WINDOW.
    'Call frmBed_DirtyFlag_Throw
    Call Refresh_frmEditTab
  End If
  Call Global_LostFocus(txtctl)
End Sub


