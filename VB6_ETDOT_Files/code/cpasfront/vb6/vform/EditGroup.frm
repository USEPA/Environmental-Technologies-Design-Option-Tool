VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmEditGroup 
   Caption         =   "Group Properties"
   ClientHeight    =   5805
   ClientLeft      =   5955
   ClientTop       =   780
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8880
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2220
      Top             =   5340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin Threed.SSFrame ssframe_GroupBackgroundColor 
      Height          =   915
      Left            =   5040
      TabIndex        =   13
      Top             =   570
      Width           =   3675
      _Version        =   65536
      _ExtentX        =   6482
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Group Background Color"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmd_GroupBackgroundColor 
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
         TabIndex        =   14
         Top             =   330
         Width           =   1305
      End
      Begin Threed.SSPanel sspanel_GroupBackgroundColor 
         Height          =   465
         Left            =   2460
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   420
         Width           =   810
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2235
      Left            =   180
      TabIndex        =   4
      Top             =   570
      Width           =   4725
      _Version        =   65536
      _ExtentX        =   8334
      _ExtentY        =   3942
      _StockProps     =   14
      Caption         =   "Position and Size"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
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
         Left            =   1725
         TabIndex        =   11
         Text            =   "txtData(3)"
         Top             =   1710
         Width           =   1875
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
         Left            =   1725
         TabIndex        =   9
         Text            =   "txtData(2)"
         Top             =   1260
         Width           =   1875
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
         Left            =   1725
         TabIndex        =   7
         Text            =   "txtData(1)"
         Top             =   810
         Width           =   1875
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
         Left            =   1725
         TabIndex        =   5
         Text            =   "txtData(0)"
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
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
         Index           =   3
         Left            =   150
         TabIndex        =   12
         Top             =   1755
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1305
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
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
         Left            =   150
         TabIndex        =   8
         Top             =   855
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
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
         Left            =   150
         TabIndex        =   6
         Top             =   405
         Width           =   1500
      End
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
      Left            =   1920
      MaxLength       =   100
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
      Left            =   7470
      TabIndex        =   1
      Top             =   5340
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
      Left            =   6180
      TabIndex        =   0
      Top             =   5340
      Width           =   1245
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   915
      Left            =   5040
      TabIndex        =   17
      Top             =   1890
      Width           =   3675
      _Version        =   65536
      _ExtentX        =   6482
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Group Foreground Color"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmd_GroupForegroundColor 
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
         TabIndex        =   18
         Top             =   330
         Width           =   1305
      End
      Begin Threed.SSPanel sspanel_GroupForegroundColor 
         Height          =   465
         Left            =   2460
         TabIndex        =   19
         Top             =   330
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   32768
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
         Index           =   1
         Left            =   1560
         TabIndex        =   20
         Top             =   420
         Width           =   810
      End
   End
   Begin Threed.SSFrame ssframe_GroupTitleFont 
      Height          =   1095
      Left            =   180
      TabIndex        =   21
      Top             =   2910
      Width           =   8535
      _Version        =   65536
      _ExtentX        =   15055
      _ExtentY        =   1931
      _StockProps     =   14
      Caption         =   "Group Title Font"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txt_GroupTitleFont 
         Height          =   705
         Left            =   2460
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "EditGroup.frx":0000
         Top             =   270
         Width           =   5955
      End
      Begin VB.CommandButton cmd_GroupTitleFont 
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
         TabIndex        =   22
         Top             =   330
         Width           =   1305
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
         Index           =   2
         Left            =   1560
         TabIndex        =   23
         Top             =   420
         Width           =   810
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1095
      Left            =   180
      TabIndex        =   25
      Top             =   4080
      Width           =   8535
      _Version        =   65536
      _ExtentX        =   15055
      _ExtentY        =   1931
      _StockProps     =   14
      Caption         =   "Group Icon Font"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmd_GroupIconFont 
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
         TabIndex        =   27
         Top             =   330
         Width           =   1305
      End
      Begin VB.TextBox txt_GroupIconFont 
         Height          =   705
         Left            =   2460
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "EditGroup.frx":0013
         Top             =   270
         Width           =   5955
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
         Index           =   3
         Left            =   1560
         TabIndex        =   28
         Top             =   420
         Width           =   810
      End
   End
   Begin VB.Label lblDescStr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name:"
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
      Left            =   345
      TabIndex        =   3
      Top             =   165
      Width           =   1500
   End
End
Attribute VB_Name = "frmEditGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempProj As ProjectType
Dim TempGroup As GroupType

Dim EDIT_TAB As Integer
Dim EDIT_GROUP As Integer
Dim ORIGINAL_NAME As String
Dim USER_HIT_CANCEL As Boolean



Const frmEditGroup_declarations_end = 0


Sub frmEditGroup_DoEdit( _
    in_EDIT_TAB As Integer, _
    in_EDIT_GROUP As Integer, _
    out_USER_HIT_CANCEL As Boolean)
  EDIT_TAB = in_EDIT_TAB
  EDIT_GROUP = in_EDIT_GROUP
  TempProj = NowProj
  TempGroup = TempProj.Tabs(EDIT_TAB).Groups(EDIT_GROUP)
  ORIGINAL_NAME = TempGroup.Name
  frmEditGroup.Show 1
  out_USER_HIT_CANCEL = USER_HIT_CANCEL
  If (Not USER_HIT_CANCEL) Then
    TempProj.Tabs(EDIT_TAB).Groups(EDIT_GROUP) = TempGroup
    NowProj = TempProj
  End If
End Sub


Sub Refresh_frmEditGroup()
  'UPDATE TEXT VALUES TO WINDOW.
  Call AssignTextAndTag(txtDataStr(0), Trim$(TempGroup.Name))
  'DISPLAY COLORS.
  sspanel_GroupBackgroundColor.BackColor = TempGroup.GroupBackgroundColor.Color
  sspanel_GroupForegroundColor.BackColor = TempGroup.GroupForegroundColor.Color
  'DISPLAY FONT SAMPLES.
  Call FontInfo_SetFontOnControl(txt_GroupTitleFont, TempGroup.GroupTitleFont)
  Call FontInfo_SetFontOnControl(txt_GroupIconFont, TempGroup.GroupIconFont)
  'DISPLAY NUMBERS.
  Call AssignTextAndTag_WithRange(txtData(0), TempGroup.Pos.Left, 100#, Screen.Width)
  Call AssignTextAndTag_WithRange(txtData(1), TempGroup.Pos.Top, 100#, Screen.Height)
  Call AssignTextAndTag_WithRange(txtData(2), TempGroup.Pos.Width, 100#, Screen.Width)
  Call AssignTextAndTag_WithRange(txtData(3), TempGroup.Pos.Height, 100#, Screen.Height)
End Sub


Private Sub cmd_GroupBackgroundColor_Click()
  CommonDialog1.Color = TempGroup.GroupBackgroundColor.Color
  CommonDialog1.ShowColor
  TempGroup.GroupBackgroundColor.Color = CommonDialog1.Color
  Call Refresh_frmEditGroup
End Sub
Private Sub cmd_GroupForegroundColor_Click()
  CommonDialog1.Color = TempGroup.GroupForegroundColor.Color
  CommonDialog1.ShowColor
  TempGroup.GroupForegroundColor.Color = CommonDialog1.Color
  Call Refresh_frmEditGroup
End Sub


Private Sub cmd_GroupIconFont_Click()
  Call FontInfo_SetFontOnControl(CommonDialog1, TempGroup.GroupIconFont)
  CommonDialog1.Flags = cdlCFScreenFonts ' Flags property must be set
        ' to cdlCFBoth,               ' cdlCFPrinterFonts,
        ' or cdlCFScreenFonts before          ' using ShowFont method.
  CommonDialog1.ShowFont
  Call FontInfo_GetFontFromControl(CommonDialog1, TempGroup.GroupIconFont)
  Call Refresh_frmEditGroup
End Sub
Private Sub cmd_GroupTitleFont_Click()
  Call FontInfo_SetFontOnControl(CommonDialog1, TempGroup.GroupTitleFont)
  CommonDialog1.Flags = cdlCFScreenFonts ' Flags property must be set
        ' to cdlCFBoth,               ' cdlCFPrinterFonts,
        ' or cdlCFScreenFonts before          ' using ShowFont method.
  CommonDialog1.ShowFont
  Call FontInfo_GetFontFromControl(CommonDialog1, TempGroup.GroupTitleFont)
  Call Refresh_frmEditGroup
End Sub


Private Sub cmdExit_Click(Index As Integer)
  Select Case Index
    Case 0:     'OK.
      If (Trim$(UCase$(ORIGINAL_NAME)) <> Trim$(UCase$(TempGroup.Name))) Then
        If (Group_IsNameExist(TempProj, TempProj.Tabs(EDIT_TAB), TempGroup.Name)) Then
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


Sub populate_txt_GroupXFont()
Dim msg As String
  msg = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & vbCrLf & "abcdefghijklmnopqrstuvwxyz 0123456789"
  txt_GroupTitleFont.Text = msg
  txt_GroupIconFont.Text = msg
End Sub


Private Sub Form_Load()
  Call CenterOnForm(Me, frmMain)
  Call populate_txt_GroupXFont
  Call Refresh_frmEditGroup
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim txtctl As Control
Set txtctl = txtData(Index)
Call DisplayDataEntryError
Call Global_GotFocus(txtctl)
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtData_LostFocus(Index As Integer)
Dim newVal As Double
Dim txtctl As Control
Set txtctl = txtData(Index)
'Dim NowCase As Case_Type
Dim ok_to_save As Integer
Dim xmin As Double
Dim xmax As Double
Dim refresh_type As Integer
  xmin = Val(txtctl.LinkItem)
  xmax = Val(txtctl.DataField)
  ok_to_save = False
  If (ValueHasChanged(txtctl)) Then
    If (IsValidNumber(txtctl, vbDouble)) Then
      newVal = CDbl(txtctl.Text)
      If (newVal < xmin) Or (newVal > xmax) Then
        txtctl.Text = txtctl.Tag
      Else
        If (txtctl.Text <> txtctl.Tag) Then
          ok_to_save = True
        End If
      End If
    Else
      txtctl.Text = txtctl.Tag
    End If
  End If
  If (ok_to_save) Then
    'DATA LOOKS OKAY, LET'S GO AHEAD AND SAVE IT.
    Select Case Index
      Case 0: TempGroup.Pos.Left = newVal
      Case 1: TempGroup.Pos.Top = newVal
      Case 2: TempGroup.Pos.Width = newVal
      Case 3: TempGroup.Pos.Height = newVal
    End Select
    Call AssignTextAndTag(txtctl, newVal)
    'THROW DIRTY FLAG AND REFRESH WINDOW.
    'Call frmBed_DirtyFlag_Throw
    Call Refresh_frmEditGroup
  End If
  Call Global_LostFocus(txtctl)
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
      Case 0: TempGroup.Name = Trim$(txtctl.Text)
    End Select
    Call AssignTextAndTag(txtctl, txtctl.Text)
    'THROW DIRTY FLAG AND REFRESH WINDOW.
    'Call frmBed_DirtyFlag_Throw
    Call Refresh_frmEditGroup
  End If
  Call Global_LostFocus(txtctl)
End Sub


