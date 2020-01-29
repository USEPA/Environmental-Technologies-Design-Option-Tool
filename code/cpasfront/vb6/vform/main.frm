VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Clean Process Advisory System"
   ClientHeight    =   6330
   ClientLeft      =   4065
   ClientTop       =   1950
   ClientWidth     =   6450
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   6450
   Begin MSComctlLib.StatusBar StatusBar_Main 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   6045
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.PictureBox picGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   0
      Left            =   270
      ScaleHeight     =   2655
      ScaleWidth      =   4785
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   4785
      Begin MSComctlLib.ListView lvGroup 
         Height          =   2295
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   300
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   4048
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   8421376
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblGroup 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "lblGroup(0)"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   30
         Visible         =   0   'False
         Width           =   4635
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6450
      TabIndex        =   3
      Top             =   5220
      Width           =   6450
      Begin TabDlg.SSTab sstab_Main 
         Height          =   405
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   30000
         _ExtentX        =   52917
         _ExtentY        =   714
         _Version        =   327681
         TabOrientation  =   1
         Style           =   1
         TabHeight       =   520
         BackColor       =   8421504
         ForeColor       =   4210816
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "main.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "main.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "main.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   6450
      TabIndex        =   2
      Top             =   5715
      Width           =   6450
      Begin Threed.SSPanel sspanel_StatusInfo 
         Height          =   315
         Left            =   2130
         TabIndex        =   5
         Top             =   0
         Width           =   7905
         _Version        =   65536
         _ExtentX        =   13944
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "sspanel_StatusInfo"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel sspanel_Dirty 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "sspanel_Dirty"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1815
      Left            =   540
      TabIndex        =   0
      Top             =   3090
      Visible         =   0   'False
      Width           =   3105
      _Version        =   65536
      _ExtentX        =   5477
      _ExtentY        =   3201
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
      Begin MSComctlLib.ImageList il_Icons 
         Left            =   240
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.ListBox lstSorter 
         Height          =   450
         Left            =   1380
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1245
      End
      Begin VB.PictureBox picFontTest 
         Height          =   525
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   1185
         TabIndex        =   1
         Top             =   300
         Width           =   1245
      End
      Begin MSComctlLib.ImageList il_SmallIcons 
         Left            =   840
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Erase All ..."
         Index           =   10
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Revert to Saved Version ..."
         Index           =   20
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   30
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   199
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   200
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Copy Icon"
         Index           =   10
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Cu&t Icon"
         Index           =   20
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Paste Icon"
         Index           =   30
      End
   End
   Begin VB.Menu mnuTabs 
      Caption         =   "T&abs"
      Begin VB.Menu mnuTabsItem 
         Caption         =   "&Add Tab ..."
         Index           =   10
      End
      Begin VB.Menu mnuTabsItem 
         Caption         =   "&Delete Tab ..."
         Index           =   20
      End
      Begin VB.Menu mnuTabsItem 
         Caption         =   "&Edit Tab ..."
         Index           =   30
      End
      Begin VB.Menu mnuTabsItem 
         Caption         =   "&Organize Tabs ..."
         Index           =   40
      End
   End
   Begin VB.Menu mnuGroups 
      Caption         =   "&Groups"
      Begin VB.Menu mnuGroupsItem 
         Caption         =   "&Add Group ..."
         Index           =   10
      End
      Begin VB.Menu mnuGroupsItem 
         Caption         =   "&Delete Group ..."
         Index           =   20
      End
      Begin VB.Menu mnuGroupsItem 
         Caption         =   "&Edit Group ..."
         Index           =   30
      End
      Begin VB.Menu mnuGroupsItem 
         Caption         =   "&Move Group ..."
         Index           =   40
      End
      Begin VB.Menu mnuGroupsItem 
         Caption         =   "&Resize Group ..."
         Index           =   50
      End
   End
   Begin VB.Menu mnuIcons 
      Caption         =   "&Icons"
      Begin VB.Menu mnuIconsItem 
         Caption         =   "&Add Icon ..."
         Index           =   10
      End
      Begin VB.Menu mnuIconsItem 
         Caption         =   "&Delete Icon ..."
         Index           =   20
      End
      Begin VB.Menu mnuIconsItem 
         Caption         =   "&Edit Icon ..."
         Index           =   30
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsItem 
         Caption         =   "&Mode"
         Index           =   10
         Begin VB.Menu mnuToolsModeItem 
            Caption         =   "&Design Mode"
            Index           =   10
         End
         Begin VB.Menu mnuToolsModeItem 
            Caption         =   "&User Mode"
            Index           =   20
         End
         Begin VB.Menu mnuToolsModeItem 
            Caption         =   "-"
            Index           =   99
         End
         Begin VB.Menu mnuToolsModeItem 
            Caption         =   "&Toggle Between Modes"
            Index           =   100
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu mnuToolsItem 
         Caption         =   "&Options ..."
         Index           =   100
      End
      Begin VB.Menu mnuToolsItem 
         Caption         =   "&Refresh"
         Index           =   120
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&View Version History ..."
         Index           =   10
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   90
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About ..."
         Index           =   99
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Test &1"
         Index           =   200
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmMain_HIGH_GROUP_CODE As Integer
Dim frmMain_SELECTED_GROUP As Integer
Dim frmMain_SELECTED_ICON As Integer
Dim frmMain_TELL_lvGroup_Click_NOT_TO_CLEAR_SELECTED_ICON As Boolean

Dim frmMain_GROUP_MOVE_IN_PROGRESS As Boolean
Dim frmMain_GROUP_MOVE_IDX As Integer
Dim frmMain_GROUP_OFFSET_X As Double
Dim frmMain_GROUP_OFFSET_Y As Double

Const MARGIN_LV = 100
Const MARGIN_LV_FROM_LABEL = 200

Public MAIN_POPUPWINDOW_GROUP_SHOWN As Integer
Public MAIN_POPUPWINDOW_ICON_SHOWN As Integer

Dim HALT_Form_Click As Boolean





Const frmMain_declarations_end = 0


Function Get_IconStructure_Code_From_Key(ThisKey As String) As Integer
Dim IconStructure_CodeStr As String
Dim IconStructure_Code As Integer
Dim i As Integer
Dim ThisChar As String
  IconStructure_CodeStr = ""
  i = Len(ThisKey)
  Do While (1 = 1)
    ThisChar = Mid$(ThisKey, i, 1)
    If (Asc(ThisChar) >= Asc("0") And (Asc(ThisChar) <= Asc("9"))) Then
      IconStructure_CodeStr = ThisChar & IconStructure_CodeStr
    Else
      Exit Do
    End If
    i = i - 1
    If (i < 1) Then Exit Do
  Loop
  IconStructure_Code = Val(IconStructure_CodeStr)
  Get_IconStructure_Code_From_Key = IconStructure_Code
End Function


Private Sub Form_Click()
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  If (HALT_Form_Click) Then Exit Sub
  'DESELECT ANY SELECTED GROUPS.
  Call Main_Group_Unselect_All
End Sub
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then
    'HANDLE COMPLETION OF MOVE HERE.
    Call Main_Group_Complete_Move(CDbl(X), CDbl(Y))
    'DESELECT ANY SELECTED GROUPS.
    Call Main_Group_Unselect_All
    Call StatusInfo_Display(sspanel_StatusInfo, "", 12, 15)
    Call Project_DirtyFlag_Throw(NowProj)
  End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then
    If (KeyAscii = 27) Then
      'DESELECT ANY SELECTED GROUPS.
      Call Main_Group_Unselect_All
      'CANCEL GROUP MOVE.
      frmMain_GROUP_MOVE_IN_PROGRESS = False
      Call Main_Refresh_CurrentTab
      Call StatusInfo_Display(sspanel_StatusInfo, "", 12, 15)
    End If
  End If
End Sub
Private Sub Form_Load()
'Call DebugOutputFile("a")
  If (ALLOW_DESIGN_MODE) Then
    'DO NOTHING.
  Else
    'DISABLE THE TOGGLE MENU.
    mnuToolsItem(10).Visible = False
  End If
  'MISC INITS.
  frmMain_MODE = frmMain_MODE_USER    'temp: just in case it's needed in the next few lines of code!
  frmMain_HIGH_GROUP_CODE = 0
  frmMain_SELECTED_GROUP = 0
  frmMain_GROUP_MOVE_IN_PROGRESS = False
  frmMain_GROUP_MOVE_IDX = 0
  Call StatusInfo_Display(sspanel_StatusInfo, "", 12, 15)
'Call DebugOutputFile("b")
  'MISC REFRESHS.
  If (FileExists(fn_Full_MainDataFile)) Then
    If (MainDatafile_Input(NowProj) = False) Then
      Call Show_Error("Unable to load workspace.  Recommendation: " & _
          "Perform a complete re-install of the software, or else " & _
          "delete the file `" & fn_Full_MainDataFile & "` and " & _
          "re-run this program.")
      End
    End If
  Else
    Call Show_Message("About to create a new workspace file ...")
    Call Project_New(NowProj)
    If (MainDatafile_Output(NowProj) = False) Then
      Call Show_Error("Unable to save workspace.  Recommendation: " & _
          "Perform a complete re-install of the software, or else " & _
          "delete the file `" & fn_Full_MainDataFile & "` and " & _
          "re-run this program.")
      End
    End If
  End If
'Call DebugOutputFile("c")
  'POSITION THIS WINDOW.
  Me.Width = NowProj.Pos.Width
  Me.Height = NowProj.Pos.Height
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = 100
  'Call CenterOnScreen(Me)
'Call DebugOutputFile("d")
  
  'REFRESH POSITION AND SIZE OF MAIN WINDOW.
  'frmMain.Move _
      NowProj.Pos.Left, _
      NowProj.Pos.Top, _
      NowProj.Pos.Width, _
      NowProj.Pos.Height
  'COMMENTED OUT REFRESH POSITION CODE
  'TO AVOID DEALING WITH IT (FOR NOW).
'Call DebugOutputFile("e")
  
  Call Project_ReadIniVariables(NowProj)
'Call DebugOutputFile("f")
  frmMain_MODE = NowProj.StartupMode
  Call Project_DirtyFlag_RefreshScreen(NowProj)
'Call DebugOutputFile("w")
  Call Main_Refresh_Tabs
'Call DebugOutputFile("y")
  sstab_Main.Tab = 0
  Call Main_Refresh_CurrentTab
'Call DebugOutputFile("z")
End Sub


Sub Main_Refresh_Tabs()
Dim i As Integer
  
  'REFRESH NUMBER OF TABS, AND TAB CAPTIONS.
  sstab_Main.TabsPerRow = 10
  sstab_Main.Tabs = NowProj.Tabs_Count
  For i = 1 To NowProj.Tabs_Count
    sstab_Main.TabCaption(i - 1) = NowProj.Tabs(i).Name
  Next i
End Sub

Private Sub SSTab1_DblClick()

End Sub


Sub Main_EnableDisable_Stuff()
Dim IsEnabled_Menus As Boolean
Dim IsEnabled_TabBar As Boolean
Dim IsVisible_DesignMenus As Boolean
  IsEnabled_Menus = True
  IsEnabled_TabBar = True
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then
    IsEnabled_Menus = False
    IsEnabled_TabBar = False
  End If
  Select Case frmMain_MODE
    Case frmMain_MODE_DESIGN:
      IsVisible_DesignMenus = True
      mnuToolsModeItem(10).Checked = True
      mnuToolsModeItem(20).Checked = False
    Case frmMain_MODE_USER:
      IsVisible_DesignMenus = False
      mnuToolsModeItem(10).Checked = False
      mnuToolsModeItem(20).Checked = True
  End Select
  mnuFile.Enabled = IsEnabled_Menus
  mnuEdit.Enabled = IsEnabled_Menus
  mnuTabs.Enabled = IsEnabled_Menus
  mnuGroups.Enabled = IsEnabled_Menus
  mnuIcons.Enabled = IsEnabled_Menus
  mnuTools.Enabled = IsEnabled_Menus
  mnuHelp.Enabled = IsEnabled_Menus
  sstab_Main.Enabled = IsEnabled_TabBar
  'MAKE ITEMS VISIBLE/INVISIBLE DEPENDING ON USER/DESIGN MODE.
  mnuEdit.Visible = IsVisible_DesignMenus
  mnuTabs.Visible = IsVisible_DesignMenus
  mnuGroups.Visible = IsVisible_DesignMenus
  mnuIcons.Visible = IsVisible_DesignMenus
  mnuFileItem(10).Visible = IsVisible_DesignMenus
  mnuFileItem(20).Visible = IsVisible_DesignMenus
  mnuFileItem(30).Visible = IsVisible_DesignMenus
  mnuFileItem(199).Visible = IsVisible_DesignMenus
End Sub


Sub Main_Unload_All_Groups()
Dim i As Integer
  On Error Resume Next
  'UN-DISPLAY CURRENT GROUPS.
  For i = 1 To frmMain_HIGH_GROUP_CODE
    Unload lblGroup(i)
    Unload lvGroup(i)
    Unload picGroup(i)
  Next i
  'CLEAR ALL ICON IMAGES.
  il_Icons.ListImages.Clear
  il_SmallIcons.ListImages.Clear
End Sub
Sub Main_Refresh_CurrentTab()
Dim NowTab As Integer
Dim Now_fn_Pic As String
Dim i As Integer
Dim j As Integer

'File_IsExists

  'SET CURRENT TAB CODE.
  NowTab = sstab_Main.Tab
  
  'UNLOAD ALL OF THE GROUP OBJECTS ON THIS TAB.
  Call Main_Unload_All_Groups
  
  'DISPLAY NEW BACKGROUND COLOR.
  Me.BackColor = NowProj.Tabs(NowTab + 1).TabBackgroundColor.Color
  
  'DISPLAY CURRENT GROUPS.
Dim gr As GroupType
Dim ic As IconType
Dim font_height As Double
Dim clmX As ColumnHeader
Dim imgX As ListImage
Dim itmX As ListItem
Dim fn_This_Image As String
Dim This_Icon_Key As String
Dim WhichIcon As Integer
Dim nPos As PositionInfoType
  frmMain_HIGH_GROUP_CODE = NowProj.Tabs(NowTab + 1).Groups_Count
  For i = 1 To frmMain_HIGH_GROUP_CODE
    gr = NowProj.Tabs(NowTab + 1).Groups(i)
    Load picGroup(i)
    Load lblGroup(i)
    Load lvGroup(i)
    'SET UP PICTURE (HOLDS LABEL AND LISTVIEW CONTROLS).
    picGroup(i).Move _
        gr.Pos.Left, _
        gr.Pos.Top, _
        gr.Pos.Width, _
        gr.Pos.Height
    picGroup(i).BackColor = gr.GroupBackgroundColor.Color
    picGroup(i).ForeColor = gr.GroupForegroundColor.Color
    picGroup(i).Visible = True
    picGroup(i).ZOrder
    'SET UP LABEL CONTROL.
    Set lblGroup(i).Container = picGroup(i)
    font_height = FontTest_GetHeight( _
        picFontTest, _
        gr.GroupTitleFont)
    Call FontInfo_SetFontOnControl(lblGroup(i), gr.GroupTitleFont)
    lblGroup(i).Move _
        0, _
        0, _
        picGroup(i).Width, _
        font_height
    lblGroup(i).Caption = gr.Name
    lblGroup(i).BackColor = gr.GroupBackgroundColor.Color
    lblGroup(i).ForeColor = gr.GroupForegroundColor.Color
    lblGroup(i).Visible = True
    lblGroup(i).ZOrder
    'SET UP LISTVIEW CONTROL.
    Set lvGroup(i).Container = picGroup(i)
    'Call FontInfo_SetFontOnControl(lvGroup(i), gr.GroupIconFont)
    nPos.Left = MARGIN_LV
    nPos.Top = lblGroup(i).Height + MARGIN_LV_FROM_LABEL
    nPos.Width = picGroup(i).Width - 2 * MARGIN_LV
    nPos.Height = picGroup(i).Height - lblGroup(i).Height - 2 * MARGIN_LV - MARGIN_LV_FROM_LABEL
    If (nPos.Width < 1000) Then nPos.Width = 1000
    If (nPos.Height < 1000) Then nPos.Height = 1000
    lvGroup(i).Move _
        nPos.Left, _
        nPos.Top, _
        nPos.Width, _
        nPos.Height
    lvGroup(i).BackColor = gr.GroupBackgroundColor.Color
    lvGroup(i).ForeColor = gr.GroupForegroundColor.Color
    lvGroup(i).Visible = True
    lvGroup(i).ZOrder
    'SORT ICONS WITHIN CURRENT LISTVIEW CONTROL; ALSO, DETERMINE
    'WHICH APPLICATIONS ARE MISSING.
Dim ThisIconName As String
Dim ThisIconTarget0 As String
Dim ThisIconTarget As String
Dim NumIconsToDisplay As Integer
Dim FileDoesExist As Boolean
Dim RefuseToDisplay As Boolean
    lstSorter.Clear
    NumIconsToDisplay = 0
    For j = 1 To gr.Icons_Count
      ThisIconName = gr.Icons(j).Name
      ThisIconTarget0 = gr.Icons(j).fn_ApplicationLink
      ThisIconTarget = String_PrepareForApplicationLaunch(ThisIconTarget0)
      FileDoesExist = File_IsExists(ThisIconTarget)
      If (Not FileDoesExist) Then
        ThisIconName = ThisIconName & " (*)"
      End If
      RefuseToDisplay = False
      If (Not NowProj.DisplayUninstalledApplications) And (Not FileDoesExist) Then
        If (frmMain_MODE <> frmMain_MODE_DESIGN) Then
          'IF FILE DOES NOT EXIST, AND USER DOES NOT WANT TO SEE THE
          'NON-EXISTING FILES, AND THE CURRENT MODE IS "USER",
          'THEN DO NOT DISPLAY THIS ICON.
          RefuseToDisplay = True
        End If
      End If
      If (RefuseToDisplay = False) Then
        lstSorter.AddItem ThisIconName
        lstSorter.ItemData(lstSorter.NewIndex) = j
        NumIconsToDisplay = NumIconsToDisplay + 1
      End If
    Next j
    'REFRESH ICONS ON LISTVIEW CONTROL.
    Set clmX = lvGroup(i).ColumnHeaders.Add( _
      , , "Name", lvGroup(i).Width / 2)
    lvGroup(i).View = lvwReport
    'SORTING RELATED MODIFICATION.
    'For j = 1 To gr.Icons_Count
    For j = 1 To NumIconsToDisplay
      'SORTING RELATED MODIFICATION.
      'ic = gr.Icons(j)
      ic = gr.Icons(lstSorter.ItemData(j - 1))
      ThisIconName = lstSorter.List(j - 1)
      fn_This_Image = fpath_Icons & "\" & ic.fn_IconImage
      'SORTING RELATED MODIFICATION.
      'This_Icon_Key = Trim$(Str$(i)) & "-" & Trim$(Str$(j))
      This_Icon_Key = Trim$(Str$(i)) & "-" & Trim$(Str$(lstSorter.ItemData(j - 1)))
      On Error Resume Next
      Set imgX = il_Icons.ListImages.Add( _
        , This_Icon_Key, LoadPicture(fn_This_Image))
      Set imgX = il_SmallIcons.ListImages.Add( _
        , This_Icon_Key, LoadPicture(fn_This_Image))
      On Error GoTo 0
      WhichIcon = il_SmallIcons.ListImages.Count
      lvGroup(i).Icons = il_Icons
      lvGroup(i).SmallIcons = il_SmallIcons
      'Set itmX = lvGroup(i).ListItems.Add( _
      '    , This_Icon_Key, ic.Name)   ', 1)  '!!!!!!!!!!!!
      Set itmX = lvGroup(i).ListItems.Add( _
          , This_Icon_Key, ThisIconName)   ', 1)  '!!!!!!!!!!!!
      itmX.Icon = WhichIcon
      itmX.SmallIcon = WhichIcon
    Next j
    'SET THE VIEW AND ARRANGE PROPERTIES ON THE LISTVIEW CONTROL.
    lvGroup(i).View = NowProj.lvGroups_View
    lvGroup(i).Arrange = NowProj.lvGroups_Arrange
  
'    Dim clmX As ColumnHeader
'  Set clmX = lvGroup(0).ColumnHeaders.Add( _
'      , , "author", lvGroup(0).Width / 3)
'  'Set clmX = lvGroup(0).ColumnHeaders.Add( _
'  '    , , "author id", lvGroup(0).Width / 3, lvwColumnCenter)
'  'Set clmX = lvGroup(0).ColumnHeaders.Add( _
'  '    , , "bday", lvGroup(0).Width / 3)
'
'  lvGroup(0).View = lvwReport
'
'  Dim imgX As ListImage
'  'Dim imgX2 As ListImage
'  'Set imgX = ImageList1.ListImages.Add( _
'      , , LoadPicture(App.Path & "\note06.ico"))
'  Set imgX = ImageList1.ListImages.Add( _
'      , , LoadPicture(App.Path & "\note06.ico"))
'  Set imgX = ImageList1.ListImages.Add( _
'      , , LoadPicture(App.Path & "\note07.ico"))
'  Set imgX = ImageList2.ListImages.Add( _
'      , , LoadPicture(App.Path & "\w.bmp"))
'  Set imgX = ImageList2.ListImages.Add( _
'      , , LoadPicture(App.Path & "\w.bmp"))
'  'Set imgX = ImageList2.ListImages.Add( _
'  '    , , LoadPicture(App.Path & "\note06.ico"))
'  lvGroup(0).Icons = ImageList1
'  lvGroup(0).SmallIcons = ImageList2
'
'  With Combo1
'    .AddItem "Icon"         '0
'    .AddItem "SmallIcon"    '1
'    .AddItem "List"         '2
'    .AddItem "Report"       '3
'  End With
'
'  Dim myDb As Database, myRs As Recordset
'  Set myDb = DBEngine.Workspaces(0).OpenDatabase("d:\develop\vb5\biblio.mdb")
'  Set myRs = myDb.OpenRecordset("authors", dbOpenDynaset)
'  Dim itmX As ListItem
'  Dim nowCount As Integer
'  nowCount = 0
'Dim WhichIcon As Integer
'WhichIcon = 1
'  While Not myRs.EOF
'    Set itmX = lvGroup(0).ListItems.Add( _
'        , , CStr(myRs!author), 1)
'    itmX.Icon = WhichIcon
'    itmX.SmallIcon = WhichIcon
'    WhichIcon = WhichIcon + 1
'    If (WhichIcon > 2) Then WhichIcon = 1
'
'    'If Not IsNull(myRs!au_id) Then
'    '  itmX.SubItems(1) = CStr(myRs!au_id)
'    'End If
'    'If Not IsNull(myRs![year born]) Then
'    '  itmX.SubItems(2) = myRs![year born]
'    'End If
'    myRs.MoveNext
'    nowCount = nowCount + 1
'    If (nowCount > 100) Then GoTo blah
'  Wend

  
  
  
  Next i

  'DISPLAY CURRENT PICTURE.
  Now_fn_Pic = Trim$(NowProj.Tabs(NowTab + 1).fn_BackgroundImage)
  If (Now_fn_Pic = "") Then
    Set frmMain.Picture = LoadPicture()
  Else
    Now_fn_Pic = fpath_Backgrounds & "\" & Now_fn_Pic
    Set frmMain.Picture = LoadPicture(Now_fn_Pic)
  End If

  'PERFORM ENABLE/DISABLES.
  Call Main_EnableDisable_Stuff

End Sub


Sub Main_Group_Complete_Move(X As Double, Y As Double)
Dim NowTab As Integer
Dim gr As GroupType
  NowTab = sstab_Main.Tab
  gr = NowProj.Tabs(NowTab + 1).Groups(frmMain_GROUP_MOVE_IDX)
  gr.Pos.Left = X - frmMain_GROUP_OFFSET_X
  gr.Pos.Top = Y - frmMain_GROUP_OFFSET_Y
  NowProj.Tabs(NowTab + 1).Groups(frmMain_GROUP_MOVE_IDX) = gr
  frmMain_GROUP_MOVE_IN_PROGRESS = False
  frmMain_GROUP_MOVE_IDX = 0
  Call Main_Refresh_CurrentTab
  'PERFORM ENABLE/DISABLES.
  Call Main_EnableDisable_Stuff
End Sub
Sub Main_Group_Select_For_Move(idx As Integer)
Dim NowTab As Integer
Dim gr As GroupType
Dim i As Integer
  frmMain_GROUP_MOVE_IDX = idx
  NowTab = sstab_Main.Tab
  For i = 1 To NowProj.Tabs(NowTab + 1).Groups_Count
    gr = NowProj.Tabs(NowTab + 1).Groups(i)
    If (i <> frmMain_GROUP_MOVE_IDX) Then
      picGroup(i).Visible = False
    Else
      picGroup(i).Visible = True
      lblGroup(i).Visible = False
      lvGroup(i).Visible = False
    End If
  Next i
  'PERFORM ENABLE/DISABLES.
  Call Main_EnableDisable_Stuff
End Sub
Sub Main_Group_Select(idx As Integer)
Dim NowTab As Integer
Dim gr As GroupType
Dim i As Integer
  'FIRST, UNSELECT ALL OTHER GROUPS (NO MULTIPLE SELECTIONS ALLOWED!).
  ''''Call Main_Group_Unselect_All
  NowTab = sstab_Main.Tab
  For i = 1 To NowProj.Tabs(NowTab + 1).Groups_Count
    If (i <> idx) Then
      Call Main_Group_Unselect_Visually(i)
    End If
  Next i
  'SELECT THIS GROUP.
  frmMain_SELECTED_GROUP = idx
  gr = NowProj.Tabs(NowTab + 1).Groups(idx)
  If (frmMain_MODE = frmMain_MODE_DESIGN) Then
    'ONLY HIGHLIGHT GROUP WHEN IN DESIGN MODE.
    lblGroup(idx).BackColor = &HFFFFFF - gr.GroupBackgroundColor.Color
    lblGroup(idx).ForeColor = &HFFFFFF - gr.GroupForegroundColor.Color
    picGroup(idx).BackColor = &HFFFFFF - gr.GroupBackgroundColor.Color
    picGroup(idx).ForeColor = &HFFFFFF - gr.GroupForegroundColor.Color
  End If
End Sub
Sub Main_Group_Unselect_Visually(idx As Integer)
Dim NowTab As Integer
Dim gr As GroupType
  NowTab = sstab_Main.Tab
  gr = NowProj.Tabs(NowTab + 1).Groups(idx)
  lblGroup(idx).BackColor = gr.GroupBackgroundColor.Color
  lblGroup(idx).ForeColor = gr.GroupForegroundColor.Color
  picGroup(idx).BackColor = gr.GroupBackgroundColor.Color
  picGroup(idx).ForeColor = gr.GroupForegroundColor.Color
End Sub
Sub Main_Group_Unselect_All()
Dim NowTab As Integer
Dim gr As GroupType
Dim i As Integer
'Debug.Print "Main_Group_Unselect_All"
  frmMain_SELECTED_GROUP = 0
  NowTab = sstab_Main.Tab
  For i = 1 To NowProj.Tabs(NowTab + 1).Groups_Count
    Call Main_Group_Unselect_Visually(i)
    'gr = NowProj.Tabs(NowTab + 1).Groups(i)
    'lblGroup(i).BackColor = gr.GroupBackgroundColor.Color
    'lblGroup(i).ForeColor = gr.GroupForegroundColor.Color
    'picGroup(i).BackColor = gr.GroupBackgroundColor.Color
    'picGroup(i).ForeColor = gr.GroupForegroundColor.Color
    ''Set lvGroup(i).SelectedItem = Nothing
    Set lvGroup(i).SelectedItem = Nothing
  Next i
  Call StatusInfo_Display(sspanel_StatusInfo, "", 14, 0)
End Sub
Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Dim li As ListItem
Dim Dummy As String
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  If (frmMain_MODE <> frmMain_MODE_DESIGN) Then Exit Sub
  If (button = 2) Then
'Debug.Print "lvGroup_MouseDown"
    'Call Main_Group_Select(Index)
    'If (frmMain_SELECTED_GROUP <= 0) Then
    '  Exit Sub
    'End If
    HALT_Form_Click = True
    PopupMenu frmMain.mnuTabs
    HALT_Form_Click = False
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (file_query_unload(NowProj) = False) Then
    Cancel = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  NowProj.StartupMode = frmMain_MODE
  Call Project_WriteIniVariables(NowProj)
  Call Close_All_Windows
End Sub


Private Sub lblGroup_Click(Index As Integer)
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  Call Main_Group_Select(Index)
End Sub
Private Sub lvGroup_Click(Index As Integer)
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  Call Main_Group_Select(Index)
  If (Not frmMain_TELL_lvGroup_Click_NOT_TO_CLEAR_SELECTED_ICON) Then
    frmMain_SELECTED_ICON = 0
  End If
  frmMain_TELL_lvGroup_Click_NOT_TO_CLEAR_SELECTED_ICON = False
  'Debug.Print "click"
End Sub
Private Sub lvGroup_DblClick(Index As Integer)
Dim NowTab As Integer
  If (frmMain_SELECTED_GROUP <= 0) Or (frmMain_SELECTED_ICON <= 0) Then
    'Call Show_Error("You must first select a group and an icon.")
    'NO ICON SELECTED; EXIT OUT OF HERE.
    Exit Sub
  End If
  NowTab = sstab_Main.Tab
  Call Launch_Application( _
      NowProj, _
      NowTab + 1, _
      frmMain_SELECTED_GROUP, _
      frmMain_SELECTED_ICON)
  'MsgBox "Testing! ", vbInformation
  'Call lvGroup_ItemClick(Index, li)
  'HALT_Form_Click = True
  'PopupMenu frmMain.mnuIcons    ', vbPopupMenuCenterAlign
  'HALT_Form_Click = False
End Sub
''''Private Sub lvGroup_ItemClick(Index As Integer, ByVal Item As ComctlLib.ListItem)
Private Sub lvGroup_ItemClick(Index As Integer, ByVal Item As ListItem)
Dim NowTab As Integer
Dim gr As GroupType
Dim msg As String
'Dim IconStructure_CodeStr As String
Dim IconStructure_Code As Integer
'Dim i As Integer
'Dim ThisChar As String
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  Call Main_Group_Select(Index)
  If (frmMain_SELECTED_GROUP <= 0) Then
    Exit Sub
  End If
  'SORTING RELATED MODIFICATION.
  'frmMain_SELECTED_ICON = Item.Index
  IconStructure_Code = Get_IconStructure_Code_From_Key(Item.Key)
  frmMain_SELECTED_ICON = IconStructure_Code
  frmMain_TELL_lvGroup_Click_NOT_TO_CLEAR_SELECTED_ICON = True
  'Debug.Print "lvGroup_ItemClick"
  'NowTab = sstab_Main.Tab
  'gr = NowProj.Tabs(NowTab + 1).Groups(frmMain_SELECTED_GROUP)
  'msg = gr.Icons(Item.Index).LongName '& " (" & Trim$(Str$(Item.Index)) & ")"
  'Call StatusInfo_Display(sspanel_StatusInfo, msg, 14, 0)
End Sub
Private Sub lvGroup_KeyPress(Index As Integer, KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call lvGroup_DblClick(Index)
  End If
End Sub
Private Sub lvGroup_MouseDown(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Dim li As ListItem
Dim Dummy As String
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  If (button = 2) Then
'Debug.Print "lvGroup_MouseDown"
    Call Main_Group_Select(Index)
    If (frmMain_SELECTED_GROUP <= 0) Then
      Exit Sub
    End If
    If (frmMain_MODE <> frmMain_MODE_DESIGN) Then Exit Sub
    'IS THE MOUSE POSITIONED ON AN ICON?
    Set li = lvGroup(Index).HitTest(X, Y)
    On Error Resume Next
    Dummy = li.Text
    If (Err <> 0) Then
      'NO -- DO THE "GROUP" POPUP MENU.
      HALT_Form_Click = True
      PopupMenu frmMain.mnuGroups    ', vbPopupMenuCenterAlign
      HALT_Form_Click = False
    Else
      'YES -- DO THE "ICON" POPUP MENU.
      Call lvGroup_ItemClick(Index, li)
      HALT_Form_Click = True
      PopupMenu frmMain.mnuIcons    ', vbPopupMenuCenterAlign
      HALT_Form_Click = False
    End If
  End If
End Sub
Private Sub lvGroup_MouseMove(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NowTab As Integer
Dim gr As GroupType
Dim msg_short As String
Dim msg_long As String
Dim li As ListItem
Dim Screen_X As Double
Dim Screen_Y As Double
Dim Dummy As String
Dim ThisIconCode As Integer
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  If (frmMain_MODE = frmMain_MODE_DESIGN) Then
    'BECAUSE THIS IS DESIGN MODE, DO _NOT_ AUTOMATICALLY
    'SELECT THIS GROUP, AND DO _NOT_ AUTOMATICALLY
    'DISPLAY THE POP-UP WINDOW.
    Exit Sub
  End If
  Select Case frmMain_MODE
    Case frmMain_MODE_USER:
      If (frmMain_SELECTED_GROUP <> Index) Then
        Call Main_Group_Select(Index)
      Else
        'GROUP HAS NOT CHANGED: DO NOT VISUALLY SELECT IT
        '(IF THE Main_Group_Select() SUBROUTINE IS CALLED,
        'THE WINDOW KIND OF FLICKERS WHICH IS ANNOYING).
      End If
    Case frmMain_MODE_DESIGN:
      Call Main_Group_Select(Index)
  End Select
  If (frmMain_SELECTED_GROUP <= 0) Then
    Exit Sub
  End If
  'IS THE MOUSE POINTER ON AN ICON?
  Set li = lvGroup(Index).HitTest(X, Y)
  On Error Resume Next
  Dummy = li.Text
  If (Err <> 0) Then
    'NO -- HIDE THE POPUP WINDOW, AND EXIT OUT OF HERE.
    'HIDE THE POPUP WINDOW (IF SHOWN).
    Call Main_PopupWindow_HideIt
    'ALSO: NOTE THAT NO ICON IS SELECTED.
    frmMain_SELECTED_ICON = 0
    Exit Sub
  End If
  'IS THIS MESSAGE ALREADY DISPLAYED?
  'SORTING RELATED MODIFICATION.
  ThisIconCode = Get_IconStructure_Code_From_Key(li.Key)
  'If (MAIN_POPUPWINDOW_ICON_SHOWN = li.Index) And _
  '    (MAIN_POPUPWINDOW_GROUP_SHOWN = frmMain_SELECTED_GROUP) Then
  If (MAIN_POPUPWINDOW_ICON_SHOWN = ThisIconCode) And _
      (MAIN_POPUPWINDOW_GROUP_SHOWN = frmMain_SELECTED_GROUP) Then
    'YES -- EXIT OUT OF HERE.
    Exit Sub
  End If
  'DISPLAY THE MESSAGES.
  NowTab = sstab_Main.Tab
  gr = NowProj.Tabs(NowTab + 1).Groups(frmMain_SELECTED_GROUP)
  'DISPLAY SHORT MESSAGE.
  'SORTING RELATED MODIFICATION.
  'msg_short = gr.Icons(li.Index).LongName      '& " (" & Trim$(Str$(Item.Index)) & ")"
  msg_short = gr.Icons(ThisIconCode).LongName      '& " (" & Trim$(Str$(Item.Index)) & ")"
  Call StatusInfo_Display(sspanel_StatusInfo, msg_short, 14, 0)
  'DISPLAY LONG MESSAGE.
  If (Not NowProj.ShowDescriptionText) Then
    'USER DOES _NOT_ WANT TO SEE THE POPUP WINDOWS.
    'EXIT OUTTA HERE.
    Exit Sub
  End If
  Screen_X = frmMain.Left + picGroup(Index).Left + _
      lvGroup(Index).Left + X
  Screen_Y = frmMain.Top + picGroup(Index).Top + _
      lvGroup(Index).Top + Y
  'SORTING RELATED MODIFICATION.
  'msg_long = gr.Icons(li.Index).DescriptionText
  msg_long = gr.Icons(ThisIconCode).DescriptionText
'msg_long = "Fate of Volatile Organics in Wastewater Treatment Plants (FaVOr) predicts the fate of volatile organic compounds in activated sludge wastewater treatment plants. test"
  Call frmPopupWindow.frmPopupWindow_Show( _
    msg_long, _
    Screen_X, Screen_Y)
  MAIN_POPUPWINDOW_GROUP_SHOWN = frmMain_SELECTED_GROUP
  'SORTING RELATED MODIFICATION.
  'MAIN_POPUPWINDOW_ICON_SHOWN = li.Index
  MAIN_POPUPWINDOW_ICON_SHOWN = ThisIconCode
End Sub


Private Sub mnuFileItem_Click(Index As Integer)
Dim msg As String
Dim RetVal As Integer
  Select Case Index
    Case 10:      'ERASE ALL.
      'Call Project_New
      msg = "This command will erase all tabs, groups, and " & _
          "icons that have been generated.  They will be replaced " & _
          "with a very simple set of icons.  Are you sure you want " & _
          "to proceed?"
      RetVal = MsgBox(msg, _
            vbYesNo + vbQuestion, "Delete Everything?")
      If (RetVal <> vbYes) Then Exit Sub
      Call Main_Unload_All_Groups
      Call Project_New(NowProj)
      Call Main_Refresh_Tabs
      sstab_Main.Tab = 0
      Call Main_Refresh_CurrentTab
    Case 20:      'REVERT TO SAVED VERSION.
    Case 30:      'SAVE.
      'OUTPUT THE WORKSPACE FILE.
      Call MainDatafile_Output(NowProj)
      'CLEAR THE DIRTY FLAG.
      Call Project_DirtyFlag_Clear(NowProj)
    Case 200:     'EXIT.
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub mnuGroupsItem_Click(Index As Integer)
Dim name_old As String
Dim name_new As String
Dim is_aborted As Integer
Dim NowTab As Integer
Dim i As Integer
Dim ta As TabType
Dim gr As GroupType
Dim msg As String
Dim RetVal As Integer
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  Select Case Index
    Case 10:      'ADD.
      'INPUT NAME OF NEW GROUP.
      NowTab = sstab_Main.Tab
      name_old = Group_GetNewNameDefault(NowProj, NowProj.Tabs(NowTab + 1))
      name_new = name_old
      Do While (1 = 1)
        name_new = frmNewName.frmNewName_GetName( _
            "Enter Name of New Group", _
            "Each Group on this Tab must have a unique name.", _
            name_new, _
            is_aborted)
        If (is_aborted) Then Exit Sub
        If (Not Group_IsNameExist(NowProj, NowProj.Tabs(NowTab + 1), name_new)) Then
          Exit Do
        End If
        Call Show_Error("That name already exists.  Choose another name.")
      Loop
      'SET DEFAULT PROPERTIES OF GROUP.
      ta = NowProj.Tabs(NowTab + 1)
      ta.Groups_Count = ta.Groups_Count + 1
      ReDim Preserve ta.Groups(1 To ta.Groups_Count)
      Call Group_SetDefaults(NowProj, gr)
      gr.Name = name_new
      ta.Groups(ta.Groups_Count) = gr
      NowProj.Tabs(NowTab + 1) = ta
      'THROW DIRTY FLAG, REFRESH WINDOW.
      Call Project_DirtyFlag_Throw(NowProj)
      Call Main_Refresh_CurrentTab
    Case 20:      'DELETE.
      If (frmMain_SELECTED_GROUP <= 0) Then
        Call Show_Error("You must first select a group.")
        Exit Sub
      End If
      NowTab = sstab_Main.Tab
      ta = NowProj.Tabs(NowTab + 1)
      If (ta.Groups_Count <= 1) Then
        Call Show_Error("You cannot delete the last group on a tab.")
        Exit Sub
      End If
      msg = "Deleting the group named `" & _
          NowProj.Tabs(NowTab + 1).Groups(frmMain_SELECTED_GROUP).Name & _
          "` from the tab named `" & NowProj.Tabs(NowTab + 1).Name & _
          "` is an irreversible process, and will result in the " & _
          "deletion of all icons contained by that group (if any).  " & _
          "Are you sure " & _
          "you want to proceed?"
      RetVal = MsgBox(msg, _
            vbYesNo + vbQuestion, "Delete Group?")
      If (RetVal <> vbYes) Then Exit Sub
      'DELETE THE GROUP.
      For i = frmMain_SELECTED_GROUP To ta.Groups_Count - 1
        ta.Groups(i) = ta.Groups(i + 1)
      Next i
      ta.Groups_Count = ta.Groups_Count - 1
      ReDim Preserve ta.Groups(1 To ta.Groups_Count)
      NowProj.Tabs(NowTab + 1) = ta
      'THROW DIRTY FLAG, REFRESH WINDOW.
      Call Project_DirtyFlag_Throw(NowProj)
      Call Main_Refresh_CurrentTab
    Case 30:      'EDIT.
      If (frmMain_SELECTED_GROUP <= 0) Then
        Call Show_Error("You must first select a group.")
        Exit Sub
      End If
      NowTab = sstab_Main.Tab
      Dim USER_HIT_CANCEL As Boolean
      Call frmEditGroup.frmEditGroup_DoEdit( _
          NowTab + 1, _
          frmMain_SELECTED_GROUP, _
          USER_HIT_CANCEL)
      If (Not USER_HIT_CANCEL) Then
        'THROW DIRTY FLAG, REFRESH WINDOW.
        Call Project_DirtyFlag_Throw(NowProj)
        Call Main_Refresh_CurrentTab
      End If
    Case 40:      'MOVE.
      If (frmMain_SELECTED_GROUP <= 0) Then
        Call Show_Error("You must first select a group.")
        Exit Sub
      End If
      frmMain_GROUP_MOVE_IN_PROGRESS = True
      Call Main_Group_Select_For_Move(frmMain_SELECTED_GROUP)
      'picGroup(frmMain_SELECTED_GROUP).DragMode = 1
      Call StatusInfo_Display(sspanel_StatusInfo, _
          "  Drag selected group to " & _
          "move it, or hit Escape to cancel move.", 12, 15)
    Case 50:      'RESIZE.
      
      Call Show_Error("not implemented yet ! .... ")
      
  End Select
End Sub


Private Sub mnuHelpItem_Click(Index As Integer)
Dim RetVal As Integer
Dim CmdLine As String
  Select Case Index
    Case 10:      'VIEW VERSION HISTORY.
      Call Launch_Notepad(fpath_dbase & "\readme_cpasfront.txt")
    Case 99:     'About ...
      frmAbout.Show 1
    Case 200:     'TEST 1.
      ChDrive "x:"
      ChDir "\\cen-server\srcsafe\cpas10\code\cpasfront\vb5\asap\help"
      CmdLine = "winhelp asap.hlp"
      RetVal = 0 * Shell(CmdLine, 1)
  End Select
End Sub

Private Sub mnuIconsItem_Click(Index As Integer)
Dim name_old As String
Dim name_new As String
Dim is_aborted As Integer
Dim NowTab As Integer
Dim ta As TabType
Dim gr As GroupType
Dim ic As IconType
Dim msg As String
Dim RetVal As Integer
Dim i As Integer
  Select Case Index
    Case 10:      'ADD.
      If (frmMain_SELECTED_GROUP <= 0) Then
        Call Show_Error("You must first select a group.")
        Exit Sub
      End If
      'INPUT NAME OF NEW GROUP.
      NowTab = sstab_Main.Tab
      name_old = Icon_GetNewNameDefault(NowProj, _
          NowProj.Tabs(NowTab + 1).Groups(frmMain_SELECTED_GROUP))
      name_new = name_old
      Do While (1 = 1)
        name_new = frmNewName.frmNewName_GetName( _
            "Enter Name of New Icon", _
            "Each Icon in this Group must have a unique name.", _
            name_new, _
            is_aborted)
        If (is_aborted) Then Exit Sub
        If (Not Icon_IsNameExist(NowProj, _
            NowProj.Tabs(NowTab + 1).Groups(frmMain_SELECTED_GROUP), name_new)) Then
          Exit Do
        End If
        Call Show_Error("That name already exists.  Choose another name.")
      Loop
      'SET DEFAULT PROPERTIES OF ICON.
      ta = NowProj.Tabs(NowTab + 1)
      gr = ta.Groups(frmMain_SELECTED_GROUP)
      gr.Icons_Count = gr.Icons_Count + 1
      ReDim Preserve gr.Icons(1 To gr.Icons_Count)
      ic = gr.Icons(gr.Icons_Count)
      Call Icon_SetDefaults(NowProj, ic)
      ic.Name = name_new
      gr.Icons(gr.Icons_Count) = ic
      ta.Groups(frmMain_SELECTED_GROUP) = gr
      NowProj.Tabs(NowTab + 1) = ta
      'THROW DIRTY FLAG, REFRESH WINDOW.
      Call Project_DirtyFlag_Throw(NowProj)
      Call Main_Refresh_CurrentTab
    Case 20:      'DELETE.
      If (frmMain_SELECTED_GROUP <= 0) Or (frmMain_SELECTED_ICON <= 0) Then
        Call Show_Error("You must first select a group and an icon.")
        Exit Sub
      End If
      'MsgBox lvGroup(frmMain_SELECTED_GROUP).ListItems(frmMain_SELECTED_ICON).Text, vbInformation
      NowTab = sstab_Main.Tab
      msg = "Deleting the icon named `" & _
          NowProj.Tabs(NowTab + 1).Groups(frmMain_SELECTED_GROUP).Icons(frmMain_SELECTED_ICON).Name & _
          "` from the group named `" & _
          NowProj.Tabs(NowTab + 1).Groups(frmMain_SELECTED_GROUP).Name & _
          "` is an irreversible process.  Are you sure " & _
          "you want to proceed?"
      RetVal = MsgBox(msg, _
            vbYesNo + vbQuestion, "Delete Icon?")
      If (RetVal <> vbYes) Then Exit Sub
      'DELETE THE ICON.
      ta = NowProj.Tabs(NowTab + 1)
      gr = ta.Groups(frmMain_SELECTED_GROUP)
      For i = frmMain_SELECTED_ICON To gr.Icons_Count - 1
        gr.Icons(i) = gr.Icons(i + 1)
      Next i
      gr.Icons_Count = gr.Icons_Count - 1
      If (gr.Icons_Count >= 1) Then
        ReDim Preserve gr.Icons(1 To gr.Icons_Count)
      End If
      ta.Groups(frmMain_SELECTED_GROUP) = gr
      NowProj.Tabs(NowTab + 1) = ta
      'THROW DIRTY FLAG, REFRESH WINDOW.
      Call Project_DirtyFlag_Throw(NowProj)
      Call Main_Refresh_CurrentTab
    Case 30:      'EDIT.
      If (frmMain_SELECTED_GROUP <= 0) Or (frmMain_SELECTED_ICON <= 0) Then
        Call Show_Error("You must first select a group and an icon.")
        Exit Sub
      End If
      NowTab = sstab_Main.Tab
      Dim USER_HIT_CANCEL As Boolean
      Call frmEditIcon.frmEditIcon_DoEdit( _
          NowTab + 1, _
          frmMain_SELECTED_GROUP, _
          frmMain_SELECTED_ICON, _
          USER_HIT_CANCEL)
      If (Not USER_HIT_CANCEL) Then
        'THROW DIRTY FLAG, REFRESH WINDOW.
        Call Project_DirtyFlag_Throw(NowProj)
        Call Main_Refresh_CurrentTab
      End If
  End Select
End Sub


Private Sub mnuTabsItem_Click(Index As Integer)
Dim name_old As String
Dim name_new As String
Dim is_aborted As Integer
Dim NowTab As Integer
Dim ta As TabType
Dim msg As String
Dim RetVal As Integer
Dim i As Integer
Dim USER_HIT_CANCEL As Boolean
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  Select Case Index
    Case 10:      'ADD.
      'INPUT NAME OF NEW TAB.
      NowTab = sstab_Main.Tab
      name_old = Tab_GetNewNameDefault(NowProj)
      name_new = name_old
      Do While (1 = 1)
        name_new = frmNewName.frmNewName_GetName( _
            "Enter Name of New Tab", _
            "Each Tab in the program must have a unique name.", _
            name_new, _
            is_aborted)
        If (is_aborted) Then Exit Sub
        If (Not Tab_IsNameExist(NowProj, name_new)) Then
          Exit Do
        End If
        Call Show_Error("That name already exists.  Choose another name.")
      Loop
      'SET DEFAULT PROPERTIES OF TAB.
      NowProj.Tabs_Count = NowProj.Tabs_Count + 1
      ReDim Preserve NowProj.Tabs(1 To NowProj.Tabs_Count)
      Call Tab_SetDefaults(NowProj, ta)
      ta.Name = name_new
      NowProj.Tabs(NowProj.Tabs_Count) = ta
      'THROW DIRTY FLAG, REFRESH WINDOW.
      Call Project_DirtyFlag_Throw(NowProj)
      Call Main_Refresh_Tabs
      Call Main_Refresh_CurrentTab
    Case 20:      'DELETE.
      'If (frmMain_SELECTED_GROUPxxxx <= 0) Then
      '  Call Show_Error("You must first select a tab.")
      '  Exit Sub
      'End If
      NowTab = sstab_Main.Tab
      If (NowProj.Tabs_Count <= 1) Then
        Call Show_Error("You cannot delete the last tab.")
        Exit Sub
      End If
      If (Trim$(UCase$(NowProj.Tabs(NowTab + 1).Name)) = _
          Trim$(UCase$("CPAS Main Tools"))) Then
        Call Show_Error("You cannot delete the tab titled `CPAS Main Tools`.")
        Exit Sub
      End If
      msg = "Deleting the tab named `" & _
          NowProj.Tabs(NowTab + 1).Name & _
          "` is an irreversible process, and will result in " & _
          "the deletion of all groups and icons contained by " & _
          "that tab (if any).  Are you sure " & _
          "you want to proceed?"
      RetVal = MsgBox(msg, _
            vbYesNo + vbQuestion, "Delete Tab?")
      If (RetVal <> vbYes) Then Exit Sub
      'DELETE THE TAB.
      For i = NowTab + 1 To NowProj.Tabs_Count - 1
        NowProj.Tabs(i) = NowProj.Tabs(i + 1)
      Next i
      NowProj.Tabs_Count = NowProj.Tabs_Count - 1
      ReDim Preserve NowProj.Tabs(1 To NowProj.Tabs_Count)
      'THROW DIRTY FLAG, REFRESH WINDOW.
      Call Project_DirtyFlag_Throw(NowProj)
      Call Main_Refresh_Tabs
      Call Main_Refresh_CurrentTab
    Case 30:      'EDIT.
      NowTab = sstab_Main.Tab
      Call frmEditTab.frmEditTab_DoEdit( _
          NowTab + 1, _
          USER_HIT_CANCEL)
      If (Not USER_HIT_CANCEL) Then
        'THROW DIRTY FLAG, REFRESH WINDOW.
        Call Project_DirtyFlag_Throw(NowProj)
        Call Main_Refresh_Tabs
        Call Main_Refresh_CurrentTab
      End If
    Case 40:      'ORGANIZE.
      Call frmEditTabOrder.frmEditTabOrder_DoEdit( _
        USER_HIT_CANCEL)
      If (Not USER_HIT_CANCEL) Then
        'THROW DIRTY FLAG, REFRESH WINDOW.
        Call Project_DirtyFlag_Throw(NowProj)
        Call Main_Refresh_Tabs
        Call Main_Refresh_CurrentTab
      End If
  End Select
End Sub


Sub Do_Refresh_Command()
  Call Main_Refresh_Tabs
  Call Main_Refresh_CurrentTab
  'HIDE THE POPUP WINDOW (IF SHOWN).
  Call Main_PopupWindow_HideIt
End Sub


Private Sub mnuToolsItem_Click(Index As Integer)
Dim USER_HIT_CANCEL As Boolean
  Select Case Index
    Case 100:     'OPTIONS.
      Call frmToolsOptions.frmToolsOptions_DoEdit( _
        USER_HIT_CANCEL)
      If (Not USER_HIT_CANCEL) Then
        '''THROW DIRTY FLAG, REFRESH WINDOW.
        ''Call Project_DirtyFlag_Throw(NowProj)
        'REFRESH WINDOW.
        '(DO NOT THROW DIRTY FLAG: ALL OF THE TOOLS--OPTIONS VARIABLES
        'ARE STORED TO THE .INI FILE INSTEAD OF THE MAIN DATA FILE.)
        Call Main_Refresh_Tabs
        Call Main_Refresh_CurrentTab
      End If
    Case 120:     'REFRESH.
      Call Do_Refresh_Command
  End Select
End Sub


Private Sub mnuToolsModeItem_Click(Index As Integer)
Dim OLD_frmMain_MODE As Integer
  OLD_frmMain_MODE = frmMain_MODE
  Select Case Index
    Case 10:      'DESIGN MODE.
      frmMain_MODE = frmMain_MODE_DESIGN
    Case 20:      'USER MODE.
      frmMain_MODE = frmMain_MODE_USER
    Case 100:     'TOGGLE BETWEEN MODES.
      Select Case frmMain_MODE
        Case frmMain_MODE_DESIGN: frmMain_MODE = frmMain_MODE_USER
        Case frmMain_MODE_USER: frmMain_MODE = frmMain_MODE_DESIGN
      End Select
  End Select
  'UPDATE THE WINDOW.
  Select Case frmMain_MODE
    Case frmMain_MODE_DESIGN:
      Call Main_EnableDisable_Stuff
      'CLEARING THE SELECTED GROUP IS EASIER THAN UPDATING
      'THE DISPLAY (WHICH WOULD PROBABLY CONFUSE THE USER ANYWAY).
      frmMain_SELECTED_GROUP = 0
      'HIDE THE POPUP WINDOW (IF SHOWN).
      Call Main_PopupWindow_HideIt
    Case frmMain_MODE_USER:
      Call Main_EnableDisable_Stuff
  End Select
  If (OLD_frmMain_MODE <> frmMain_MODE) Then
    Call Do_Refresh_Command
  End If
End Sub


Private Sub picGroup_Click(Index As Integer)
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  Call Main_Group_Select(Index)
End Sub

Private Sub picGroup_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Dim form_X As Double
Dim form_Y As Double
  form_X = picGroup(Index).Left + X
  form_Y = picGroup(Index).Top + Y
  Call Form_DragDrop(Source, CSng(form_X), CSng(form_Y))
End Sub

Private Sub picGroup_MouseDown(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then
    frmMain_GROUP_OFFSET_X = X
    frmMain_GROUP_OFFSET_Y = Y
    picGroup(Index).Drag
    Exit Sub
  End If
End Sub
Private Sub picGroup_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then
    Exit Sub
  End If
End Sub


Public Sub Main_PopupWindow_HideIt()
  Call frmPopupWindow.frmPopupWindow_Hide
  Call StatusInfo_Display(sspanel_StatusInfo, "", 14, 0)
  frmMain.MAIN_POPUPWINDOW_GROUP_SHOWN = 0
  frmMain.MAIN_POPUPWINDOW_ICON_SHOWN = 0
End Sub


Private Sub sstab_Main_Click(PreviousTab As Integer)
  If (frmMain_GROUP_MOVE_IN_PROGRESS) Then Exit Sub
  'HIDE THE POPUP WINDOW (IF SHOWN).
  Call Main_PopupWindow_HideIt
  'DESELECT EVERYTHING.
  frmMain_SELECTED_GROUP = 0
  frmMain_SELECTED_ICON = 0
  frmMain_TELL_lvGroup_Click_NOT_TO_CLEAR_SELECTED_ICON = False
  'REFRESH WINDOW FOR THE NEW TAB.
  Call Main_Refresh_CurrentTab
End Sub

