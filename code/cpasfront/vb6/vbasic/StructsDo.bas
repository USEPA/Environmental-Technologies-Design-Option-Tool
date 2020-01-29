Attribute VB_Name = "StructsDo"
Option Explicit

Const FN_ICONIMAGE_DEFAULT = "default.ico"
  
  
  
  
  
Const StructsDo_declarations_end = 0


Sub FontInfo_GetFontFromControl( _
    picX As Control, _
    fi As FontInfoType)
  fi.FontBold = picX.FontBold
  fi.FontItalic = picX.FontItalic
  fi.FontName = picX.FontName
  fi.FontSize = picX.FontSize
  fi.FontStrikeThru = picX.FontStrikeThru
  fi.FontUnderline = picX.FontUnderline
End Sub
Sub FontInfo_SetFontOnControl( _
    picX As Control, _
    fi As FontInfoType)
  picX.FontBold = fi.FontBold
  picX.FontItalic = fi.FontItalic
  picX.FontName = fi.FontName
  picX.FontSize = fi.FontSize
  picX.FontStrikeThru = fi.FontStrikeThru
  picX.FontUnderline = fi.FontUnderline
End Sub
Function FontTest_GetHeight( _
    picX As Control, _
    fi As FontInfoType) _
    As Double
  Call FontInfo_SetFontOnControl(picX, fi)
  FontTest_GetHeight = CDbl(picX.TextHeight("W"))
End Function
Sub FontInfo_AssignDefault(fi As FontInfoType)
  fi.FontBold = False
  fi.FontItalic = False
  fi.FontName = "Arial"
  fi.FontSize = 12#
  fi.FontStrikeThru = False
  fi.FontUnderline = False
End Sub


Function Icon_GetNewNameDefault( _
    proj As ProjectType, _
    gr As GroupType) As String
Dim name_new As String
Dim i As Integer   'DETERMINE DEFAULT NAME OF NEW RECORD.
  i = 1
  Do While (1 = 1)
    name_new = "Icon " & Trim$(Str$(i))
    If (Not Icon_IsNameExist(proj, gr, name_new)) Then Exit Do
    i = i + 1
  Loop
  Icon_GetNewNameDefault = name_new   'RETURN THIS DEFAULT NAME.
End Function
Function Icon_IsNameExist( _
    proj As ProjectType, _
    gr As GroupType, _
    name_find As String) As Integer
Dim RetVal As Integer
  RetVal = Icon_GetIndex(proj, gr, name_find)
  Icon_IsNameExist = IIf(RetVal = 0, False, True)
End Function
Function Icon_GetIndex( _
    proj As ProjectType, _
    gr As GroupType, _
    name_find As String) As Integer
Dim i As Integer
Dim Found As Integer
  Found = False
  For i = 1 To gr.Icons_Count
    If (Trim$(UCase$(gr.Icons(i).Name)) = Trim$(UCase$(name_find))) Then
      Found = True: Exit For
    End If
  Next i
  Icon_GetIndex = IIf(Found, i, 0)
End Function
Sub Icon_SetDefaults(proj As ProjectType, ic As IconType)
  'ic.Name = ...   'NOTE: THE .Name FIELD IS NOT SET HERE!
  ic.LongName = "Longer Title of Application"
  ic.DescriptionText = "Very long (up to about 100 words) " & _
    "description of this application."
  ic.fn_IconImage = FN_ICONIMAGE_DEFAULT  '"default.ico"
  ic.fn_ApplicationLink = ""
  ic.fn_ApplicationLink_Dir = ""
End Sub


Function Group_GetNewNameDefault( _
    proj As ProjectType, _
    ta As TabType) As String
Dim name_new As String
Dim i As Integer   'DETERMINE DEFAULT NAME OF NEW RECORD.
  i = 1
  Do While (1 = 1)
    name_new = "Group " & Trim$(Str$(i))
    If (Not Group_IsNameExist(proj, ta, name_new)) Then Exit Do
    i = i + 1
  Loop
  Group_GetNewNameDefault = name_new   'RETURN THIS DEFAULT NAME.
End Function
Function Group_IsNameExist( _
    proj As ProjectType, _
    ta As TabType, _
    name_find As String) As Integer
Dim RetVal As Integer
  RetVal = Group_GetIndex(proj, ta, name_find)
  Group_IsNameExist = IIf(RetVal = 0, False, True)
End Function
Function Group_GetIndex( _
    proj As ProjectType, _
    ta As TabType, _
    name_find As String) As Integer
Dim i As Integer
Dim Found As Integer
  Found = False
  For i = 1 To ta.Groups_Count
    If (Trim$(UCase$(ta.Groups(i).Name)) = Trim$(UCase$(name_find))) Then
      Found = True: Exit For
    End If
  Next i
  Group_GetIndex = IIf(Found, i, 0)
End Function
Sub Group_SetDefaults(proj As ProjectType, gr As GroupType)
  'gr.Name = ...   'NOTE: THE .Name FIELD IS NOT SET HERE!
  gr.Icons_Count = 0
  gr.GroupBackgroundColor.Color = &H808000
  gr.GroupForegroundColor.Color = &HFFFFC0
  Call FontInfo_AssignDefault(gr.GroupTitleFont)
  Call FontInfo_AssignDefault(gr.GroupIconFont)
  gr.Pos.Left = 1000
  gr.Pos.Top = 1000
  gr.Pos.Width = 4000
  gr.Pos.Height = 3000
End Sub


Function Tab_GetNewNameDefault( _
    proj As ProjectType) As String
Dim name_new As String
Dim i As Integer   'DETERMINE DEFAULT NAME OF NEW RECORD.
  i = 1
  Do While (1 = 1)
    name_new = "Tab " & Trim$(Str$(i))
    If (Not Tab_IsNameExist(proj, name_new)) Then Exit Do
    i = i + 1
  Loop
  Tab_GetNewNameDefault = name_new   'RETURN THIS DEFAULT NAME.
End Function
Function Tab_IsNameExist( _
    proj As ProjectType, _
    name_find As String) As Integer
Dim RetVal As Integer
  RetVal = Tab_GetIndex(proj, name_find)
  Tab_IsNameExist = IIf(RetVal = 0, False, True)
End Function
Function Tab_GetIndex( _
    proj As ProjectType, _
    name_find As String) As Integer
Dim i As Integer
Dim Found As Integer
  Found = False
  For i = 1 To proj.Tabs_Count
    If (Trim$(UCase$(proj.Tabs(i).Name)) = Trim$(UCase$(name_find))) Then
      Found = True: Exit For
    End If
  Next i
  Tab_GetIndex = IIf(Found, i, 0)
End Function
Sub Tab_SetDefaults(proj As ProjectType, ta As TabType)
  'ta.Name = ...   'NOTE: THE .Name FIELD IS NOT SET HERE!
  ta.Groups_Count = 0
  ta.fn_BackgroundImage = ""
  ta.TabBackgroundColor.Color = &H808000
End Sub


Sub Project_ReadIniVariables(proj As ProjectType)
  proj.lvGroups_View = CInt(Val(INI_GetSetting0_Defaults(fn_OldFileList, _
    "ToolsOptions_Display", "lvGroups_View", "0")))
  proj.lvGroups_Arrange = CInt(Val(INI_GetSetting0_Defaults(fn_OldFileList, _
    "ToolsOptions_Display", "lvGroups_Arrange", "2")))
  proj.ShowDescriptionText = CBool(Val(INI_GetSetting0_Defaults(fn_OldFileList, _
    "ToolsOptions_Display", "ShowDescriptionText", "-1")))
  proj.MinimizeOnApplicationExecution = CBool(Val(INI_GetSetting0_Defaults(fn_OldFileList, _
    "ToolsOptions_Display", "MinimizeOnApplicationExecution", "-1")))
  proj.StartupMode = CInt(Val(INI_GetSetting0_Defaults(fn_OldFileList, _
    "ToolsOptions_Display", "StartupMode", Trim$(Str$(frmMain_MODE_USER)))))
  proj.DisplayUninstalledApplications = CBool(Val(INI_GetSetting0_Defaults(fn_OldFileList, _
    "ToolsOptions_Display", "DisplayUninstalledApplications", "0")))    'defaults to FALSE
End Sub
Sub Project_WriteIniVariables(proj As ProjectType)
  Call ini_putsetting0(fn_OldFileList, "ToolsOptions_Display", _
    "lvGroups_View", Trim$(Str$(proj.lvGroups_View)))
  Call ini_putsetting0(fn_OldFileList, "ToolsOptions_Display", _
    "lvGroups_Arrange", Trim$(Str$(proj.lvGroups_Arrange)))
  Call ini_putsetting0(fn_OldFileList, "ToolsOptions_Display", _
    "ShowDescriptionText", Trim$(Str$(CInt(proj.ShowDescriptionText))))
  Call ini_putsetting0(fn_OldFileList, "ToolsOptions_Display", _
    "MinimizeOnApplicationExecution", Trim$(Str$(CInt(proj.MinimizeOnApplicationExecution))))
  'NOTE: DO NOT ALLOW THE USER TO SAVE THE StartupMode VARIABLE AS frmMain_MODE_DESIGN.
  'FORCE THIS VARIABLE TO frmMain_MODE_USER.
  Call ini_putsetting0(fn_OldFileList, "ToolsOptions_Display", _
    "StartupMode", Trim$(Str$(frmMain_MODE_USER)))
  'Call ini_putsetting0(fn_OldFileList, "ToolsOptions_Display", _
  '  "StartupMode", Trim$(Str$(proj.StartupMode)))
  Call ini_putsetting0(fn_OldFileList, "ToolsOptions_Display", _
    "DisplayUninstalledApplications", Trim$(Str$(CInt(proj.DisplayUninstalledApplications))))
End Sub
Sub Project_New(proj As ProjectType)
Dim ta As TabType
Dim gr As GroupType
Dim ic As IconType
  proj.dirty = False
  proj.Pos.Width = 9600
  proj.Pos.Height = 8500
  proj.Pos.Left = (Screen.Width - proj.Pos.Width) / 2#
  'proj.Pos.Top = (Screen.Height - proj.Pos.Height) / 2#
  proj.Pos.Top = 100  '(Screen.Height - proj.Pos.Height) / 2#
  proj.lvGroups_View = 0        'ICONS!
  proj.lvGroups_Arrange = 2     'TOP!
  proj.ShowDescriptionText = True
  proj.MinimizeOnApplicationExecution = True
  proj.DisplayUninstalledApplications = False
  proj.Tabs_Count = 2
  ReDim proj.Tabs(1 To 2)
  'TAB #1 : CPAS MAIN TOOLS.
  ta.Name = "CPAS Main Tools"
  ReDim ta.Groups(1 To 1)
  ta.Groups_Count = 1
  ta.fn_BackgroundImage = "main.bmp"
  ta.TabBackgroundColor.Color = &H808000
  gr.Name = "SIMULATE"
  'SET UP THREE ICONS.
  gr.Icons_Count = 3
  ReDim gr.Icons(1 To 3)
  ic.Name = "AdDesignS"
  ic.LongName = "Adsorption Design Software"
  ic.DescriptionText = "Adsorption Design Software (AdDesignS) " & _
    "provides sizing and performance estimations for single " & _
    "and multiple component adsorption of gas and liquid " & _
    "phase streams to fixed bed adsorbers."
  ic.fn_IconImage = "ads.bmp"
  ic.fn_ApplicationLink = "<C>\ads\adss.exe"
  gr.Icons(1) = ic
  ic.Name = "ASAP"
  ic.LongName = "Aeration Simulation Analysis Program"
  ic.DescriptionText = "Aeration Systems Analysis Program (ASAP) " & _
    "provides sizing and performance estimations for " & _
    "packed tower, bubble, and surface aeration systems " & _
    "for the removal of volatile organic compounds (VOCs) " & _
    "from aqueous streams."
  ic.fn_IconImage = "asap.bmp"
  ic.fn_ApplicationLink = "<C>\asap\asap.exe"
  gr.Icons(2) = ic
  ic.Name = "FaVOr"
  ic.LongName = "Fate of Volatile Organics in Wastewater Treatment Plants"
  ic.DescriptionText = "Fate of Volatile Organics in Wastewater " & _
    "Treatment Plants (FaVOr) predicts the fate of volatile " & _
    "organic compounds in activated sludge wastewater " & _
    "treatment plants."
  ic.fn_IconImage = "favor.bmp"
  ic.fn_ApplicationLink = "<C>\favor\favor.exe"
  gr.Icons(3) = ic
  'SET UP THREE ICONS (END).
  gr.GroupBackgroundColor.Color = &H808000
  gr.GroupForegroundColor.Color = &HFFFFC0
  Call FontInfo_AssignDefault(gr.GroupTitleFont)
  Call FontInfo_AssignDefault(gr.GroupIconFont)
  gr.Pos.Left = 5500
  gr.Pos.Top = 400
  gr.Pos.Width = 3600
  gr.Pos.Height = 2450
  ta.Groups(1) = gr
  proj.Tabs(1) = ta
  'TAB #2 : MISCELLANEOUS DOCUMENTATION.
  ta.Name = "Miscellaneous Documentation"
  ReDim ta.Groups(1 To 1)
  ta.Groups_Count = 1
  ta.fn_BackgroundImage = ""   '"default.bmp"
  ta.TabBackgroundColor.Color = &H808000
  gr.Name = "DOCUMENTATION"
  gr.Icons_Count = 0
  gr.GroupBackgroundColor.Color = &H808000
  gr.GroupForegroundColor.Color = &HFFFFC0
  Call FontInfo_AssignDefault(gr.GroupTitleFont)
  Call FontInfo_AssignDefault(gr.GroupIconFont)
  gr.Pos.Left = 3000
  gr.Pos.Top = 1400
  gr.Pos.Width = 4000
  gr.Pos.Height = 3000
  ta.Groups(1) = gr
  proj.Tabs(2) = ta
End Sub
Sub Project_DirtyFlag_Clear(ByRef proj As ProjectType)
  proj.dirty = False
  Call Project_DirtyFlag_RefreshScreen(proj)
End Sub
Sub Project_DirtyFlag_RefreshScreen(proj As ProjectType)
  If (proj.dirty) Then
    frmMain.sspanel_Dirty = ""
    frmMain.sspanel_Dirty.ForeColor = QBColor(4 + 8)
    frmMain.sspanel_Dirty = "Data Changed"
  Else
    frmMain.sspanel_Dirty = ""
    frmMain.sspanel_Dirty.ForeColor = QBColor(0)
    frmMain.sspanel_Dirty = "Data Unchanged"
  End If
End Sub
Sub Project_DirtyFlag_Throw(ByRef proj As ProjectType)
  proj.dirty = True
  Call Project_DirtyFlag_RefreshScreen(proj)
End Sub


Function INI_GetSetting0_Defaults( _
    fn_ini As String, _
    ini_header As String, _
    INI_VarName As String, _
    use_default_if_null As String) As String
Dim retstr As String
  retstr = Trim$(INI_GetSetting0(fn_ini, ini_header, INI_VarName))
  If (retstr = "") Then retstr = use_default_if_null
  INI_GetSetting0_Defaults = retstr
End Function


