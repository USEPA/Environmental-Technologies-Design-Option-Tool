Attribute VB_Name = "LastFewMod"
Option Explicit


Type LastFewFilesType
  WhichApp As Integer
  WhichForm As Integer
  MenuIDNum_LastFewStartsAt As Integer
  MenuIDNum_FinalSeparator As Integer
  FileNames(1 To 4) As String
  INI_VariablePrefix As String
End Type

Global Current_LastFewFilesRec As LastFewFilesType

Global Const LASTFEW_WHICHAPP_STEPP = 1
Global Const LASTFEW_WHICHAPP_ASAP = 2
Global Const LASTFEW_WHICHAPP_ADSIM = 3

Global Const LASTFEW_STEPP_contam_prop_form = 101
Global Const LASTFEW_ASAP_frmPTADScreen1 = 201
Global Const LASTFEW_ASAP_frmPTADScreen2 = 202
Global Const LASTFEW_ASAP_frmBubble_DESIGN = 203
Global Const LASTFEW_ASAP_frmBubble_RATING = 204
Global Const LASTFEW_ASAP_frmSurface_DESIGN = 205
Global Const LASTFEW_ASAP_frmSurface_RATING = 206
Global Const LASTFEW_ADSIM_frmPFPSDM = 301

Sub LastFewFiles_ChangeCaption(MenuItemID As Integer, ChangeTo As String)

  Call LastFewFiles_ChangeSomething(MenuItemID, "c", ChangeTo, 0)

End Sub

Sub LastFewFiles_ChangeSomething(MenuItemID As Integer, ChangeWhat As String, StrParam1 As String, IntParam1 As Integer)
Dim mm As Menu

  Select Case Current_LastFewFilesRec.WhichApp
    'Case LASTFEW_WHICHAPP_STEPP
    '  Select Case Current_LastFewFilesRec.WhichForm
    '    Case LASTFEW_STEPP_contam_prop_form
    '      Set mm = contam_prop_form!mnuFile(MenuItemID)
    '  End Select
    Case LASTFEW_WHICHAPP_ASAP
      Select Case Current_LastFewFilesRec.WhichForm
        Case LASTFEW_ASAP_frmPTADScreen1
          Set mm = frmPTADScreen1!mnuFile(MenuItemID)
        Case LASTFEW_ASAP_frmPTADScreen2
          Set mm = frmPTADScreen2!mnuFile(MenuItemID)
        Case LASTFEW_ASAP_frmBubble_DESIGN
          Set mm = frmBubble!mnuFile(MenuItemID)
        Case LASTFEW_ASAP_frmBubble_RATING
          Set mm = frmBubble!mnuFile(MenuItemID)
        Case LASTFEW_ASAP_frmSurface_DESIGN
          Set mm = frmSurface!mnuFile(MenuItemID)
        Case LASTFEW_ASAP_frmSurface_RATING
          Set mm = frmSurface!mnuFile(MenuItemID)
      End Select
    'Case LASTFEW_WHICHAPP_ADSIM
    '  Select Case Current_LastFewFilesRec.WhichForm
    '    Case LASTFEW_ADSIM_frmPFPSDM
    '      Set mm = frmPFPSDM!mnuFileItem(MenuItemID)
    '  End Select
  End Select

  Call LastFewFiles_ChangeSomething0(mm, ChangeWhat, StrParam1, IntParam1)

End Sub

Sub LastFewFiles_ChangeSomething0(mm As Menu, ChangeWhat As String, StrParam1 As String, IntParam1 As Integer)

  ''''On Error Resume Next
  If (UCase$(ChangeWhat) = "C") Then
    mm.Caption = StrParam1
  ElseIf (UCase$(ChangeWhat) = "V") Then
    mm.Visible = IntParam1
  Else
    'Do nothing.
  End If

End Sub

Sub LastFewFiles_ChangeVisibility(MenuItemID As Integer, ChangeTo As Integer)

  Call LastFewFiles_ChangeSomething(MenuItemID, "v", "", ChangeTo)

End Sub

Sub LastFewFiles_DisplayList()
Dim i As Integer
Dim j As Integer
Dim NumVisible As Integer
Dim NewCaption As String

  NumVisible = 0
  For i = 1 To 4
    j = Current_LastFewFilesRec.MenuIDNum_LastFewStartsAt + i - 1
    If (Current_LastFewFilesRec.FileNames(i) <> "") Then
      NewCaption = "&" & Trim$(Str$(i)) & " " & Current_LastFewFilesRec.FileNames(i)
      Call LastFewFiles_ChangeCaption(j, NewCaption)

      'Current_LastFewFilesRec.FileNames(i))
      Call LastFewFiles_ChangeVisibility(j, True)
      NumVisible = NumVisible + 1
    Else
      Call LastFewFiles_ChangeCaption(j, "")
      Call LastFewFiles_ChangeVisibility(j, False)
    End If
  Next i
  
  If (NumVisible = 0) Then
    Call LastFewFiles_ChangeVisibility(Current_LastFewFilesRec.MenuIDNum_FinalSeparator, False)
  Else
    Call LastFewFiles_ChangeVisibility(Current_LastFewFilesRec.MenuIDNum_FinalSeparator, True)
  End If

End Sub

Sub LastFewFiles_InitializeList(WhichApp As Integer, WhichForm As Integer)
Dim i As Integer
Dim thisvarname As String

  Current_LastFewFilesRec.WhichApp = WhichApp
  Current_LastFewFilesRec.WhichForm = WhichForm
  Current_LastFewFilesRec.MenuIDNum_LastFewStartsAt = 191
  Current_LastFewFilesRec.MenuIDNum_FinalSeparator = 199

  Select Case WhichApp
    Case LASTFEW_WHICHAPP_STEPP
      Select Case WhichForm
        Case LASTFEW_STEPP_contam_prop_form
          Current_LastFewFilesRec.INI_VariablePrefix = "MAIN"
      End Select
    Case LASTFEW_WHICHAPP_ASAP
      Select Case WhichForm
        Case LASTFEW_ASAP_frmPTADScreen1
          Current_LastFewFilesRec.INI_VariablePrefix = "PTAD1"
        Case LASTFEW_ASAP_frmPTADScreen2
          Current_LastFewFilesRec.INI_VariablePrefix = "PTAD2"
        Case LASTFEW_ASAP_frmBubble_DESIGN
          Current_LastFewFilesRec.INI_VariablePrefix = "BUB1"
        Case LASTFEW_ASAP_frmBubble_RATING
          Current_LastFewFilesRec.INI_VariablePrefix = "BUB2"
        Case LASTFEW_ASAP_frmSurface_DESIGN
          Current_LastFewFilesRec.INI_VariablePrefix = "SUR1"
        Case LASTFEW_ASAP_frmSurface_RATING
          Current_LastFewFilesRec.INI_VariablePrefix = "SUR2"
      End Select
    Case LASTFEW_WHICHAPP_ADSIM
      Select Case WhichForm
        Case LASTFEW_ADSIM_frmPFPSDM
          Current_LastFewFilesRec.INI_VariablePrefix = "MAIN"
      End Select
  End Select

  For i = 1 To 4
    thisvarname = Current_LastFewFilesRec.INI_VariablePrefix & "_OldFile" & Trim$(Str$(i))
    Current_LastFewFilesRec.FileNames(i) = Trim$(INI_Getsetting(thisvarname))
  Next i

  'Update display from internal memory.
  Call LastFewFiles_DisplayList

End Sub

Sub LastFewFiles_MoveFilenameToTop(fn As String)
Dim Found As Integer
Dim i As Integer
Dim fn_This As String
Dim thisvarname As String

  Found = 0
  For i = 1 To 4
    fn_This = Trim$(Current_LastFewFilesRec.FileNames(i))
    If (fn_This <> "") Then
      If (UCase$(fn_This) = UCase$(fn)) Then
        Found = i
        Exit For
      End If
    End If
  Next i

  If (Found <> 0) Then
    For i = Found - 1 To 1 Step -1
      Current_LastFewFilesRec.FileNames(i + 1) = Current_LastFewFilesRec.FileNames(i)
    Next i
  Else
    For i = 3 To 1 Step -1
      Current_LastFewFilesRec.FileNames(i + 1) = Current_LastFewFilesRec.FileNames(i)
    Next i
  End If
  Current_LastFewFilesRec.FileNames(1) = UCase$(Trim$(fn))

  'Update display from internal memory.
  Call LastFewFiles_DisplayList

  'Update the .INI file.
  For i = 1 To 4
    thisvarname = Current_LastFewFilesRec.INI_VariablePrefix & "_OldFile" & Trim$(Str$(i))
    Call INI_PutSetting(thisvarname, UCase$(Trim$(Current_LastFewFilesRec.FileNames(i))))
  Next i

End Sub

