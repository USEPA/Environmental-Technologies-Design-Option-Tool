Attribute VB_Name = "FileIO"
Option Explicit





Const FileIO_declarations_end = 0


Sub PositionInfo_ProjectFile_Write(f As Integer, pi As PositionInfoType)
  Call ProjectFile_Write(f, pi.Left, "pi.Left")
  Call ProjectFile_Write(f, pi.Top, "pi.Top")
  Call ProjectFile_Write(f, pi.Width, "pi.Width")
  Call ProjectFile_Write(f, pi.Height, "pi.Height")
End Sub
Sub PositionInfo_ProjectFile_Read(f As Integer, pi As PositionInfoType)
  Call ProjectFile_Read(f, pi.Left, "pi.Left")
  Call ProjectFile_Read(f, pi.Top, "pi.Top")
  Call ProjectFile_Read(f, pi.Width, "pi.Width")
  Call ProjectFile_Read(f, pi.Height, "pi.Height")
End Sub


Sub FontInfo_ProjectFile_Write(f As Integer, fi As FontInfoType)
  Call ProjectFile_Write(f, fi.FontBold, "fi.FontBold")
  Call ProjectFile_Write(f, fi.FontItalic, "fi.FontItalic")
  Call ProjectFile_Write(f, fi.FontName, "fi.FontName")
  Call ProjectFile_Write(f, fi.FontSize, "fi.FontSize")
  Call ProjectFile_Write(f, fi.FontStrikeThru, "fi.FontStrikeThru")
  Call ProjectFile_Write(f, fi.FontUnderline, "fi.FontUnderline")
End Sub
Sub FontInfo_ProjectFile_Read(f As Integer, fi As FontInfoType)
  Call ProjectFile_Read(f, fi.FontBold, "fi.FontBold")
  Call ProjectFile_Read(f, fi.FontItalic, "fi.FontItalic")
  Call ProjectFile_Read(f, fi.FontName, "fi.FontName")
  Call ProjectFile_Read(f, fi.FontSize, "fi.FontSize")
  Call ProjectFile_Read(f, fi.FontStrikeThru, "fi.FontStrikeThru")
  Call ProjectFile_Read(f, fi.FontUnderline, "fi.FontUnderline")
End Sub


Sub Icon_ProjectFile_Write(f As Integer, ic As IconType)
  Call ProjectFile_Write(f, ic.Name, "ic.Name")
  Call ProjectFile_Write(f, ic.LongName, "ic.LongName")
  Call ProjectFile_Write(f, ic.DescriptionText, "ic.DescriptionText")
  Call ProjectFile_Write(f, ic.fn_IconImage, "ic.fn_IconImage")
  Call ProjectFile_Write(f, ic.fn_ApplicationLink, "ic.fn_ApplicationLink")
  Call ProjectFile_Write(f, ic.fn_ApplicationLink_Dir, "ic.fn_ApplicationLink_Dir")
End Sub
Sub Icon_ProjectFile_Read(f As Integer, ic As IconType)
  Call ProjectFile_Read(f, ic.Name, "ic.Name")
  Call ProjectFile_Read(f, ic.LongName, "ic.LongName")
  Call ProjectFile_Read(f, ic.DescriptionText, "ic.DescriptionText")
  Call ProjectFile_Read(f, ic.fn_IconImage, "ic.fn_IconImage")
  Call ProjectFile_Read(f, ic.fn_ApplicationLink, "ic.fn_ApplicationLink")
  Call ProjectFile_Read(f, ic.fn_ApplicationLink_Dir, "ic.fn_ApplicationLink_Dir")
End Sub


Sub Group_ProjectFile_Write(f As Integer, gr As GroupType)
Dim i As Integer
  Call ProjectFile_Write(f, gr.Name, "gr.Name")
  Call ProjectFile_Write(f, gr.Icons_Count, "gr.Icons_Count")
  For i = 1 To gr.Icons_Count
    Call Icon_ProjectFile_Write(f, gr.Icons(i))
  Next i
  Call ProjectFile_Write(f, gr.GroupBackgroundColor.Color, "gr.GroupBackgroundColor.Color")
  Call ProjectFile_Write(f, gr.GroupForegroundColor.Color, "gr.GroupForegroundColor.Color")
  Call FontInfo_ProjectFile_Write(f, gr.GroupTitleFont)
  Call FontInfo_ProjectFile_Write(f, gr.GroupIconFont)
  Call PositionInfo_ProjectFile_Write(f, gr.Pos)
End Sub
Sub Group_ProjectFile_Read(f As Integer, gr As GroupType)
Dim i As Integer
  Call ProjectFile_Read(f, gr.Name, "gr.Name")
  Call ProjectFile_Read(f, gr.Icons_Count, "gr.Icons_Count")
  If (gr.Icons_Count >= 1) Then
    ReDim gr.Icons(1 To gr.Icons_Count)
  End If
  For i = 1 To gr.Icons_Count
    Call Icon_ProjectFile_Read(f, gr.Icons(i))
  Next i
  Call ProjectFile_Read(f, gr.GroupBackgroundColor.Color, "gr.GroupBackgroundColor.Color")
  Call ProjectFile_Read(f, gr.GroupForegroundColor.Color, "gr.GroupForegroundColor.Color")
  Call FontInfo_ProjectFile_Read(f, gr.GroupTitleFont)
  Call FontInfo_ProjectFile_Read(f, gr.GroupIconFont)
  Call PositionInfo_ProjectFile_Read(f, gr.Pos)
End Sub


Sub Tab_ProjectFile_Write(f As Integer, ta As TabType)
Dim i As Integer
  Call ProjectFile_Write(f, ta.Name, "ta.Name")
  Call ProjectFile_Write(f, ta.Groups_Count, "ta.Groups_Count")
  For i = 1 To ta.Groups_Count
    Call Group_ProjectFile_Write(f, ta.Groups(i))
  Next i
  Call ProjectFile_Write(f, ta.fn_BackgroundImage, "ta.fn_BackgroundImage")
  Call ProjectFile_Write(f, ta.TabBackgroundColor.Color, "ta.TabBackgroundColor.Color")
End Sub
Sub Tab_ProjectFile_Read(f As Integer, ta As TabType)
Dim i As Integer
  Call ProjectFile_Read(f, ta.Name, "ta.Name")
  Call ProjectFile_Read(f, ta.Groups_Count, "ta.Groups_Count")
  If (ta.Groups_Count >= 1) Then
    ReDim ta.Groups(1 To ta.Groups_Count)
  End If
  For i = 1 To ta.Groups_Count
    Call Group_ProjectFile_Read(f, ta.Groups(i))
  Next i
  Call ProjectFile_Read(f, ta.fn_BackgroundImage, "ta.fn_BackgroundImage")
  Call ProjectFile_Read(f, ta.TabBackgroundColor.Color, "ta.TabBackgroundColor.Color")
End Sub


'RETURNS:
'   - TRUE = SAVE WENT OK.
'   - FALSE = SAVE DID NOT GO OK.
Function MainDatafile_Output(proj As ProjectType) As Boolean
Dim f As Integer
Dim i As Integer
  'STORE CURRENT WINDOW WIDTH/HEIGHT.
  proj.Pos.Width = frmMain.Width
  proj.Pos.Height = frmMain.Height
  
  'STORE THE DATA FILE.
  On Error GoTo err_MainDatafile_Output
  f = FreeFile
  Call StatusInfo_Display(frmMain.sspanel_StatusInfo, _
      "File I/O taking place, please wait ... ", _
      12, 15)
  Open fn_Full_MainDataFile For Output As #f
  Call ProjectFile_Write(f, get_program_version_with_build_info(), "Version")
  Call ProjectFile_Write(f, Now, "Save date/time")
  Call ProjectFile_Write(f, proj.Tabs_Count, "proj.Tabs_Count")
  For i = 1 To proj.Tabs_Count
    Call Tab_ProjectFile_Write(f, proj.Tabs(i))
  Next i
  Call PositionInfo_ProjectFile_Write(f, proj.Pos)
  Close #f
exit_Save_went_okay:
  'SAVE WENT OKAY.
  Call StatusInfo_Display(frmMain.sspanel_StatusInfo, "", 12, 15)
  MainDatafile_Output = True
  Exit Function
exit_Save_did_not_go_okay:
  'SAVE DID NOT GO OKAY.
  Call StatusInfo_Display(frmMain.sspanel_StatusInfo, "", 12, 15)
  MainDatafile_Output = False
  Exit Function
err_MainDatafile_Output:
  Call Show_Trapped_Error("MainDatafile_Output")
  Resume exit_Save_did_not_go_okay
End Function
'RETURNS:
'   - TRUE = LOAD WENT OK.
'   - FALSE = LOAD DID NOT GO OK.
Function MainDatafile_Input(proj As ProjectType) As Boolean
Dim f As Integer
Dim i As Integer
Dim Dummy As String
  On Error GoTo err_MainDatafile_Input
  f = FreeFile
  Call StatusInfo_Display(frmMain.sspanel_StatusInfo, _
      "File I/O taking place, please wait ... ", _
      12, 15)
  Open fn_Full_MainDataFile For Input As #f
  Call ProjectFile_Read(f, Dummy, "Version")
  Call ProjectFile_Read(f, Dummy, "Save date/time")
  Call ProjectFile_Read(f, proj.Tabs_Count, "proj.Tabs_Count")
  If (proj.Tabs_Count >= 1) Then
    ReDim proj.Tabs(1 To proj.Tabs_Count)
  End If
  For i = 1 To proj.Tabs_Count
    Call Tab_ProjectFile_Read(f, proj.Tabs(i))
  Next i
  Call PositionInfo_ProjectFile_Read(f, proj.Pos)
  Close #f
exit_Load_went_okay:
  'LOAD WENT OKAY.
  MainDatafile_Input = True
  Call StatusInfo_Display(frmMain.sspanel_StatusInfo, "", 12, 15)
  Exit Function
exit_Load_did_not_go_okay:
  'LOAD DID NOT GO OKAY.
  MainDatafile_Input = False
  Call StatusInfo_Display(frmMain.sspanel_StatusInfo, "", 12, 15)
  Exit Function
err_MainDatafile_Input:
  Call Show_Trapped_Error("MainDatafile_Input")
  Resume exit_Load_did_not_go_okay
End Function


Sub ProjectFile_Read(f As Integer, ByRef RetVal As Variant, Optional optDummy1 As Variant)
Dim outputstr$
Dim outlin As String
Dim sub_name As String
Dim input1 As String
Dim input2 As String
  Input #f, input1, input2
  sub_name = "ProjectFile_Read"
  Select Case VarType(RetVal)
    Case vbBoolean
      RetVal = Val(input1)
    Case vbByte, vbInteger, vbLong, vbCurrency
      RetVal = Val(input1)
    Case vbSingle, vbDouble
      RetVal = Val(input1)
    Case vbString, vbDate
      RetVal = input1
    Case vbObject
        MsgBox sub_name & " vbObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbError
        MsgBox sub_name & " vbError not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbDataObject
        MsgBox sub_name & " vbDataObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbVariant
        MsgBox sub_name & " vbVariant not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbArray
        MsgBox sub_name & " vbArray not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbEmpty
        MsgBox sub_name & " vbEmpty not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbNull
        MsgBox sub_name & " vbNull not implemented"
        GoTo EXIT_FALSE_VALUE
  End Select
  GoTo EXIT_OK
EXIT_FALSE_VALUE:
  Print #f, "   - - - ERROR IN " & sub_name & "() - - -"
  Exit Sub
EXIT_OK:
  Exit Sub
End Sub
Sub ProjectFile_Write(f As Integer, v As Variant, s As String)
Dim outputstr$
Dim outlin As String
Dim sub_name As String
  sub_name = "ProjectFile_Write"
  Select Case VarType(v)
    Case vbBoolean
        outputstr$ = IIf(v, "1", "0")
    Case vbByte, vbInteger, vbLong, vbCurrency
        outputstr$ = Trim$(CStr(v))
    Case vbSingle, vbDouble
        outputstr$ = Trim$(CStr(v))
    Case vbString, vbDate
        outputstr$ = CStr(v)
    Case vbObject
        MsgBox sub_name & " vbObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbError
        MsgBox sub_name & " vbError not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbDataObject
        MsgBox sub_name & " vbDataObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbVariant
        MsgBox sub_name & " vbVariant not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbArray
        MsgBox sub_name & " vbArray not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbEmpty
        MsgBox sub_name & " vbEmpty not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbNull
        MsgBox sub_name & " vbNull not implemented"
        GoTo EXIT_FALSE_VALUE
  End Select
  outlin = Chr$(34) & Trim$(outputstr$) & Chr$(34) & "," & _
      Chr$(34) & s & Chr$(34)
  'outlin = Trim$(outputstr$)
  'If (Len(outlin) > 27) Then
  '  outlin = outlin & "    "
  'Else
  '  Do While (1 = 1)
  '    If (Len(outlin) >= 27) Then Exit Do
  '    outlin = outlin & " "
  '  Loop
  'End If
  'outlin = outlin & s
  Print #f, outlin
  GoTo EXIT_OK
EXIT_FALSE_VALUE:
  Print #f, "   - - - ERROR IN " & sub_name & "() - - -"
  Exit Sub
EXIT_OK:
  Exit Sub
End Sub


'RETURNS:
'- true = it's okay to unload this file now.
'- false = cancel the unload.
Function file_query_unload(proj As ProjectType) As Integer
Dim RetVal As Integer
Dim msg As String
  If (proj.dirty = False) Then
    file_query_unload = True
    Exit Function
  End If
  msg = "Do you want to save the changes you made to the design " & _
      "of this workspace ?"
  RetVal = MsgBox(msg, vbCritical + vbYesNoCancel, App.Title)
  Select Case RetVal
    Case vbYes:
      If (MainDatafile_Output(proj) = True) Then
        'SAVE WENT OK; IT'S NOW OKAY TO UNLOAD THIS FILE.
        file_query_unload = True
      Else
        'SAVE FAILED; DON'T UNLOAD THIS FILE.
        file_query_unload = False
      End If
      Exit Function
    Case vbNo:
      file_query_unload = True
      Exit Function
    Case vbCancel:
      file_query_unload = False
      Exit Function
  End Select
End Function

