Attribute VB_Name = "IniFile_Preferences"
Option Explicit





Const IniFile_Preferences_decl_end = True


Function INI_GetSetting_Into_Var( _
    in_HeaderName As String, _
    in_VariableName As String, _
    out_StoreVar As Variant) _
    As Boolean
On Error GoTo err_ThisFunc
Dim sReturn As String
  sReturn = INI_GetSetting0( _
      fn_OldFileList, _
      in_HeaderName, _
      in_VariableName)
  sReturn = Trim$(sReturn)
  If (sReturn = "") Then
    '
    ' DO NOTHING!
    '
  Else
    '
    ' STORE THIS DATA.
    '
    Select Case VarType(out_StoreVar)
      Case vbBoolean:
        out_StoreVar = CBool(Val(sReturn))
      Case vbByte:
        out_StoreVar = CByte(Val(sReturn))
      Case vbInteger:
        out_StoreVar = CInt(Val(sReturn))
      Case vbLong:
        out_StoreVar = CLng(Val(sReturn))
      Case vbString, vbDate:
        out_StoreVar = CStr(sReturn)
      Case vbDouble:
        out_StoreVar = CDbl(Val(sReturn))
      Case vbSingle:
        out_StoreVar = CSng(Val(sReturn))
    End Select
  End If
exit_normally_ThisFunc:
  INI_GetSetting_Into_Var = True
  Exit Function
exit_err_ThisFunc:
  INI_GetSetting_Into_Var = False
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("PrefEnvironment_SetDefaults")
  Resume exit_err_ThisFunc
End Function
Function INI_PutSetting_From_Var( _
    in_HeaderName As String, _
    in_VariableName As String, _
    in_StoreVar As Variant) _
    As Boolean
On Error GoTo err_ThisFunc
Dim sWrite As String
  '
  ' TRANSFER DATA TO STRING VARIABLE.
  '
  Select Case VarType(in_StoreVar)
    Case vbBoolean, vbByte, vbInteger, vbLong, vbDouble, vbSingle:
      sWrite = Trim$(Str$(in_StoreVar))
    Case vbString, vbDate:
      sWrite = Trim$(in_StoreVar)
  End Select
  Call ini_putsetting0( _
      fn_OldFileList, _
      in_HeaderName, _
      in_VariableName, _
      sWrite)
exit_normally_ThisFunc:
  INI_PutSetting_From_Var = True
  Exit Function
exit_err_ThisFunc:
  INI_PutSetting_From_Var = False
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("INI_PutSetting_From_Var")
  Resume exit_err_ThisFunc
End Function


Function PrefEnvironment_SetDefaults() _
    As Boolean
On Error GoTo err_ThisFunc
  With PrefEnvironment
    '
    ' NUMERICAL DISPLAY FORMAT.
    '
    .NumFormat_Greater1000 = NUMFORMAT_4SIGFIG
    .NumFormat_Less0_001 = NUMFORMAT_4SIGFIG
    .NumFormat_Other = NUMFORMAT_4SIGFIG
    '
    ' MISCELLANEOUS.
    '
    .FontSize_Lists = 8
  End With
exit_normally_ThisFunc:
  PrefEnvironment_SetDefaults = True
  Exit Function
exit_err_ThisFunc:
  PrefEnvironment_SetDefaults = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("PrefEnvironment_SetDefaults")
  Resume exit_err_ThisFunc
End Function
Function PrefEnvironment_LoadFromINI() _
    As Boolean
On Error GoTo err_ThisFunc
  With PrefEnvironment
    '
    ' NUMERICAL DISPLAY FORMAT.
    '
    Call INI_GetSetting_Into_Var("PrefEnvironment", "NumFormat_Greater1000", .NumFormat_Greater1000)
    Call INI_GetSetting_Into_Var("PrefEnvironment", "NumFormat_Less0_001", .NumFormat_Less0_001)
    Call INI_GetSetting_Into_Var("PrefEnvironment", "NumFormat_Other", .NumFormat_Other)
    '
    ' MISCELLANEOUS.
    '
    Call INI_GetSetting_Into_Var("PrefEnvironment", "FontSize_Lists", .FontSize_Lists)
  End With
exit_normally_ThisFunc:
  PrefEnvironment_LoadFromINI = True
  Exit Function
exit_err_ThisFunc:
  PrefEnvironment_LoadFromINI = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("PrefEnvironment_LoadFromINI")
  Resume exit_err_ThisFunc
End Function
Function PrefEnvironment_SaveToINI() _
    As Boolean
On Error GoTo err_ThisFunc
  With PrefEnvironment
    '
    ' NUMERICAL DISPLAY FORMAT.
    '
    Call INI_PutSetting_From_Var("PrefEnvironment", "NumFormat_Greater1000", .NumFormat_Greater1000)
    Call INI_PutSetting_From_Var("PrefEnvironment", "NumFormat_Less0_001", .NumFormat_Less0_001)
    Call INI_PutSetting_From_Var("PrefEnvironment", "NumFormat_Other", .NumFormat_Other)
    '
    ' MISCELLANEOUS.
    '
    Call INI_PutSetting_From_Var("PrefEnvironment", "FontSize_Lists", .FontSize_Lists)
  End With
exit_normally_ThisFunc:
  PrefEnvironment_SaveToINI = True
  Exit Function
exit_err_ThisFunc:
  PrefEnvironment_SaveToINI = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("PrefEnvironment_SaveToINI")
  Resume exit_err_ThisFunc
End Function


