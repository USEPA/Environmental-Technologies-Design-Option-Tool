Attribute VB_Name = "MiscUI"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

Global Const CBOYAXISTYPE_C_CO = 1
Global Const CBOYAXISTYPE_UG_L = 2
Global Const CBOYAXISTYPE_MG_L = 3
Global Const CBOYAXISTYPE_G_L = 4
Global Const CBOYAXISTYPE_PPB = 5
Global Const CBOYAXISTYPE_PPM = 6






Const MiscUI_declarations_end = True


'
' THIS FUNCTION RETURNS THE CONVERSION FACTOR THAT THE VALUE OF
' Results.CP() MUST BE MULTIPLIED BY TO GET THE VALUE IN
' THE DESIRED UNITS OF DISPLAY.
'
'      If (intIsPSDMInRoomModel = False) Then
'        'the .CP() units are C/Co
'      Else
'        If (intAnyCrCloseToZero = True) Then
'          'the .CP() units are ug/L
'        Else
'          'the .CP() units are Cr/Cr,ss
'        End If
'      End If
'
Function CBOYAXISTYPE_GetUnitConversion( _
    intCBOYAXISTYPE As Integer, _
    intIsPSDMInRoomModel As Integer, _
    intAnyCrCloseToZero As Integer, _
    intComponentNum As Integer, _
    intBedPhase As Integer, _
    OUT_strYAxisTitle As String) _
    As Double
Dim strConcName As String
Dim dblRetVal As Double
Dim dbl_Cr_ss As Double     'ug/L
Dim dbl_Co As Double        'ug/L
Dim dbl_ConvertTo_ug_L As Double      'what it takes to convert the .CP() value to ug/L
Dim dbl_ConvertFrom_ppm_To_ug_L As Double
Dim dbl_ConvertFrom_ug_L_To_ppm As Double
Dim dbl_Pressure_Pa As Double
Dim dbl_R_J_gmol_K As Double
Dim dbl_T_K As Double
Dim dbl_MolecWeight As Double
  '
  ' SET UP THE CONVERSION FACTOR FOR ug/L <===> ppm.
  '
  Select Case intBedPhase
    '
    '////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////   LIQUID PHASE
    Case 0:
      dbl_ConvertFrom_ug_L_To_ppm = 1 / 1000#
    '
    '////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////   GAS PHASE
    Case 1:
      dbl_R_J_gmol_K = 8.31451
      dbl_Pressure_Pa = Results.Bed.Pressure * 101325#
      dbl_T_K = Results.Bed.Temperature + 273.15
      dbl_MolecWeight = Results.Component(intComponentNum).MW
      dbl_ConvertFrom_ppm_To_ug_L = _
          1# / 1000000# * _
          (dbl_Pressure_Pa) / (dbl_R_J_gmol_K) / (dbl_T_K) * _
          1000000# * dbl_MolecWeight / 1000#
      dbl_ConvertFrom_ug_L_To_ppm = 1# / dbl_ConvertFrom_ppm_To_ug_L
  End Select
  '
  ' DETERMINE HOW TO CONVERT .CP() INTO ug/L.
  '
  If (intIsPSDMInRoomModel = True) Then
    strConcName = "Cr"
    dbl_Cr_ss = Results.psdmroom_Crss(intComponentNum)
    If (intAnyCrCloseToZero = True) Then
      dbl_ConvertTo_ug_L = 1#
    Else
      dbl_ConvertTo_ug_L = dbl_Cr_ss
    End If
  Else
    strConcName = "C"
    dbl_Co = 1000# * Results.Component(intComponentNum).InitialConcentration
    ' THE PREVIOUS LINE CONVERTS mg/L TO ug/L
    dbl_ConvertTo_ug_L = dbl_Co
  End If
  '
  ' THE MAIN CODE.
  '
  Select Case intCBOYAXISTYPE
    Case CBOYAXISTYPE_C_CO:
      If (intIsPSDMInRoomModel = False) Then
        OUT_strYAxisTitle = "C/Co"
      Else
        If (intAnyCrCloseToZero = True) Then
          OUT_strYAxisTitle = "(  ERROR -- UNAVAILABLE!!!  )"
        Else
          OUT_strYAxisTitle = "Cr/Cr,ss"
        End If
      End If
      dblRetVal = 1#
    Case CBOYAXISTYPE_UG_L:
      OUT_strYAxisTitle = strConcName & ", µg/L"
      dblRetVal = dbl_ConvertTo_ug_L
    Case CBOYAXISTYPE_MG_L:
      OUT_strYAxisTitle = strConcName & ", mg/L"
      dblRetVal = dbl_ConvertTo_ug_L / 1000#
    Case CBOYAXISTYPE_G_L:
      OUT_strYAxisTitle = strConcName & ", g/L"
      dblRetVal = dbl_ConvertTo_ug_L / 1000# / 1000#
    Case CBOYAXISTYPE_PPB:
      OUT_strYAxisTitle = strConcName & ", ppb"
      dblRetVal = dbl_ConvertTo_ug_L * dbl_ConvertFrom_ug_L_To_ppm * 1000#
    Case CBOYAXISTYPE_PPM:
      OUT_strYAxisTitle = strConcName & ", ppm"
      dblRetVal = dbl_ConvertTo_ug_L * dbl_ConvertFrom_ug_L_To_ppm
  End Select
  CBOYAXISTYPE_GetUnitConversion = dblRetVal
End Function

Sub ShellExecute_LocalFile( _
    in_Filename As String)
  Call ShellExecute(0&, vbNullString, in_Filename, vbNullString, vbNullString, vbNormalFocus)
End Sub
Sub ShellExecute_URL( _
    in_URL As String)
  Call ShellExecute(0&, vbNullString, in_URL, vbNullString, vbNullString, vbNormalFocus)
End Sub


Sub CalcStatus_Set(newVal As Boolean)
  If (newVal) Then
    Call GenericStatus_Set("Calculating -- please wait.")
  Else
    Call GenericStatus_Set("")
  End If
End Sub
Sub GenericStatus_Set(fn_Text As String)
  frmMain.sspanel_Status = fn_Text
End Sub
Sub DirtyStatus_Set(newVal As Boolean)
  If (IsThisADemo() = True) Then
    frmMain.sspanel_Dirty = "* DEMO VERSION *"
    frmMain.sspanel_Dirty.ForeColor = QBColor(12)
  Else
    If (newVal) Then
      frmMain.sspanel_Dirty = "Data Changed"
      frmMain.sspanel_Dirty.ForeColor = QBColor(12)
    Else
      frmMain.sspanel_Dirty = "Unchanged"
      frmMain.sspanel_Dirty.ForeColor = QBColor(0)
    End If
  End If
End Sub
Sub DirtyStatus_Set_Current()
  Call DirtyStatus_Set(Project_Is_Dirty)
End Sub
Sub DirtyStatus_Throw()
  Project_Is_Dirty = True
  Call DirtyStatus_Set_Current
End Sub


Sub frmMain_Close_All_Windows()
Dim ifc%
Dim i%
  On Error Resume Next
  ifc% = Forms.Count - 1
  For i% = ifc% To 0 Step -1
    'If (Forms(i%).name <> "frmMain") And _
       (Forms(i%).name <> "frmProgress") Then
    If (Forms(i%).Name <> "frmMain") Then
      Unload Forms(i%)
    End If
  Next i%
End Sub


Sub CenterOnScreen(frm_to_center As Form)
  frm_to_center.Left = (Screen.Width - frm_to_center.Width) / 2
  frm_to_center.Top = (Screen.Height - frm_to_center.Height) / 2
End Sub
Sub CenterOnForm(frm_to_center As Form, Frm As Form)
  frm_to_center.Left = Frm.Left + (Frm.Width - frm_to_center.Width) / 2
  frm_to_center.Top = Frm.Top + (Frm.Height - frm_to_center.Height) / 2
End Sub


Sub Show_Message00(msg As String, flags As Integer, WinTitle As String)
  MsgBox msg, flags, WinTitle
End Sub
Sub Show_Message0(msg As String, flags As Integer)
  Call Show_Message00(msg, vbInformation, AppName_For_Display_Short)
End Sub
Sub Show_Message(msg As String)
  Call Show_Message0(msg, vbInformation)
End Sub
Sub Show_Error(msg As String)
  Beep
  Call Show_Message0(msg, vbExclamation)
End Sub
Sub Show_Trapped_Error(subname As String)
  Call Show_Error("An error #" & Trim$(Str$(Err)) & _
      " has occurred in routine " & Trim$(subname) & _
      ": `" & Trim$(Error$) & "`.  Ending this operation.")
End Sub


Sub Launch_Notepad(fn_edit As String)
Dim CmdLine As String
Dim RetVal As Integer
  CmdLine = "notepad " & fn_edit
  RetVal = 0 * Shell(CmdLine, 3)
End Sub

