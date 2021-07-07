Attribute VB_Name = "FileOptMod"
'This module will handle File Menu Options such
'as Print, Save, Open, etc.

Global FileName As String
Global OldFileName As String
Global Const FILEID_STEPP = "StEPP Contaminant File"   'File Identifier for a StEPP Contaminant Design File
Global JustLoadedFile As Integer   'Tells if a StEPP Design File has just been loaded for use in displaying information with cboSelectContaminant Click Event
Global Stepp_ClientProgram  As String

Sub LoadFileStEPP(FileName As String)
Dim Ctl As Control
Set Ctl = contam_prop_form.CommonDialog1

    On Error Resume Next
'    contam_prop_form!CMDialog1.DefaultExt = "stp"
'    contam_prop_form!CMDialog1.Filter = "StEPP Files (*.stp)|*.stp"
'    contam_prop_form!CMDialog1.DialogTitle = "Load StEPP Contaminant File"
'    contam_prop_form!CMDialog1.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
'    contam_prop_form!CMDialog1.CancelError = True
'    contam_prop_form!CMDialog1.Action = 1
'    FileName$ = contam_prop_form!CMDialog1.FileName
    Ctl.DefaultExt = "stp"
    Ctl.Filter = "StEPP Files (*.stp)|*.stp"
    Ctl.DialogTitle = "Load StEPP Contaminant File"
    Ctl.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    Ctl.CancelError = True
    Ctl.Action = 1
    FileName$ = Ctl.FileName
    If Err = 32755 Then   'Cancel selected by user
       FileName$ = ""
    End If

End Sub

Sub LoadStEPPDesign(OverrideFilename As String)
    Dim FileID As String, msg As String
    Dim i As Integer
    Dim NamePlusCAS As String

    contam_prop_form!cmdSelectContaminant.SetFocus

    If (OverrideFilename <> "") Then
      FileName$ = OverrideFilename
    Else
      Call LoadFileStEPP(FileName)
    End If
    
    If FileName$ <> "" Then
       FileID = ""
       Open FileName$ For Random As #1 Len = Len(phprop)
       On Error Resume Next
       Get #1, 1, FileID
       If FileID <> FILEID_STEPP Then
          msg = "Invalid StEPP Design File"
          MsgBox msg, 48, "Error"
          Close #1
          Exit Sub
       End If
       
       Get #1, 2, NumSelectedChemicals

       contam_prop_form!cboSelectContaminant.Clear
       contam_prop_form!cboSelectContaminant.Enabled = True

       For i = 1 To NumSelectedChemicals
           Get #1, i + 2, PropContaminant(i)
           If (SteppLink_SpecifiedPressure <> "") Then
             PropContaminant(i).OperatingPressure = CDbl(SteppLink_SpecifiedPressure)
           End If
           If (SteppLink_SpecifiedTemperature <> "") Then
             PropContaminant(i).OperatingTemperature = CDbl(SteppLink_SpecifiedTemperature)
           End If
       Next i
       
       For i = 1 To NumSelectedChemicals
           NamePlusCAS = " " & Trim$(Str$(PropContaminant(i).CasNumber)) & "  " & Trim$(PropContaminant(i).Name)
           contam_prop_form!cboSelectContaminant.AddItem NamePlusCAS
       Next i

       PreviouslySelectedIndex = -1
       JustLoadedFile = True
       phprop = PropContaminant(1)

       contam_prop_form!cboSelectContaminant.ListIndex = 0
       contam_prop_form!mnuFile(4).Enabled = True
       contam_prop_form!mnuFile(5).Enabled = True
       contam_prop_form!mnuFile(7).Enabled = True
       contam_prop_form!cmdUnselectContaminant.Enabled = True
'*** Modification v1019 by David R. Hokanson (16may2000)
'       Call frmmain.frmMain_Reset_DemoVersionDisablings
       Call contam_prop_form.frmMain_Reset_DemoVersionDisablings
'*** End Modification v1019 by David R. Hokanson (16may2000)
       

       Close #1
              
       'Add this file to the last-few-files list.
       Call LastFewFiles_MoveFilenameToTop(FileName$)
    
    End If

End Sub

Sub SaveFileStEPP(FileName As String)
Dim Ctl As Control
Set Ctl = contam_prop_form.CommonDialog1

    On Error Resume Next
'    contam_prop_form!CMDialog1.DefaultExt = "stp"
'    contam_prop_form!CMDialog1.Filter = "StEPP Files (*.stp)|*.stp"
'    contam_prop_form!CMDialog1.DialogTitle = "Save StEPP Contaminant File"
'    contam_prop_form!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
'    contam_prop_form!CMDialog1.CancelError = True
'    contam_prop_form!CMDialog1.Action = 2
'    FileName$ = contam_prop_form!CMDialog1.FileName
    Ctl.DefaultExt = "stp"
    Ctl.Filter = "StEPP Files (*.stp)|*.stp"
    Ctl.DialogTitle = "Save StEPP Contaminant File"
    Ctl.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    Ctl.CancelError = True
    Ctl.Action = 2
    FileName$ = Ctl.FileName
    If Err = 32755 Then   'Cancel selected by user
       FileName$ = ""
    End If

End Sub

Sub SaveStEPPDesign()
    Dim i As Integer
    Dim FileID As String
  
  If (IsThisADemo() = True) Then
    Call Demo_ShowError("Saving is not allowed in the demonstration version.")
    Exit Sub
  End If

    If FileName$ = "" Then
       Call SaveFileStEPP(FileName)
    End If

    If FileName$ <> "" Then
       FileID = FILEID_STEPP


       PropContaminant(PreviouslySelectedIndex) = phprop

       Open FileName$ For Random As #1 Len = Len(phprop)
       
       Put #1, 1, FileID
       Put #1, 2, NumSelectedChemicals
      
       For i = 1 To NumSelectedChemicals
           Put #1, i + 2, PropContaminant(i)
       Next i

       Close #1
       contam_prop_form!mnuFile(4).Enabled = True
       
'*** Modification v1019 by David R. Hokanson (16may2000)
'       Call frmmain.frmMain_Reset_DemoVersionDisablings
       Call contam_prop_form.frmMain_Reset_DemoVersionDisablings
'*** End Modification v1019 by David R. Hokanson (16may2000)

       'Add this file to the last-few-files list.
       Call LastFewFiles_MoveFilenameToTop(FileName$)

    Else   'Cancel Selected by user
       FileName$ = OldFileName$
    End If
End Sub

'Sub XOLD_ini_initializethisprogram(SpecifiedProgramType As String)
'Dim infotype As String
'Dim defaultresult As String
'Dim returnvalue As String * 125
'Dim storage As String * 144
'Dim newdata As String
'Dim defaultnumber As Long, valid As Long, string_size As Long
'Dim f As Integer, i As Integer, Value As Integer
'Dim temp As String
'
'  'Set global variable to specified program type
'  INI_ProgramType = Trim$(SpecifiedProgramType)
'
'  'get windows directory and look for adsim.ini
'  Value = GetWindowsDirectory(ByVal storage, ByVal Len(storage))
'  INI_WindowsDir = Trim$(Left$(storage, Value))
'
'  'if does not exist, create ini file
'  'INI_FileName = INI_WindowsDir & "\" & INI_ProgramType & ".ini"
'  INI_FileName = Global_fpath_dir_CPAS & "\DBASE\" & INI_ProgramType & ".ini"
'
'
'  If (Dir(INI_FileName)) = "" Then
'    '======>  Program Specified Initialization!  <======
'    f = FreeFile
'    Open INI_FileName For Output As f
'    Print #f, "[stepp]"
'    Print #f, "app_path="
'    Print #f, "has_seen_disclaimer=0"
'    Print #f, "has_seen_steppinfo=0"
'    Print #f, "oldfile1="
'    Print #f, "oldfile2="
'    Print #f, "oldfile3="
'    Print #f, "oldfile4="
'
'    Close #f
'  End If
'
'  'temp = ini_getsetting(INI_FileName, INI_ProgramType, "app_path")
'  temp = INI_Getsetting("app_path")
'
'  'if incorrect path set with programs current path being used now
'  If ((StrComp(temp, App.Path) <> 0) Or (temp = "")) Then
'    Call INI_PutSetting("app_path", Trim$(App.Path))
'  End If
'
'  ChDir App.Path
'
'End Sub

