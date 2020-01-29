VERSION 5.00
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmDyeStudy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing Dye Study Data"
   ClientHeight    =   5925
   ClientLeft      =   1560
   ClientTop       =   3360
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Rows"
      Height          =   315
      Left            =   5280
      TabIndex        =   12
      Top             =   4350
      Width           =   1605
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert Rows"
      Height          =   315
      Left            =   5280
      TabIndex        =   11
      Top             =   4050
      Width           =   1605
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste from Excel"
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Top             =   3735
      Width           =   1605
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy to Excel"
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      Top             =   3435
      Width           =   1605
   End
   Begin VB.CommandButton cmdDisplayResultsDisp 
      Caption         =   "D&isplay Results for Dispersion Model"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4980
      TabIndex        =   8
      Top             =   1890
      Width           =   2385
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset to Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      Top             =   3120
      Width           =   1605
   End
   Begin VB.CommandButton cmdDisplayResults 
      Caption         =   "    &Display Results for Tanks In Series Model"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4980
      TabIndex        =   5
      Top             =   1245
      Width           =   2385
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   5790
      TabIndex        =   3
      Text            =   "txtData(0)"
      Top             =   2670
      Width           =   1635
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Ca&lculate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      Top             =   825
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Top             =   180
      Width           =   1065
   End
   Begin VCIF1Lib.F1Book f1book_dyestudy 
      Height          =   5655
      Left            =   240
      OleObjectBlob   =   "DyeStudy.frx":0000
      TabIndex        =   0
      Top             =   90
      Width           =   4365
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"DyeStudy.frx":0567
      Height          =   1065
      Left            =   4800
      TabIndex        =   7
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label lblDesc 
      Caption         =   "Last Calculated:"
      Height          =   465
      Index           =   0
      Left            =   4860
      TabIndex        =   4
      Top             =   2610
      Width           =   825
   End
   Begin VB.Menu mnuItem 
      Caption         =   "Edit"
      Index           =   1
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy to Excel"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste from Excel"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert Rows"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Rows"
      End
   End
End
Attribute VB_Name = "frmDyeStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Integer
Dim TempProj As Project_Type


Const frmDyeStudy_declarations_end = 0

  
'RETURNS:
'  TRUE = USER HIT OK
'  FALSE = USER HIT CANCEL
Public Function frmDyeStudy_DoEdit() As Integer
Dim is_aborted As Integer
Dim name_new As String
  
  'IMPORT THIS PROJECT FROM MEMORY TO THE FORM.
  TempProj = nowproj
  
  'SHOW THE FORM.
  frmDyeStudy.Show 1
  
  'UPDATE MEMORY.
  If TempProj.dirty Then
    nowproj = TempProj
  End If
  
  'RETURN TO MAIN WINDOW.
  frmDyeStudy_DoEdit = nowproj.dirty

End Function


Private Sub cmdCalculate_Click()
Dim calcdate As String
Dim fn_this As String
Dim out_file As String
Dim count_next As Boolean
Dim textline1 As String
Dim textline2 As String
Dim i As Integer
Dim Got_Count As Boolean
Dim f As String
Dim Not_Yet As Boolean
Dim StrTextLine As String
Dim strTheta As String
Dim strE As String


If TempProj.DyeStudy(1).time = " " Then
  Call Show_Error("No data has been entered into the grid")
Else
  Call Pec_Run
  out_file = App.Path & "\exes\outpt2.txt"
  count_next = False
  Got_Count = False
  Not_Yet = False
  TempProj.Predicted_count = 0
  If FileExists(out_file) Then
    f = FreeFile
    Open out_file For Input As #f
      Do While Not EOF(f)
      Input #f, textline1
      Select Case textline1
        Case textline1
          If (InStr(1, textline1, "----") > 0) Then
           count_next = True
          End If
      
          If (InStr(1, textline1, "----") = 0) And count_next = True Then
            If TempProj.Predicted_count = 0 Then
              TempProj.Predicted_count = Val(textline1)
              ReDim TempProj.Predicted(1 To TempProj.Predicted_count)
              i = 0
            Else
              StrTextLine = Trim$(Parser_RemoveDuplicateSeparators(" ", textline1))
              If (2 <> Parser_GetNumArgs(" ", StrTextLine)) Then
                Call Show_Message("Output file is corrupt, please call Vendor", _
                  vbExclamation, App.title)
              Else
                TempProj.Predicted_Available = True
                Call Parser_GetArg(" ", StrTextLine, 1, strTheta)
                Call Parser_GetArg(" ", StrTextLine, 2, strE)
                  i = i + 1
                  TempProj.Predicted(i).Predicted_Theta = Val(strTheta)
                  TempProj.Predicted(i).Predicted_E = Val(strE)
              End If
            End If
         End If
      End Select
      
    Loop
    Close #f
  End If
  
  out_file = App.Path & "\exes\outpt3.txt"
  count_next = False
  Got_Count = False
  Not_Yet = False
  TempProj.Experimental_count = 0
  If FileExists(out_file) Then
    f = FreeFile
    Open out_file For Input As #f
      Do While Not EOF(f)
      Input #f, textline1
      Select Case textline1
        Case textline1
          If (InStr(1, textline1, "----") > 0) Then
           count_next = True
          End If
      
          If (InStr(1, textline1, "----") = 0) And count_next = True Then
            If TempProj.Experimental_count = 0 Then
              TempProj.Experimental_count = Val(textline1)
              ReDim TempProj.Experimental(1 To TempProj.Experimental_count)
              i = 0
            Else
              StrTextLine = Trim$(Parser_RemoveDuplicateSeparators(" ", textline1))
              If (2 <> Parser_GetNumArgs(" ", StrTextLine)) Then
                Call Show_Message("Output file is corrupt, please call Vendor", _
                  vbExclamation, App.title)
              Else
                TempProj.Predicted_Available = True
                Call Parser_GetArg(" ", StrTextLine, 1, strTheta)
                Call Parser_GetArg(" ", StrTextLine, 2, strE)
                  i = i + 1
                  TempProj.Experimental(i).Experimental_Theta = Val(strTheta)
                  TempProj.Experimental(i).Experimental_E = Val(strE)
              End If
            End If
         End If
      End Select
      
    Loop
    Close #f
    
    'run dispersion model only if outpt3.txt exists
    Call PecDisp_Run
    out_file = App.Path & "\exes\outpt4.txt"
    count_next = False
    Got_Count = False
    Not_Yet = False
    TempProj.PredictedDispClosed_count = 0
    If FileExists(out_file) Then
      f = FreeFile
      Open out_file For Input As #f
        Do While Not EOF(f)
        Input #f, textline1
        Select Case textline1
          Case textline1
            If (InStr(1, textline1, "----") > 0) Then
             count_next = True
            End If
        
            If (InStr(1, textline1, "----") = 0) And count_next = True Then
              If TempProj.PredictedDispClosed_count = 0 Then
                TempProj.PredictedDispClosed_count = Val(textline1)
                ReDim TempProj.DispClosed(1 To TempProj.PredictedDispClosed_count)
                i = 0
              Else
                StrTextLine = Trim$(Parser_RemoveDuplicateSeparators(" ", textline1))
                If (2 <> Parser_GetNumArgs(" ", StrTextLine)) Then
                  Call Show_Message("Output file is corrupt, please call Vendor", _
                    vbExclamation, App.title)
                Else
                  Call Parser_GetArg(" ", StrTextLine, 1, strTheta)
                  Call Parser_GetArg(" ", StrTextLine, 2, strE)
                    i = i + 1
                    TempProj.DispClosed(i).PredictedDispClosed_Theta = Val(strTheta)
                    TempProj.DispClosed(i).PredictedDispClosed_E = Val(strE)
                End If
              End If
           End If
        End Select
        
      Loop
      Close #f
    End If
    
    out_file = App.Path & "\exes\outpt5.txt"
    count_next = False
    Got_Count = False
    Not_Yet = False
    TempProj.PredictedDispOpen_count = 0
    If FileExists(out_file) Then
      f = FreeFile
      Open out_file For Input As #f
        Do While Not EOF(f)
        Input #f, textline1
        Select Case textline1
          Case textline1
            If (InStr(1, textline1, "----") > 0) Then
             count_next = True
            End If
        
            If (InStr(1, textline1, "----") = 0) And count_next = True Then
              If TempProj.PredictedDispOpen_count = 0 Then
                TempProj.PredictedDispOpen_count = Val(textline1)
                ReDim TempProj.DispOpen(1 To TempProj.PredictedDispOpen_count)
                i = 0
              Else
                StrTextLine = Trim$(Parser_RemoveDuplicateSeparators(" ", textline1))
                If (2 <> Parser_GetNumArgs(" ", StrTextLine)) Then
                  Call Show_Message("Output file is corrupt, please call Vendor", _
                    vbExclamation, App.title)
                Else
                  Call Parser_GetArg(" ", StrTextLine, 1, strTheta)
                  Call Parser_GetArg(" ", StrTextLine, 2, strE)
                    i = i + 1
                    TempProj.DispOpen(i).PredictedDispOpen_Theta = Val(strTheta)
                    TempProj.DispOpen(i).PredictedDispOpen_E = Val(strE)
                End If
              End If
           End If
        End Select
        
      Loop
      Close #f
    
    End If
    
  End If
  
  
  TempProj.dyestudy_calcdate = Now
  txtData(0) = TempProj.dyestudy_calcdate
  IsCalculated = True
  Call DirtyFlag_Throw(TempProj)
End If

End Sub

Sub Pec_Run()
Dim fn_FortranModuleEXE As String
Dim fpath_run As String
Dim fpath_save As String
Dim success As Integer

Dim calctime_start As String
Dim calctime_end As String
Dim msg As String
Dim elapsed_min As Double
Dim CmdLine As String
Dim f As String
Dim i As Integer
Dim output_text As String
Dim last_line As Integer

 
  'WRITE INPUT FILES FOR FORTRAN MODULE.
  FortranLink_fn_MainInput = App.Path & "\exes\input.txt"
  FortranLink_fn_MainOutput = App.Path & "\exes\outpt.txt"
  FortranLink_fn_MainOutput2 = App.Path & "\exes\outpt2.txt"
  FortranLink_fn_MainOutput3 = App.Path & "\exes\outpt3.txt"
  
  Call PecLink_WriteInputFile(FortranLink_fn_MainInput)
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput)
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput2)
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput3)
  
  'CALL FORTRAN MODULE.
  Call ChangeDir_Exes
  fn_FortranModuleEXE = App.Path & "\exes\pec.exe"
  CmdLine = fn_FortranModuleEXE
  calctime_start = Now
  Call FortranLink_ExecAndWaitForProcess(CmdLine)
  calctime_end = Now
  Call ChangeDir_Main
     
  'DID IT SUCCEED?
  success = (Dir(FortranLink_fn_MainOutput3) <> "")
  If (success) Then
    elapsed_min = DateDiff("s", calctime_start, calctime_end) / 60#
    msg = "Calculations succeeded." & vbCrLf & _
        vbCrLf & _
        "    Calculations began at " & calctime_start & vbCrLf & _
        "    Calculations ended at " & calctime_end & vbCrLf & _
        vbCrLf & _
        "    Total elapsed time = " & qstr(elapsed_min) & " minutes"
        
    f = FreeFile
    Open FortranLink_fn_MainOutput For Input As #f
    TempProj.dyestudy_output = ""
    
    Do While Not EOF(f)
      Line Input #f, output_text
        TempProj.dyestudy_output = TempProj.dyestudy_output _
          + vbCrLf + output_text
    Loop
    Close #f
  Else
    msg = "Calculations failed."
    TempProj.Predicted_Available = False
  End If
  Call Show_Message(msg, vbExclamation, "Tanks in Series")
  
End Sub


Sub PecDisp_Run()
Dim fn_FortranModuleEXE As String
Dim fpath_run As String
Dim fpath_save As String
Dim success As Integer

Dim calctime_start As String
Dim calctime_end As String
Dim msg As String
Dim elapsed_min As Double
Dim CmdLine As String
Dim f As String
Dim i As Integer
Dim output_text As String
Dim last_line As Integer

 
  'WRITE INPUT FILES FOR FORTRAN MODULE.
  FortranLink_fn_MainInput = App.Path & "\exes\pecinput.txt"
  FortranLink_fn_MainOutput = App.Path & "\exes\pecoutpt.txt"
  FortranLink_fn_MainOutput2 = App.Path & "\exes\outpt4.txt"
  FortranLink_fn_MainOutput3 = App.Path & "\exes\outpt5.txt"
  
  Call PecLink_WriteInputFile(FortranLink_fn_MainInput)
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput)
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput2)
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput3)
  
  'CALL FORTRAN MODULE.
  Call ChangeDir_Exes
  fn_FortranModuleEXE = App.Path & "\exes\pecdisp.exe"
  CmdLine = fn_FortranModuleEXE
  calctime_start = Now
  Call FortranLink_ExecAndWaitForProcess(CmdLine)
  calctime_end = Now
  Call ChangeDir_Main
     
  'DID IT SUCCEED?
  success = (Dir(FortranLink_fn_MainOutput3) <> "")
  If (success) Then
    elapsed_min = DateDiff("s", calctime_start, calctime_end) / 60#
    msg = "Calculations succeeded." & vbCrLf & _
        vbCrLf & _
        "    Calculations began at " & calctime_start & vbCrLf & _
        "    Calculations ended at " & calctime_end & vbCrLf & _
        vbCrLf & _
        "    Total elapsed time = " & qstr(elapsed_min) & " minutes"
        
    f = FreeFile
    Open FortranLink_fn_MainOutput For Input As #f
    TempProj.dyestudydisp_output = ""
    
    Do While Not EOF(f)
      Line Input #f, output_text
        TempProj.dyestudydisp_output = TempProj.dyestudydisp_output _
          + vbCrLf + output_text
    Loop
    Close #f
  Else
    msg = "Calculations failed."
    TempProj.Predicted_Available = False
  End If
  Call Show_Message(msg, vbExclamation, "Dispersion")
  

  
End Sub


Sub PecLink_WriteInputFile(fn_InputFile As String)

Dim f As Integer
Dim i As Integer
 
  'WRITE THE FORTRAN INPUT FILE.
  f = FreeFile
  Open fn_InputFile For Output As #f
  '---- MAIN INPUTS --------------------------------------------------------------------------------------------
  Call WriteFortranInput(f, TempProj.dyestudy_count - 1, "")
  For i = 1 To nowproj.dyestudy_count
    Call WriteFortranInput(f, TempProj.DyeStudy(i).time, " ")
    Call WriteFortranInput(f, TempProj.DyeStudy(i).concentration, " ")
  Next i
  
  Close #f
  

End Sub



Private Sub cmdDelete_Click()
  Dim n1R As Integer
  Dim n2R As Integer
  Dim n1C As Integer
  Dim n2C As Integer
  
  n1R = frmDyeStudy.f1book_dyestudy.SelStartRow
  n2R = frmDyeStudy.f1book_dyestudy.SelEndRow
  n1C = frmDyeStudy.f1book_dyestudy.SelStartCol
  n2C = frmDyeStudy.f1book_dyestudy.SelEndCol
  
  frmDyeStudy.f1book_dyestudy.DeleteRange n1R, n1C, n2R, n2C, F1ShiftRows
  IsCalculated = False
  Call DirtyFlag_Throw(TempProj)
  End Sub

Private Sub cmdDisplayResults_Click()
Dim fn_this As String

  'see if data changed and not calculated
  If Not IsCalculated Then
    Call Show_Message("Data has changed, please calculate first.", _
    vbExclamation, App.title)
  Else
    'look for output.txt and if not there,display message
    fn_this = App.Path & "\exes\outpt.txt"
    If (FileExists(fn_this) = False) Then
      Call Show_Message("No output file was found, please calculate.", _
      vbExclamation, App.title)
    Else
      Call Launch_Notepad(App.Path & "\exes\outpt.txt")
    End If
  End If
End Sub

Private Sub cmdDisplayResultsDisp_Click()
  Dim fn_this As String
  'see if data changed and not calculated
  If Not IsCalculated Then
    Call Show_Message("Data was changed, please calculate first.", _
    vbExclamation, App.title)
  Else
    'look for pecoutput.txt and if not there,display message
    fn_this = App.Path & "\exes\pecoutpt.txt"
    If (FileExists(fn_this) = False) Then
      Call Show_Message("No output file was found, please calculate.", _
      vbExclamation, App.title)
    Else
      Call Launch_Notepad(App.Path & "\exes\pecoutpt.txt")
    End If
  End If
End Sub

Private Sub cmdInsert_Click()
  Dim n1R As Integer
  Dim n2R As Integer
  Dim n1C As Integer
  Dim n2C As Integer
  
  n1R = frmDyeStudy.f1book_dyestudy.SelStartRow
  n2R = frmDyeStudy.f1book_dyestudy.SelEndRow
  n1C = frmDyeStudy.f1book_dyestudy.SelStartCol
  n2C = frmDyeStudy.f1book_dyestudy.SelEndCol
  
  frmDyeStudy.f1book_dyestudy.InsertRange n1R, n1C, n2R, n2C, F1ShiftRows
  IsCalculated = False
  Call DirtyFlag_Throw(TempProj)
End Sub

Private Sub cmdOK_Click()
  nowproj = TempProj
  Unload Me
End Sub



Private Sub cmdPaste_Click()
Dim i As Integer
Dim temp_count As Integer
Dim oldVal As String
Dim newVal As String


  On Error GoTo ErrReset
  Call cmdReset_Click
  
  frmDyeStudy.f1book_dyestudy.EditPasteValues
  frmDyeStudy.f1book_dyestudy.SetFocus
  
  For i = 1 To 1600
    f1book_dyestudy.Row = i
    f1book_dyestudy.Col = 1
    newVal = Me.f1book_dyestudy.Entry
    TempProj.DyeStudy(i).time = newVal
    f1book_dyestudy.Col = 2
    newVal = Me.f1book_dyestudy.Entry
    TempProj.DyeStudy(i).concentration = newVal
    If newVal = " " Then
      TempProj.dyestudy_count = i
      Exit For
    End If
  Next i
  
  Call DirtyFlag_Throw(TempProj)
  Call refresh_frmDyeStudy(TempProj)

  Exit Sub
  
ErrReset:
  Call Show_Error("Data in clipboard is not valid.")
  Call cmdReset_Click
  Exit Sub
  
End Sub

Private Sub cmdReset_Click()
  Dim i As Integer
  
  ReDim TempProj.DyeStudy(1 To 1600)
  TempProj.dyestudy_count = 1600
  For i = 1 To 1600
    f1book_dyestudy.EntryRC(i, 1) = " "
    f1book_dyestudy.EntryRC(i, 2) = " "
  Next i
  For i = 1 To TempProj.dyestudy_count
      TempProj.DyeStudy(i).time = " "
      TempProj.DyeStudy(i).concentration = " "
    Next i
  txtData(0) = ""
  
  Call DirtyFlag_Throw(TempProj)
  Call refresh_frmDyeStudy(TempProj)
  
End Sub

Private Sub cmdCopy_Click()
    frmDyeStudy.f1book_dyestudy.EditCopy
End Sub

Private Sub f1book_dyestudy_EndEdit(EditString As String, Cancel As Integer)
Dim idx As Integer
Dim newVal As String
Dim oldVal As String

  idx = f1book_dyestudy.Row
  If (idx < 1) Then Exit Sub

  Select Case f1book_dyestudy.Col
    Case 1:
      If (Not IsValidNumber0(EditString, vbDouble)) Then
        EditString = " "
        Call Show_Error("Invalid number. ")
      End If
      oldVal = TempProj.DyeStudy(idx).time
      newVal = CDbl(Val(EditString))
      TempProj.DyeStudy(idx).time = newVal
      If oldVal = " " Then
        TempProj.dyestudy_count = idx + 1
      End If
      
      'REFRESH DyeStudy window, ESPECIALLY THE GRIDS.
      Call refresh_frmDyeStudy(TempProj)
    Case 2:
      newVal = CDbl(Val(EditString))
      If (Not IsValidNumber0(EditString, vbDouble)) Then
        EditString = ""
        Call Show_Error("Invalid number. ")
      End If
      TempProj.DyeStudy(idx).concentration = newVal
      
      'REFRESH DyeStudy window, ESPECIALLY THE GRIDS.
      Call refresh_frmDyeStudy(TempProj)
      
  End Select
  IsCalculated = False
  Call DirtyFlag_Throw(TempProj)
End Sub



Private Sub Form_Load()

  Call CenterOnForm(Me, frmMain)
  IsCalculated = True
  If UCase(get_program_releasetype()) = "BETA" Then
    cmdCalculate.Enabled = False
  Else
    cmdCalculate.Enabled = True
  End If
  Call refresh_frmDyeStudy(TempProj)
  
End Sub



Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub

Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Sub Populate_frmDyeStudy_Units()
  Call unitsys_register(frmDyeStudy, lblDesc(0), txtData(0), Nothing, "", _
      "", "", "", "", 0#, False)
End Sub


Private Sub mnuCopy_Click()
   frmDyeStudy.f1book_dyestudy.EditCopy

End Sub


Private Sub mnuDelete_Click()
  Dim n1R As Integer
  Dim n2R As Integer
  Dim n1C As Integer
  Dim n2C As Integer
  
  n1R = frmDyeStudy.f1book_dyestudy.SelStartRow
  n2R = frmDyeStudy.f1book_dyestudy.SelEndRow
  n1C = frmDyeStudy.f1book_dyestudy.SelStartCol
  n2C = frmDyeStudy.f1book_dyestudy.SelEndCol
  
  frmDyeStudy.f1book_dyestudy.DeleteRange n1R, n1C, n2R, n2C, F1ShiftRows
  IsCalculated = False
  Call DirtyFlag_Throw(TempProj)
  
End Sub

Private Sub mnuInsert_Click()
  Dim n1R As Integer
  Dim n2R As Integer
  Dim n1C As Integer
  Dim n2C As Integer
  
  n1R = frmDyeStudy.f1book_dyestudy.SelStartRow
  n2R = frmDyeStudy.f1book_dyestudy.SelEndRow
  n1C = frmDyeStudy.f1book_dyestudy.SelStartCol
  n2C = frmDyeStudy.f1book_dyestudy.SelEndCol
  
  frmDyeStudy.f1book_dyestudy.InsertRange n1R, n1C, n2R, n2C, F1ShiftRows
  IsCalculated = False
  Call DirtyFlag_Throw(TempProj)
  
End Sub

Private Sub mnuPaste_Click()
Dim i As Integer
Dim temp_count As Integer
Dim oldVal As String
Dim newVal As String


  On Error GoTo ErrReset
  Call cmdReset_Click
  
  frmDyeStudy.f1book_dyestudy.EditPasteValues
  frmDyeStudy.f1book_dyestudy.SetFocus
  
  For i = 1 To 1600
    f1book_dyestudy.Row = i
    f1book_dyestudy.Col = 1
    newVal = Me.f1book_dyestudy.Entry
    TempProj.DyeStudy(i).time = newVal
    f1book_dyestudy.Col = 2
    newVal = Me.f1book_dyestudy.Entry
    TempProj.DyeStudy(i).concentration = newVal
    If newVal = " " Then
      TempProj.dyestudy_count = i
      Exit For
    End If
  Next i
  
  Call DirtyFlag_Throw(TempProj)
  Call refresh_frmDyeStudy(TempProj)

  Exit Sub
  
ErrReset:
  Call Show_Error("Data in clipboard is not valid.")
  Call cmdReset_Click
  Exit Sub
  
End Sub
