VERSION 5.00
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmDyeStudy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing Dye Study Data"
   ClientHeight    =   5925
   ClientLeft      =   2625
   ClientTop       =   2385
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste"
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
      Left            =   5700
      TabIndex        =   7
      Top             =   3420
      Width           =   1065
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
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
      Left            =   5700
      TabIndex        =   6
      Top             =   3030
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Display Results"
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
      Left            =   5220
      TabIndex        =   5
      Top             =   1980
      Width           =   1815
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   5790
      TabIndex        =   3
      Text            =   "txtData(0)"
      Top             =   2490
      Width           =   1635
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Ca&lculate"
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
      Left            =   5340
      TabIndex        =   2
      Top             =   1620
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
      Left            =   5700
      TabIndex        =   1
      Top             =   510
      Width           =   1065
   End
   Begin VCIF1Lib.F1Book f1book_dyestudy 
      Height          =   5655
      Left            =   210
      OleObjectBlob   =   "DyeStudy.frx":0000
      TabIndex        =   0
      Top             =   90
      Width           =   4365
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"DyeStudy.frx":0567
      Height          =   1185
      Left            =   4890
      TabIndex        =   8
      Top             =   4500
      Width           =   2535
   End
   Begin VB.Label lblDesc 
      Caption         =   "Last Calculated:"
      Height          =   465
      Index           =   0
      Left            =   4860
      TabIndex        =   4
      Top             =   2430
      Width           =   825
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
  TempProj = NowProj
  
  'SHOW THE FORM.
  frmDyeStudy.Show 1
  
  'UPDATE MEMORY.
  If (Not USER_HIT_CANCEL) Then
    NowProj = TempProj
  End If
  
  'RETURN TO MAIN WINDOW.
  frmDyeStudy_DoEdit = Not USER_HIT_CANCEL

End Function


Private Sub cmdCalculate_Click()
Dim calcdate As String
Dim fn_this As String


'modify FortranLink_Run to take parameters for output file name,
'input file name, and fortran executable name or copy code
'to new procedure and modify?

'(1.) If it exists, delete the "output.txt" file from the EXES subdirectory
'(2.) Using the user-entered data, generate the "input.txt" file in the EXES
'subdirectory
'(3.) Call the PEC.EXE program in the EXES subdirectory; use code similar to
'that found in the FortranLink_Run() subroutine
'(4.) Check for the existence of "output.txt" in EXES; if it does not exist,
'show an error message
'(4a.) Populate date/time field with calculation date/time
'(5.) Display the "output.txt" file using Launch_Notepad()

If TempProj.DyeStudy(1).time = -1 Then
  Call Show_Error("No data has been entered into the grid")
Else
  Call Pec_Run
  TempProj.dyestudy_calcdate = Now
  txtData(0) = TempProj.dyestudy_calcdate
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
'  FortranLink_fn_MainOutput = App.Path & "\exes\output.txt"
  FortranLink_fn_MainOutput = App.Path & "\exes\outpt.txt"

  Call PecLink_WriteInputFile(FortranLink_fn_MainInput)
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput)
  
  'CALL FORTRAN MODULE.
  Call ChangeDir_Exes
  fn_FortranModuleEXE = App.Path & "\exes\pec.exe"
  CmdLine = fn_FortranModuleEXE
  calctime_start = Now
  Call FortranLink_ExecAndWaitForProcess(CmdLine)
  calctime_end = Now
  Call ChangeDir_Main
     
  'DID IT SUCCEED?
  success = (Dir(FortranLink_fn_MainOutput) <> "")
  If (success) Then
    elapsed_min = DateDiff("s", calctime_start, calctime_end) / 60#
    msg = "Calculations succeeded." & vbCrLf & _
        vbCrLf & _
        "    Calculations began at " & calctime_start & vbCrLf & _
        "    Calculations ended at " & calctime_end & vbCrLf & _
        vbCrLf & _
        "    Total elapsed time = " & qstr(elapsed_min) & " minutes"
    
    ' put output.txt into a comma delimited string to be saved in Access table
    f = FreeFile
    Open FortranLink_fn_MainOutput For Input As #f
    TempProj.dyestudy_output = ""
    Do While Not EOF(f)
      Line Input #f, output_text
      ' this puts an extra comma at the beginning, but the code in
      ' File_Save_Latest... to put it in the output.txt file takes it out
        TempProj.dyestudy_output = TempProj.dyestudy_output _
          + ", " + output_text
    Loop
    Close #f
  Else
    msg = "Calculations failed."
  End If
  Call Show_Error(msg) 'show_error

  
End Sub


Sub PecLink_WriteInputFile(fn_InputFile As String)

Dim f As Integer
Dim i As Integer
 
  'WRITE THE FORTRAN INPUT FILE.
  f = FreeFile
  Open fn_InputFile For Output As #f
  '---- MAIN INPUTS --------------------------------------------------------------------------------------------
  Call WriteFortranInput(f, TempProj.dyestudy_count - 1, "")
  For i = 1 To NowProj.dyestudy_count
    Call WriteFortranInput(f, TempProj.DyeStudy(i).time, "")
    Call WriteFortranInput(f, TempProj.DyeStudy(i).concentration, "")
  Next i
  
  Close #f
  

End Sub


Private Sub cmdOK_Click()
  USER_HIT_CANCEL = False
  If (Not USER_HIT_CANCEL) Then
    NowProj = TempProj
  End If
  Unload Me
End Sub




Private Sub Command1_Click()
Dim fn_this As String
  'look for output.txt and if not there,display message
'  fn_this = MAIN_APP_PATH & "\exes\output.txt"
  fn_this = MAIN_APP_PATH & "\exes\outpt.txt"
  If (FileExists(fn_this) = False) Then
    Call Show_Message("No output file was found, please calculate.", _
    vbExclamation, App.title)
  Else
'    Call Launch_Notepad(App.Path & "\exes\output.txt")
    Call Launch_Notepad(App.Path & "\exes\outpt.txt")
  End If
End Sub

Private Sub cmdPaste_Click()
Dim i As Integer
Dim temp_count As Integer
Dim oldVal As String
Dim newVal As String

  Call cmdReset_Click
  
  frmDyeStudy.f1book_dyestudy.EditPaste
  frmDyeStudy.f1book_dyestudy.SetFocus
  
  For i = 1 To 400
    f1book_dyestudy.Row = i
    f1book_dyestudy.Col = 1
    On Error GoTo ErrReset
    newVal = Me.f1book_dyestudy.Entry
    TempProj.DyeStudy(i).time = newVal
    f1book_dyestudy.Col = 2
    newVal = Me.f1book_dyestudy.Entry
    TempProj.DyeStudy(i).concentration = newVal
    If newVal = -1 Then
      TempProj.dyestudy_count = i
      Exit For
    End If
  Next i
  
  Call DirtyFlag_Throw(TempProj)
  Call refresh_frmDyeStudy(TempProj)
  
ErrReset:
  Call Show_Error("Data in clipboard is not valid.")
  Call cmdReset_Click
  Exit Sub
  
End Sub

Private Sub cmdReset_Click()
  Dim i As Integer
  
  ReDim TempProj.DyeStudy(1 To 400)
  TempProj.dyestudy_count = 400
  For i = 1 To 400
    f1book_dyestudy.EntryRC(i, 1) = -1
    f1book_dyestudy.EntryRC(i, 2) = -1
  Next i
  For i = 1 To TempProj.dyestudy_count
      TempProj.DyeStudy(i).time = -1
      TempProj.DyeStudy(i).concentration = -1
    Next i
  txtData(0) = ""
  
  Call DirtyFlag_Throw(TempProj)
  Call refresh_frmDyeStudy(TempProj)
  
End Sub

Private Sub f1book_dyestudy_EndEdit(EditString As String, Cancel As Integer)
Dim idx As Integer
Dim newVal As Double
Dim oldVal As Double

  idx = f1book_dyestudy.Row
  If (idx < 1) Then Exit Sub

  Select Case f1book_dyestudy.Col
    Case 1:
      oldVal = TempProj.DyeStudy(idx).time
      newVal = CDbl(EditString)
      TempProj.DyeStudy(idx).time = newVal
      If oldVal = -1 Then
        TempProj.dyestudy_count = idx + 1
      End If
      
      'REFRESH DyeStudy window, ESPECIALLY THE GRIDS.
      Call refresh_frmDyeStudy(TempProj)
    Case 2:
      newVal = CDbl(EditString)
      TempProj.DyeStudy(idx).concentration = newVal
      
      'REFRESH DyeStudy window, ESPECIALLY THE GRIDS.
      Call refresh_frmDyeStudy(TempProj)
      
  End Select
  Call DirtyFlag_Throw(TempProj)
End Sub



Private Sub Form_Load()
  'MISC INITS.
  Call CenterOnForm(Me, frmDyeStudy)
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

