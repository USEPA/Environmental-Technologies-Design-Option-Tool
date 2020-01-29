Attribute VB_Name = "FortranLink"
Option Explicit

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'      SHELL EXECUTION PROCESS -- BEGINS --
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type


Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
        hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
        lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
        lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As _
        PROCESS_INFORMATION) As Long

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'      SHELL EXECUTION PROCESS -- ENDS --
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


Global FortranLink_fn_MainInput As String
Global FortranLink_fn_MainOutput As String
Global FortranLink_Components_Count As Integer
Global FortranLink_fn_Components() As String
Global FortranLink_fn_pH As String



Const FortranLink_declarations_end = 0


Public Sub FortranLink_ExecAndWaitForProcess(CmdLine$)
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret&

' Initialize the STARTUPINFO structure:
start.cb = Len(start)

' Start the shelled application:
ret& = CreateProcessA(0&, CmdLine$, 0&, 0&, 1&, _
   NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

' Wait for the shelled application to finish:
ret& = WaitForSingleObject(proc.hProcess, INFINITE)
ret& = CloseHandle(proc.hProcess)

End Sub


Sub WriteFortranInput(f As Integer, v As Variant, s As String)
Dim outputstr$
Dim outlin As String

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
        MsgBox "WriteFortranInput vbObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbError
        MsgBox "WriteFortranInput vbError not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbDataObject
        MsgBox "WriteFortranInput vbDataObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbVariant
        MsgBox "WriteFortranInput vbVariant not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbArray
        MsgBox "WriteFortranInput vbArray not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbEmpty
        MsgBox "WriteFortranInput vbEmpty not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbNull
        MsgBox "WriteFortranInput vbNull not implemented"
        GoTo EXIT_FALSE_VALUE
  End Select

  outlin = Trim$(outputstr$)
  If (Len(outlin) > 27) Then
    outlin = outlin & "    "
  Else
    Do While (1 = 1)
      If (Len(outlin) >= 27) Then Exit Do
      outlin = outlin & " "
    Loop
  End If
  outlin = outlin & s
  Print #f, outlin
  GoTo EXIT_OK

EXIT_FALSE_VALUE:
  Print #f, "   - - - ERROR IN WriteFortranInput() - - -"
  Exit Sub
  
EXIT_OK:
  Exit Sub
  
End Sub


Sub FortranLink_SetFilenames()
Dim i As Integer
Dim temp_ntarget As Integer
  'SET FILENAMES.
  FortranLink_fn_MainInput = App.Path & "\exes\adox_in.txt"
  FortranLink_fn_MainOutput = App.Path & "\exes\adox_out.txt"
  temp_ntarget = NowProj.TargetCompounds_Count - 1    'do not include NOM.
  'NowProj.ncomp = (NowProj.ntarget * 2) + 14
  'TEMPORARILY, ASSUME THE MAXIMUM # OF COMPONENTS.
  FortranLink_Components_Count = _
      ((temp_ntarget * 2) + 14)
      '(temp_ntarget) + (temp_ntarget + 6) + 7
  ReDim FortranLink_fn_Components(1 To FortranLink_Components_Count)
  For i = 1 To FortranLink_Components_Count
    FortranLink_fn_Components(i) = App.Path & "\exes\" & _
        "comp" & Trim$(Str$(i)) & ".txt"
  Next i
  FortranLink_fn_pH = App.Path & "\exes\cmp_ph.txt"
End Sub


Sub FortranLink_WritePathFile()
Dim f As Integer
Dim fn_PathFile As String
Dim i As Integer
Dim qq As String

  'SET FILENAMES.
  Call FortranLink_SetFilenames

  'OUTPUT FILENAMES.
  fn_PathFile = App.Path & "\exes\adoxpath.txt"
  f = FreeFile
  Open fn_PathFile For Output As #f
  qq = Chr$(34)
  Print #f, qq & Trim$(FortranLink_fn_MainInput) & qq
  Print #f, qq & Trim$(FortranLink_fn_MainOutput) & qq
  Print #f, qq & Trim$(FortranLink_fn_pH) & qq
  Print #f, qstr(FortranLink_Components_Count)
  For i = 1 To FortranLink_Components_Count
    Print #f, qq & Trim$(FortranLink_fn_Components(i)) & qq
  Next i
  Close #f

End Sub


Sub FortranLink_KillCompoundOutputFiles()
Dim fn_spec As String
Dim fn_this As String
  fn_spec = App.Path & "\exes\comp*.txt"
  fn_this = Trim$(Dir(fn_spec))
  Do While (1 = 1)
    If (fn_this = "") Then Exit Do
    Call Kill_If_It_Exists(App.Path & "\exes\" & fn_this)
    fn_this = Dir
  Loop
  fn_this = App.Path & "\exes\cmp_ph.txt"
  Call Kill_If_It_Exists(fn_this)
End Sub

Sub Kill_If_It_Exists(fn As String)
  If (FileExists(fn)) Then
    Kill fn
  End If
End Sub
'show_error

Sub FortranLink_Run()
Dim fn_FortranModuleEXE As String
Dim fpath_run As String
Dim fpath_save As String
Dim success As Integer

Dim calctime_start As String
Dim calctime_end As String
Dim msg As String
Dim elapsed_min As Double
Dim CmdLine As String

  'REMOVE ALL PREVIOUS COMPOUND OUTPUT FILES (IF ANY).
  Call FortranLink_KillCompoundOutputFiles
  
  'WRITE INPUT FILES FOR FORTRAN MODULE.
  Call FortranLink_WritePathFile
  Call FortranLink_WriteInputFile
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput)
  
  'CALL FORTRAN MODULE.
  Call ChangeDir_Exes
  fn_FortranModuleEXE = "adoxfor.exe"
  CmdLine = fn_FortranModuleEXE
  calctime_start = Now
  Call FortranLink_ExecAndWaitForProcess(CmdLine)
  calctime_end = Now
  Call ChangeDir_Main
    
  'fpath_save = CurDir$
  'fpath_run = App.Path & "\exes"
  'ChDir fpath_run
  'ChDrive fpath_run
  'fn_FortranModuleEXE = App.Path & "\exes\adoxfor.exe"
  'calctime_start = Now
  'Call FortranLink_ExecAndWaitForProcess(fn_FortranModuleEXE)
  'calctime_end = Now
  'ChDir fpath_save
  'ChDrive fpath_save
  
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
  Else
    msg = "Calculations failed."
  End If
  Call Show_Error(msg) 'show_error
  
End Sub


