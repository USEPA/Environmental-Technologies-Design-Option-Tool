Attribute VB_Name = "FortranLink"
Option Explicit

Global FortranLink_Path As String
Global FortranLink_fn_MainInput As String
Global FortranLink_fn_MainOutput As String
'Global FortranLink_fn_FlowTimes As String
Global FortranLink_fn_FlowTimesInput As String
Global FortranLink_fn_BedDef() As String
Global FortranLink_fn_VarInf() As String
Global FortranLink_Version As String
Global FortranLink_fn_VarEff() As String
Global FortranLink_fn_MiscEff() As String
Global FortranLink_fn_Paths As String
Global FortranLink_EffluentStream_Count As Integer
Global FortranLink_fn_ErrorOutput As String


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


'Global FortranLink_fn_MainInput As String
'Global FortranLink_fn_MainOutput As String
'Global FortranLink_Components_Count As Integer
'Global FortranLink_fn_Components() As String
'Global FortranLink_fn_pH As String



Const FortranLink_declarations_end = 0


Public Sub FortranLink_ExecAndWaitForProcess(cmdline$)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ret&
  ' Initialize the STARTUPINFO structure:
  start.cb = Len(start)
  ' Start the shelled application:
  ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
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
      'outputstr$ = CStr(v)
      outputstr$ = Chr$(34) & CStr(v) & Chr$(34)
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


'Sub FortranLink_SetFilenames()
'Dim i As Integer
'  'SET FILENAMES.
'  FortranLink_Path = AppPath & "\SIMS\" & NowProj.code
'  FortranLink_fn_MainInput = FortranLink_Path & "\catreac_main.in"
'  FortranLink_fn_MainOutput = FortranLink_Path & "\catreac_main.out"
'  FortranLink_fn_ErrorOutput = FortranLink_Path & "\catreac_nflag.out"
'  FortranLink_fn_FlowTimesInput = FortranLink_Path & "\catreac_flowtimes.in"
'  'ReDim FortranLink_fn_BedDef(1 To NowProj.BedDefs_Count)
'  'For i = 1 To NowProj.BedDefs_Count
'  '  FortranLink_fn_BedDef(i) = FortranLink_Path & "\catreac_bed" & Trim$(Str$(i)) & ".in"
'  'Next i
'  'ReDim FortranLink_fn_VarInf(1 To NowInfluentConc_Count)
'  'For i = 1 To NowInfluentConc_Count
'  '  FortranLink_fn_VarInf(i) = FortranLink_Path & "\catreac_varinf" & Trim$(Str$(i)) & ".in"
'  'Next i
'  FortranLink_Version = "1.00"
'  'ReDim FortranLink_fn_VarEff(1 To FortranLink_EffluentStream_Count)
'  'ReDim FortranLink_fn_MiscEff(1 To FortranLink_EffluentStream_Count)
'  'For i = 1 To FortranLink_EffluentStream_Count
'  '  FortranLink_fn_VarEff(i) = FortranLink_Path & "\catreac_vareff" & Trim$(Str$(i)) & ".out"
'  '  FortranLink_fn_VarEff(i) = FortranLink_Path & "\catreac_misceff" & Trim$(Str$(i)) & ".out"
'  'Next i
'  FortranLink_fn_Paths = AppPath & "\catpath.txt"
'End Sub


'Sub FortranLink_WritePathFile()
'Dim f As Integer
'Dim fn_PathFile As String
'Dim qq As String
'  'SET FILENAMES.
'  Call FortranLink_SetFilenames
'  'OUTPUT FILENAME OF MAIN INPUTS.
'  fn_PathFile = FortranLink_fn_Paths
'  f = FreeFile
'  Open fn_PathFile For Output As #f
'  qq = Chr$(34)
'  Print #f, qq & Trim$(FortranLink_fn_MainInput) & qq
'  Close #f
'End Sub


