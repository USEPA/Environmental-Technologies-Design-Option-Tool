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



Function qstr(v As Variant) As String
  qstr = NumberToMFBString(v)
End Function



