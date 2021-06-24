Option Strict Off
Option Explicit On
Module FortranLink
	
	Public FortranLink_Path As String
	Public FortranLink_fn_MainInput As String
	Public FortranLink_fn_MainOutput As String
	'Global FortranLink_fn_FlowTimes As String
	Public FortranLink_fn_FlowTimesInput As String
	Public FortranLink_fn_BedDef() As String
	Public FortranLink_fn_VarInf() As String
	Public FortranLink_Version As String
	Public FortranLink_fn_VarEff() As String
	Public FortranLink_fn_MiscEff() As String
	Public FortranLink_fn_Paths As String
	Public FortranLink_EffluentStream_Count As Short
	Public FortranLink_fn_ErrorOutput As String
	
	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'      SHELL EXECUTION PROCESS -- BEGINS --
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	
	Private Structure STARTUPINFO
		Dim cb As Integer
		Dim lpReserved As String
		Dim lpDesktop As String
		Dim lpTitle As String
		Dim dwX As Integer
		Dim dwY As Integer
		Dim dwXSize As Integer
		Dim dwYSize As Integer
		Dim dwXCountChars As Integer
		Dim dwYCountChars As Integer
		Dim dwFillAttribute As Integer
		Dim dwFlags As Integer
		Dim wShowWindow As Short
		Dim cbReserved2 As Short
		Dim lpReserved2 As Integer
		Dim hStdInput As Integer
		Dim hStdOutput As Integer
		Dim hStdError As Integer
	End Structure
	
	Private Structure PROCESS_INFORMATION
		Dim hProcess As Integer
		Dim hThread As Integer
		Dim dwProcessID As Integer
		Dim dwThreadID As Integer
	End Structure
	
	
	Private Const NORMAL_PRIORITY_CLASS As Integer = &H20
	Private Const INFINITE As Short = -1
	Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
	Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
	'UPGRADE_WARNING: Structure PROCESS_INFORMATION may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure STARTUPINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Integer, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Integer, ByVal lpThreadAttributes As Integer, ByVal bInheritHandles As Integer, ByVal dwCreationFlags As Integer, ByVal lpEnvironment As Integer, ByVal lpCurrentDirectory As Integer, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Integer
	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'      SHELL EXECUTION PROCESS -- ENDS --
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	
	
	'Global FortranLink_fn_MainInput As String
	'Global FortranLink_fn_MainOutput As String
	'Global FortranLink_Components_Count As Integer
	'Global FortranLink_fn_Components() As String
	'Global FortranLink_fn_pH As String
	
	
	
	Const FortranLink_declarations_end As Short = 0
	
	
	Public Sub FortranLink_ExecAndWaitForProcess(ByRef cmdline As String)
		Dim proc As PROCESS_INFORMATION
		Dim start As STARTUPINFO
		Dim ret As Integer
		' Initialize the STARTUPINFO structure:
		start.cb = Len(start)
		' Start the shelled application:
		ret = CreateProcessA(0, cmdline, 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, 0, start, proc)
		' Wait for the shelled application to finish:
		ret = WaitForSingleObject(proc.hProcess, INFINITE)
		ret = CloseHandle(proc.hProcess)
	End Sub
	
	
	Sub WriteFortranInput(ByRef f As Short, ByRef v As Object, ByRef s As String)
		Dim outputstr As String
		Dim outlin As String
		'UPGRADE_ISSUE: Constant vbDataObject was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case VarType(v)
			Case VariantType.Boolean
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				outputstr = IIf(v, "1", "0")
			Case VariantType.Byte, VariantType.Short, VariantType.Integer, VariantType.Decimal
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				outputstr = Trim(CStr(v))
			Case VariantType.Single, VariantType.Double
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				outputstr = Trim(CStr(v))
			Case VariantType.String, VariantType.Date
				'outputstr$ = CStr(v)
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				outputstr = Chr(34) & CStr(v) & Chr(34)
			Case VariantType.Object
				MsgBox("WriteFortranInput vbObject not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Error
				MsgBox("WriteFortranInput vbError not implemented")
				GoTo EXIT_FALSE_VALUE
				'	Case vbDataObject
				'		MsgBox("WriteFortranInput vbDataObject not implemented")
		'		GoTo EXIT_FALSE_VALUE
			Case VariantType.Object
				MsgBox("WriteFortranInput vbVariant not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Array
				MsgBox("WriteFortranInput vbArray not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Empty
				MsgBox("WriteFortranInput vbEmpty not implemented")
				GoTo EXIT_FALSE_VALUE
			Case VariantType.Null
				MsgBox("WriteFortranInput vbNull not implemented")
				GoTo EXIT_FALSE_VALUE
		End Select
		outlin = Trim(outputstr)
		If (Len(outlin) > 27) Then
			outlin = outlin & "    "
		Else
			Do While (1 = 1)
				If (Len(outlin) >= 27) Then Exit Do
				outlin = outlin & " "
			Loop 
		End If
		outlin = outlin & s
		PrintLine(f, outlin)
		GoTo EXIT_OK
EXIT_FALSE_VALUE: 
		PrintLine(f, "   - - - ERROR IN WriteFortranInput() - - -")
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
End Module