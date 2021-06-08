Option Strict Off
Option Explicit On
Module INIFILES

	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetPrivateProfileInt Lib "kernel32"  Alias "GetPrivateProfileIntA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetSystemDirectory Lib "kernel32"  Alias "GetSystemDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	Declare Function GetWindowsDirectory Lib "kernel32"  Alias "GetWindowsDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

	Public fn_OldFileList As String 'full name including path
	Public Const fn_INI_name As String = "ADS.INI" 'non-full name NOT including Windows path
	
	'OldFiles(i,j):
	'     i = menu code
	'     j = file within that menu (file 1,2,3, or 4, where
	'         file 1 is the most recently accessed file)
	'UPGRADE_WARNING: Lower bound of array OldFiles was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public OldFiles(5, 4) As String
	
	
	
	
	
	Const INIFILES_declarations_end As Short = 0
	
	
	Function GetWindowsDir() As String
		Dim Value As Short
		Dim storage As New VB6.FixedLengthString(144)
		Value = GetWindowsDirectory(storage.Value, Len(storage.Value))
		GetWindowsDir = Trim(Left(storage.Value, Value))
	End Function
	
	
	Function GetWindowsTempDir() As String
		Dim retstr As String
		retstr = Trim(Environ("TEMP"))
		If (retstr = "") Then retstr = Trim(Environ("TMP"))
		'IF NEITHER THE TEMP OR TMP ENVIRONMENT VARIABLES EXIST
		'THEN WE USE THE WINDOWS DIRECTORY AS THE TEMPORARY
		'FILE PATH.
		If (retstr = "") Then
			retstr = Trim(GetWindowsDir())
		End If
		If (Right(retstr, 1) = "\") Then
			'REMOVE TRAILING BACKSLASH.
			retstr = Left(retstr, Len(retstr) - 1)
		End If
		GetWindowsTempDir = retstr
	End Function
	
	
	Sub ini_putsetting0(ByRef fn_ini As String, ByRef ini_header As String, ByRef INI_VarName As String, ByRef ini_newsetting As String)
		Dim lpApplicationName As String
		Dim lpKeyName As String
		Dim lpString As String
		Dim lpFileName As String
		Dim valid As Short
		lpApplicationName = ini_header
		lpKeyName = INI_VarName
		lpString = ini_newsetting
		lpFileName = fn_ini
		valid = WritePrivateProfileString(lpApplicationName, lpKeyName, lpString, lpFileName)
	End Sub
	Sub INI_PutSetting(ByRef INI_VarName As String, ByRef NewSetting As String)
		Call ini_putsetting0(fn_OldFileList, "UserSettings", INI_VarName, NewSetting)
	End Sub
	
	
	Function INI_GetSetting00(ByRef fn_ini As String, ByRef ini_header As String, ByRef INI_VarName As String) As String
		INI_GetSetting00 = INI_GetSetting0(fn_ini, ini_header, INI_VarName)
	End Function
	Function INI_GetSetting0(ByRef fn_ini As String, ByRef ini_header As String, ByRef INI_VarName As String) As String
		Dim lpApplicationName As String
		Dim lpKeyName As String
		Dim lpszDefault As String
		Dim lpReturnedString As New VB6.FixedLengthString(200)
		Dim nSize As Short
		Dim lpFileName As String
		
		Dim BytesCopied As Short
		Dim temp As String
		
		lpApplicationName = ini_header
		lpKeyName = INI_VarName
		lpszDefault = ""
		lpReturnedString.Value = ""
		nSize = Len(lpReturnedString.Value)
		lpFileName = fn_ini
		
		BytesCopied = GetPrivateProfileString(lpApplicationName, lpKeyName, lpszDefault, lpReturnedString.Value, nSize, lpFileName)
		temp = Left(Trim(lpReturnedString.Value), BytesCopied)
		INI_GetSetting0 = temp
		
	End Function
	Function INI_Getsetting(ByRef INI_VarName As String) As String
		INI_Getsetting = INI_GetSetting0(fn_OldFileList, "UserSettings", INI_VarName)
	End Function


	Sub OldFileList_Populate(ByRef num_menu As Short, ByRef menu0 As System.Windows.Forms.ToolStripSeparator, ByRef menu1 As System.Windows.Forms.ToolStripMenuItem, ByRef menu2 As System.Windows.Forms.ToolStripMenuItem, ByRef menu3 As System.Windows.Forms.ToolStripMenuItem, ByRef menu4 As System.Windows.Forms.ToolStripMenuItem)
		OldFiles(num_menu, 1) = INI_GetSetting0(fn_OldFileList, "old_files", "old_file(" & Trim(Str(num_menu)) & ",1)")
		OldFiles(num_menu, 2) = INI_GetSetting0(fn_OldFileList, "old_files", "old_file(" & Trim(Str(num_menu)) & ",2)")
		OldFiles(num_menu, 3) = INI_GetSetting0(fn_OldFileList, "old_files", "old_file(" & Trim(Str(num_menu)) & ",3)")
		OldFiles(num_menu, 4) = INI_GetSetting0(fn_OldFileList, "old_files", "old_file(" & Trim(Str(num_menu)) & ",4)")
		Call OldFileList_UpdateMenu(num_menu, menu0, menu1, menu2, menu3, menu4)
	End Sub
	Sub OldFileList_Promote(ByRef fn_newfile As String, ByRef num_menu As Short, ByRef menu0 As System.Windows.Forms.ToolStripSeparator, ByRef menu1 As System.Windows.Forms.ToolStripMenuItem, ByRef menu2 As System.Windows.Forms.ToolStripMenuItem, ByRef menu3 As System.Windows.Forms.ToolStripMenuItem, ByRef menu4 As System.Windows.Forms.ToolStripMenuItem)
		Dim i As Short
		Dim Found As Short
		'IF NOT IN CURRENT LIST, SHIFT 1-3 DOWN TO 2-4 AND REPLACE 1.
		'IF IN CURRENT LIST, SAVE, SHIFT OTHERS DOWN, AND MOVE TO 1.
		fn_newfile = LCase(fn_newfile)
		Found = -1
		For i = 1 To 4
			If (Trim(LCase(fn_newfile)) = Trim(LCase(OldFiles(num_menu, i)))) Then
				Found = i
				Exit For
			End If
		Next i
		If (Found = -1) Then
			For i = 4 To 2 Step -1
				OldFiles(num_menu, i) = OldFiles(num_menu, i - 1)
			Next i
			OldFiles(num_menu, 1) = fn_newfile
		Else
			For i = Found To 2 Step -1
				OldFiles(num_menu, i) = OldFiles(num_menu, i - 1)
			Next i
			OldFiles(num_menu, 1) = fn_newfile
		End If
		'UPDATE MENU:
		Call OldFileList_UpdateMenu(num_menu, menu0, menu1, menu2, menu3, menu4)
		'UPDATE INI FILE:
		For i = 1 To 4
			Call ini_putsetting0(fn_OldFileList, "old_files", "old_file(" & Trim(Str(num_menu)) & "," & Trim(Str(i)) & ")", OldFiles(num_menu, i))
		Next i
	End Sub
	Sub OldFileList_UpdateMenu(ByRef num_menu As Short, ByRef menu0 As System.Windows.Forms.ToolStripItem, ByRef menu1 As System.Windows.Forms.ToolStripMenuItem, ByRef menu2 As System.Windows.Forms.ToolStripMenuItem, ByRef menu3 As System.Windows.Forms.ToolStripMenuItem, ByRef menu4 As System.Windows.Forms.ToolStripMenuItem)
		Dim found_at_least_one As Short
		Dim i As Short
		Dim mnu As System.Windows.Forms.ToolStripMenuItem
		found_at_least_one = False
		For i = 1 To 4
			If (i = 1) Then mnu = menu1
			If (i = 2) Then mnu = menu2
			If (i = 3) Then mnu = menu3
			If (i = 4) Then mnu = menu4
			If (OldFiles(num_menu, i) <> "") Then
				found_at_least_one = True
				mnu.Text = "&" & Trim(Str(i)) & " - " & OldFiles(num_menu, i)
				mnu.Visible = True
			Else
				mnu.Visible = False
			End If
		Next i
		If (Not found_at_least_one) Then
			menu0.Visible = False
		Else
			menu0.Visible = True
		End If
	End Sub
End Module