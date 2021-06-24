Attribute VB_Name = "INIFILES_pre_990817"
'Option Explicit
'
'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'
''Global fn_OldFileList As String           'full name including Windows path
'Global fn_FullFilename_INI As String
'Global Const fn_INI_name = "CPAS.INI"     'non-full name NOT including Windows path
'Global fpath_INI As String
'''Global OldFiles(1 To 4) As String
'
'
'
'
'
'Const INIFILES_declarations_end = 0
'
'
'Function GetWindowsDir() As String
'Dim Value As Integer
'Dim storage As String * 144
'  Value = GetWindowsDirectory(ByVal storage, ByVal Len(storage))
'  GetWindowsDir = Trim$(Left$(storage, Value))
'End Function
'
'
'Function GetWindowsTempDir() As String
'Dim retstr As String
'  retstr = Trim$(Environ$("TEMP"))
'  If (retstr = "") Then retstr = Trim$(Environ$("TMP"))
'  'IF NEITHER THE TEMP OR TMP ENVIRONMENT VARIABLES EXIST
'  'THEN WE USE THE WINDOWS DIRECTORY AS THE TEMPORARY
'  'FILE PATH.
'  If (retstr = "") Then
'    retstr = Trim$(GetWindowsDir())
'  End If
'  If (Right$(retstr, 1) = "\") Then
'    'REMOVE TRAILING BACKSLASH.
'    retstr = Left$(retstr, Len(retstr) - 1)
'  End If
'  GetWindowsTempDir = retstr
'End Function
'
'
'Sub ini_putsetting0(fn_ini As String, ini_header As String, INI_VarName As String, ini_newsetting As String)
'Dim lpApplicationName As String
'Dim lpKeyName As String
'Dim lpString As String
'Dim lpFileName As String
'Dim valid As Integer
'  lpApplicationName = ini_header
'  lpKeyName = INI_VarName
'  lpString = ini_newsetting
'  lpFileName = fn_ini
'  valid = WritePrivateProfileString(ByVal lpApplicationName, ByVal lpKeyName, ByVal lpString, ByVal lpFileName)
'End Sub
'
'Function INI_GetSetting0_Defaults( _
'    fn_ini As String, _
'    ini_header As String, _
'    INI_VarName As String, _
'    use_default_if_null As String) As String
'Dim retstr As String
'  retstr = Trim$(INI_GetSetting0(fn_ini, ini_header, INI_VarName))
'  If (retstr = "") Then retstr = use_default_if_null
'  INI_GetSetting0_Defaults = retstr
'End Function
'
'
'Function INI_GetSetting0(fn_ini As String, ini_header As String, INI_VarName As String) As String
'Dim lpApplicationName As String
'Dim lpKeyName As String
'Dim lpszDefault As String
'Dim lpReturnedString As String * 200
'Dim nSize As Integer
'Dim lpFileName As String
'
'Dim BytesCopied As Integer
'Dim temp As String
'
'  lpApplicationName = ini_header
'  lpKeyName = INI_VarName
'  lpszDefault = ""
'  lpReturnedString = ""
'  nSize = Len(lpReturnedString)
'  lpFileName = fn_ini
'
'  BytesCopied = GetPrivateProfileString(ByVal lpApplicationName, ByVal lpKeyName, ByVal lpszDefault, ByVal lpReturnedString, ByVal nSize, ByVal lpFileName)
'  temp = Left$(Trim$(lpReturnedString), BytesCopied)
'  INI_GetSetting0 = temp
'
'End Function
'
''Sub OldFileList_Populate(num_menu As Integer, menu0 As Menu, menu1 As Menu, menu2 As Menu, menu3 As Menu, menu4 As Menu)
''
''  OldFiles(1) = INI_GetSetting0(fn_OldFileList, "old_files", "old_file(1)")
''  OldFiles(2) = INI_GetSetting0(fn_OldFileList, "old_files", "old_file(2)")
''  OldFiles(3) = INI_GetSetting0(fn_OldFileList, "old_files", "old_file(3)")
''  OldFiles(4) = INI_GetSetting0(fn_OldFileList, "old_files", "old_file(4)")
''
''  Call OldFileList_UpdateMenu(num_menu, menu0, menu1, menu2, menu3, menu4)
''
''End Sub
''
''Sub OldFileList_Promote(fn_newfile As String, num_menu As Integer, menu0 As Menu, menu1 As Menu, menu2 As Menu, menu3 As Menu, menu4 As Menu)
''Dim i As Integer
''Dim found As Integer
''
''  'IF NOT IN CURRENT LIST, SHIFT 1-3 DOWN TO 2-4 AND REPLACE 1.
''  'IF IN CURRENT LIST, SAVE, SHIFT OTHERS DOWN, AND MOVE TO 1.
''  fn_newfile = LCase$(fn_newfile)
''  found = -1
''  For i = 1 To 4
''    If (Trim$(LCase$(fn_newfile)) = Trim$(LCase$(OldFiles(i)))) Then
''      found = i
''      Exit For
''    End If
''  Next i
''  If (found = -1) Then
''    For i = 4 To 2 Step -1
''      OldFiles(i) = OldFiles(i - 1)
''    Next i
''    OldFiles(1) = fn_newfile
''  Else
''    For i = found To 2 Step -1
''      OldFiles(i) = OldFiles(i - 1)
''    Next i
''    OldFiles(1) = fn_newfile
''  End If
''
''  'UPDATE MENU:
''  Call OldFileList_UpdateMenu(num_menu, menu0, menu1, menu2, menu3, menu4)
''
''  'UPDATE INI FILE:
''  For i = 1 To 4
''    Call ini_putsetting0(fn_OldFileList, "old_files", "old_file(" & Trim$(Str$(i)) & ")", OldFiles(i))
''  Next i
''
''End Sub
''
''
''Sub OldFileList_UpdateMenu(num_menu As Integer, menu0 As Menu, menu1 As Menu, menu2 As Menu, menu3 As Menu, menu4 As Menu)
''Dim found_at_least_one As Integer
''Dim i As Integer
''Dim mnu As Menu
''
''  found_at_least_one = False
''  For i = 1 To 4
''    If (i = 1) Then Set mnu = menu1
''    If (i = 2) Then Set mnu = menu2
''    If (i = 3) Then Set mnu = menu3
''    If (i = 4) Then Set mnu = menu4
''    If (OldFiles(i) <> "") Then
''      found_at_least_one = True
''      mnu.Caption = "&" & Trim$(Str$(i)) & " - " & OldFiles(i)
''      mnu.Visible = True
''    Else
''      mnu.Visible = False
''    End If
''  Next i
''  If (Not found_at_least_one) Then
''    menu0.Visible = False
''  Else
''    menu0.Visible = True
''  End If
''
''End Sub
