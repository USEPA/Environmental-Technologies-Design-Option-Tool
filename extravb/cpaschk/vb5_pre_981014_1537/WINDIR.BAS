Attribute VB_Name = "WINDIR"
Option Explicit

' THESE 3 VARIABLES WILL STORE THE PATHS FOR THE CURRENT
' WINDOWS SESSION.  THEY BOTH END WITH A BACKSLASH. = "C:\WINNT\"
Global WinPathSystem$
Global WinPathWindows$
Global WinPathTemp$

Global AppPath$

'Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
'Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long




Const WINDIR_declarations_end = True


' MAKE SURE THAT THERE IS A BACKSLASH AT THE END OF THE PATH
Function add_backslash(Path$) As String
  If (Right$(Path$, 1) = "\") Then
    add_backslash = Path$
  Else
    add_backslash = Path$ + "\"
  End If
End Function


Function FileExists(filepath$) As Integer
  On Error GoTo ERROR_FILEEXISTS
  FileExists = (Dir$(filepath$) <> "")
  Exit Function
ERROR_FILEEXISTS:
  FileExists = False
  Exit Function
End Function


Function GetWindowsDir() As String
Dim Value As Integer
Dim storage As String * 144
  Value = GetWindowsDirectory(ByVal storage, ByVal Len(storage))
  GetWindowsDir = Trim$(Left$(storage, Value))
End Function
Function GetWindowsSystemDir() As String
Dim Value As Integer
Dim storage As String * 144
  Value = GetSystemDirectory(ByVal storage, ByVal Len(storage))
  GetWindowsSystemDir = Trim$(Left$(storage, Value))
End Function


Function GetWindowsTempDir() As String
Dim retstr As String
  retstr = Trim$(Environ$("TEMP"))
  If (retstr = "") Then retstr = Trim$(Environ$("TMP"))
  'IF NEITHER THE TEMP OR TMP ENVIRONMENT VARIABLES EXIST
  'THEN WE USE THE WINDOWS DIRECTORY AS THE TEMPORARY
  'FILE PATH.
  If (retstr = "") Then
    retstr = Trim$(GetWindowsDir())
  End If
  If (Right$(retstr, 1) = "\") Then
    'REMOVE TRAILING BACKSLASH.
    retstr = Left$(retstr, Len(retstr) - 1)
  End If
  GetWindowsTempDir = retstr
End Function


Sub get_win_paths()
Dim tstr$, tmp%
    
  'tstr$ = Space$(144)
  'GetWindowsDir
  'tmp% = GetWindowsDirectory(tstr$, 144)
  WinPathWindows$ = add_backslash(UCase$(GetWindowsDir()))
  
  'tstr$ = Space$(144)
  'tmp% = GetSystemDirectory(tstr$, 144)
  WinPathSystem$ = add_backslash(UCase$(GetWindowsSystemDir()))
  
  tstr$ = Environ$("TEMP")
  If (tstr$ = "") Then tstr$ = Environ$("TMP")
  
  ' IF NEITHER THE TEMP OR TMP ENVIRONMENT VARIABLES EXIST
  ' THEN WE USE THE WINDOWS DIRECTORY AS THE TEMPORARY
  ' FILE PATH
  If (tstr$ = "") Then
    WinPathTemp$ = WinPathWindows$
  Else
    WinPathTemp$ = add_backslash(tstr$)
  End If
  AppPath$ = add_backslash(CStr(App.Path))
End Sub




