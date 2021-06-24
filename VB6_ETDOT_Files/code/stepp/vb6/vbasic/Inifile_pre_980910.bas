Option Explicit

'allows for *.ini calls
Declare Function GetPrivateProfileInt% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$)
Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
Declare Function WritePrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpString$, ByVal lplFileName$)

Declare Function GetWindowsDirectory Lib "kernel" (ByVal lpbuffer$, ByVal nSize%) As Integer

Global INI_ProgramType As String
Global INI_FileName As String

Function GetWindowsDir () As String
Dim Value As Integer
Dim storage As String * 144

  Value = GetWindowsDirectory(ByVal storage, ByVal Len(storage))
  GetWindowsDir = Trim$(Left$(storage, Value))

End Function

Function ini_getsetting (INI_FILE As String, INI_SpecifiedProgramType As String, INI_VariableName As String) As String

  'NOTE: THE FOLLOWING ARE DUMMY VARIABLES:
  '    - INI_FILE
  '    - INI_SpecifiedProgramType
  ini_getsetting = INI_GetSetting00(INI_FileName, INI_ProgramType, INI_VariableName)

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
'  lpApplicationName = INI_SpecifiedProgramType
'  lpKeyName = INI_VariableName
'  lpszDefault = ""
'  lpReturnedString = ""
'  nSize = Len(lpReturnedString)
'  lpFileName = INI_FILE
'
'  BytesCopied = GetPrivateProfileString(ByVal lpApplicationName, ByVal lpKeyName, ByVal lpszDefault, ByVal lpReturnedString, ByVal nSize, ByVal lpFileName)
'  temp = Trim$(Left$(lpReturnedString, BytesCopied))
'  ini_getsetting = temp
'
End Function

Function INI_GetSetting00 (use_Filename As String, use_Section As String, use_VarName As String)
Dim lpApplicationName As String
Dim lpKeyName As String
Dim lpszDefault As String
Dim lpReturnedString As String * 200
Dim nSize As Integer
Dim lpFileName As String

Dim BytesCopied As Integer
Dim temp As String

  'lpApplicationName = INI_SpecifiedProgramType
  'lpKeyName = INI_VariableName
  lpszDefault = ""
  lpReturnedString = ""
  nSize = Len(lpReturnedString)
  'lpFileName = INI_WindowsDir & "\" & INI_SpecifiedProgramType & ".ini"
  BytesCopied = GetPrivateProfileString(ByVal use_Section, ByVal use_VarName, ByVal lpszDefault, ByVal lpReturnedString, ByVal nSize, ByVal use_Filename)
  temp = Left$(Trim$(lpReturnedString), BytesCopied)
  INI_GetSetting00 = temp
  
End Function

Sub ini_putsetting (INI_VariableName As String, INI_NewSetting As String)
Dim lpApplicationName As String
Dim lpKeyName As String
Dim lpString As String
Dim lpFileName As String

Dim valid As Integer

  lpApplicationName = INI_ProgramType
  lpKeyName = INI_VariableName
  lpString = INI_NewSetting
  lpFileName = INI_FileName

  valid = WritePrivateProfileString(ByVal lpApplicationName, ByVal lpKeyName, ByVal lpString, ByVal lpFileName)

End Sub

