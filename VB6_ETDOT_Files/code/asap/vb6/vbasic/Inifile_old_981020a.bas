Attribute VB_Name = "IniFileMod"
Option Explicit

'allows for adsim.ini calls
Declare Function GetPrivateProfileInt% Lib "kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$)
Declare Function GetPrivateProfileString% Lib "kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
Declare Function WritePrivateProfileString% Lib "kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpString$, ByVal lplFileName$)

Declare Function GetWindowsDirectory Lib "kernel" (ByVal lpbuffer$, ByVal nSize%) As Integer
Declare Function GetSystemDirectory Lib "kernel" (ByVal lpbuffer$, ByVal nSize%) As Integer


'Global variables:
Global INI_WindowsDir As String
Global INI_ProgramType As String
Global INI_FileName As String

Function GetWindowsDir() As String
Dim storage As String * 144
Dim value As Integer

  value = GetWindowsDirectory(ByVal storage, ByVal Len(storage))
  GetWindowsDir = Trim$(Left$(storage, value))

End Function

Function GetWindowsSystemDir() As String
Dim value As Integer
Dim storage As String * 144

  value = GetSystemDirectory(ByVal storage, ByVal Len(storage))
  GetWindowsSystemDir = Trim$(Left$(storage, value))

End Function

Function INI_GetSetting(INI_VariableName As String) As String
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
'  lpApplicationName = INI_ProgramType
'  lpKeyName = INI_VariableName
'  lpszDefault = ""
'  lpReturnedString = ""
'  nSize = Len(lpReturnedString)
'  lpFileName = INI_FileName
'
'  BytesCopied = GetPrivateProfileString(ByVal lpApplicationName, ByVal lpKeyName, ByVal lpszDefault, ByVal lpReturnedString, ByVal nSize, ByVal lpFileName)
'  temp = Trim$(Left$(lpReturnedString, BytesCopied))
'  INI_GetSetting = temp

  'ini_getsetting = INI_GetSetting0(INI_ProgramType, INI_VariableName)
  On Error Resume Next
  INI_GetSetting = Trim$(INI_GetSetting00(INI_FileName, AppProgramKey, INI_VariableName))

End Function

Function INI_GetSetting0(INI_SpecifiedProgramType As String, INI_VariableName As String) As String
Dim lpApplicationName As String
Dim lpKeyName As String
Dim lpszDefault As String
Dim lpReturnedString As String * 200
Dim nSize As Integer
Dim lpFileName As String

Dim BytesCopied As Integer
Dim temp As String

  On Error Resume Next
  lpApplicationName = INI_SpecifiedProgramType
  lpKeyName = INI_VariableName
  lpszDefault = ""
  lpReturnedString = ""
  nSize = Len(lpReturnedString)
  lpFileName = INI_WindowsDir & "\" & INI_SpecifiedProgramType & ".ini"
  
  BytesCopied = GetPrivateProfileString(ByVal lpApplicationName, ByVal lpKeyName, ByVal lpszDefault, ByVal lpReturnedString, ByVal nSize, ByVal lpFileName)
  temp = Trim$(Left$(lpReturnedString, BytesCopied))
  INI_GetSetting0 = temp

End Function

Function INI_GetSetting00(use_Filename As String, use_Section As String, use_VarName As String)
Dim lpApplicationName As String
Dim lpKeyName As String
Dim lpszDefault As String
Dim lpReturnedString As String * 200
Dim nSize As Integer
Dim lpFileName As String

Dim BytesCopied As Integer
Dim temp As String

  On Error Resume Next
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

Sub ini_initializethisprogram(SpecifiedProgramType As String)
'Dim programtype As String
Dim infotype As String
Dim defaultresult As String
Dim returnvalue As String * 125
'dim ini_location As String
'ReDim appnames(3) As String
Dim storage As String * 144
Dim newdata As String
Dim defaultnumber As Long, valid As Long, string_size As Long
Dim f As Integer, i As Integer, value As Integer
Dim temp As String

  'Set global variable to specified program type
  INI_ProgramType = Trim$(SpecifiedProgramType)

  'get windows directory and look for adsim.ini
  value = GetWindowsDirectory(ByVal storage, ByVal Len(storage))
  INI_WindowsDir = Trim$(Left$(storage, value))

  'if does not exist, create ini file
  'INI_FileName = INI_WindowsDir & "\" & INI_ProgramType & ".ini"
  
  'SELECT INI FILENAME DIRECTORY.
  If (fileexists(Global_fpath_dir_CPAS & "\DBASE")) Then
    'USE THE DBASE DIRECTORY.
    INI_FileName = Global_fpath_dir_CPAS & "\DBASE\" & INI_ProgramType & ".ini"
  Else
    'VB3 IS FLAKY ABOUT LONG FILENAMES.
    'THEREFORE, PUT THE INI FILE INTO THE WINDOWS DIRECTORY.
    INI_FileName = GetWindowsDir() & "\" & INI_ProgramType & ".ini"
  End If

  On Error Resume Next
  If (Dir(INI_FileName)) = "" Then
    '======>  Program Specified Initialization!  <======
    f = FreeFile
    Open INI_FileName For Output As f
    Print #f, "[asap]"
    Print #f, "app_path="
    Print #f, "has_seen_disclaimer=0"
    Print #f, "oldfile1="
    Print #f, "oldfile2="
    Print #f, "oldfile3="
    Print #f, "oldfile4="

  'appnames(1) = "[adsim]"
  'appnames(2) = "[asap]"
  'appnames(3) = "[stepp]"
  '
  'For i% = 1 To 3
  '  Print #f, appnames(i)
  '  Write #f,
  '  Print #f, "has_seen_disclaimer=0"
  '  Write #f,
  '  Print #f, "oldfile1="
  '  Write #f,
  '  Print #f, "oldfile2="
  '  Write #f,
  '  Print #f, "oldfile3="
  '  Write #f,
  '  Print #f, "oldfile4="
  '  Write #f,
  '  Print #f, "path="
  '  Write #f,
  'Next i%
    Close #f
  End If

'setup variables for ini call
'programtype = "stepp" '**change to name of program
'ini_location = windowsdir
'defaultresult = ""
'infotype = "app_path"
'string_size = Len(returnvalue)
'returnvalue = ""
'value = GetPrivateProfileString(ByVal INI_ProgramType, ByVal infotype, ByVal defaultresult, ByVal returnvalue, ByVal string_size, ByVal INI_FileName)
'temp = Trim$(Left$(returnvalue, value))
  temp = INI_GetSetting("app_path")

  'if incorrect path set with programs current path being used now
  If ((StrComp(temp, App.Path) <> 0) Or (temp = "")) Then
    'newdata = Trim$(app.Path)
    'valid = WritePrivateProfileString(ByVal INI_ProgramType, ByVal infotype, ByVal newdata, ByVal INI_FileName)
    Call ini_putsetting("app_path", Trim$(App.Path))
  End If

  ChDir App.Path

End Sub

Sub ini_putsetting(INI_VariableName As String, INI_NewSetting As String)
Dim lpApplicationName As String
Dim lpKeyName As String
Dim lpString As String
Dim lpFileName As String

Dim valid As Integer

  On Error Resume Next
  lpApplicationName = INI_ProgramType
  lpKeyName = INI_VariableName
  lpString = INI_NewSetting
  lpFileName = INI_FileName

  valid = WritePrivateProfileString(ByVal lpApplicationName, ByVal lpKeyName, ByVal lpString, ByVal lpFileName)

End Sub

