Attribute VB_Name = "DLL_Declare"
Option Explicit

'extern "C" int __stdcall snumGenerate(char *spNumber, int iModules[], int iVersionType, int iExpires,
'                           int iExpiresDay, int iExpiresMonth, int iExpiresYear,
'                           long longInternalSnum, int iCheck);
Declare Function snumGenerate Lib "cpaslib.dll" ( _
    ByVal spNumber As String, _
    ByRef iModules As Long, _
    ByVal iVersionType As Long, _
    ByVal iExpires As Long, _
    ByVal iExpiresDay As Long, _
    ByVal iExpiresMonth As Long, _
    ByVal iExpiresYear As Long, _
    ByVal longInternalSnum As Long, _
    ByVal iCheck As Long) As Integer

'extern "C" int __stdcall snumVerify(char *spNumber)
Declare Function snumVerify Lib "cpaslib.dll" ( _
    ByVal spNumber As String) As Integer

'extern "C" int __stdcall snumIsModulePurchased(char *spNumber, int iModule)
Declare Function snumIsModulePurchased Lib "cpaslib.dll" ( _
    ByVal spNumber As String, _
    ByVal iModule As Long) As Integer

'extern "C" int __stdcall snumIsExpirationPresent(char *spNumber);
Declare Function snumIsExpirationPresent Lib "cpaslib.dll" ( _
    ByVal spNumber As String) As Integer

'extern "C" int __stdcall snumGetExpirationDay(char *spNumber);
Declare Function snumGetExpirationDay Lib "cpaslib.dll" ( _
    ByVal spNumber As String) As Integer

'extern "C" int __stdcall snumGetExpirationMonth(char *spNumber);
Declare Function snumGetExpirationMonth Lib "cpaslib.dll" ( _
    ByVal spNumber As String) As Integer

'extern "C" int __stdcall snumGetExpirationYear(char *spNumber);
Declare Function snumGetExpirationYear Lib "cpaslib.dll" ( _
    ByVal spNumber As String) As Integer

'extern "C" int __stdcall snumGetVersionType(char *spNumber);
Declare Function snumGetVersionType Lib "cpaslib.dll" ( _
    ByVal spNumber As String) As Integer






Const DLL_Declare_declarations_end = True


'RETURNS:
'    TRUE = RETURNED SUCCESSFULLY.
'    FALSE = FAILED TO GENERATE THAT SNUM.
Function Call_snumGenerate( _
    out_spNumber As String, _
    in_iModules() As Integer, _
    in_iVersionType As Integer, _
    in_iExpires As Integer, _
    in_iExpiresDay As Integer, _
    in_iExpiresMonth As Integer, _
    in_iExpiresYear As Integer, _
    in_longInternalSnum As Long) _
    As Boolean
Dim spNumber As String * 100
Dim iModules(0 To 49) As Long
Dim iVersionType As Integer
Dim iExpires As Integer
Dim iExpiresDay As Integer
Dim iExpiresMonth As Integer
Dim iExpiresYear As Integer
Dim longInternalSnum As Long
Dim iCheck As Integer
Dim RetVal As Integer
Dim i As Integer
  ''CHANGE DIRECTORIES (IS THIS STILL REQUIRED?).
  'ChDir "X:\etdot10\license\snum\cpaslib_vbtest1\vb5"
  'ChDrive "X:\etdot10\license\snum\cpaslib_vbtest1\vb5"
  'COPY PARAMETERS INTO TRANSFER VARIABLES.
  For i = 0 To 49
    iModules(i) = CLng(in_iModules(i))
  Next i
  iVersionType = CInt(in_iVersionType)
  iExpires = CInt(in_iExpires)
  iExpiresDay = CInt(in_iExpiresDay)
  iExpiresMonth = CInt(in_iExpiresMonth)
  iExpiresYear = CInt(in_iExpiresYear)
  longInternalSnum = CLng(in_longInternalSnum)
  iCheck = 13892
  'MAKE THE DLL CALL.
  ''''Call ChangeDir_Main
  RetVal = snumGenerate( _
    spNumber, _
    iModules(0), _
    iVersionType, _
    iExpires, _
    iExpiresDay, _
    iExpiresMonth, _
    iExpiresYear, _
    longInternalSnum, _
    iCheck)
  'RETURN WHETHER THE CALL WORKED OR NOT.
  If (RetVal = 0) Then
    Call_snumGenerate = False
  Else
    Call_snumGenerate = True
    out_spNumber = ""
    For i = 1 To 100
      If (Mid$(spNumber, i, 1) = Chr$(0)) Then Exit For
      out_spNumber = out_spNumber & Mid$(spNumber, i, 1)
    Next i
  End If
End Function


