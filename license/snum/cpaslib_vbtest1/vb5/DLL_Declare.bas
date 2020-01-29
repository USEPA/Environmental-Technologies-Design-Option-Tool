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

'extern "C" int __stdcall snumCpasLicGenerate(
'    char *spCpasDir,
'    char *spWinDir,
'    char *spNumber,
'    char *spUserName,
'    char *spUserCompany)
Declare Function snumCpasLicGenerate Lib "cpaslib.dll" ( _
    ByVal spCpasDir As String, _
    ByVal spWinDir As String, _
    ByVal spNumber As String, _
    ByVal spUserName As String, _
    ByVal spUserCompany As String) As Integer






Const DLL_Declare_declarations_end = True


Private Sub old_Command1_Click()
'Dim iLen As Long
'Dim x_spTest As String
'Dim spTest(1 To 100) As Byte
'Dim retVal As Long
'Dim i As Integer
'Dim iModules(0 To 10) As Long
'  'PERFORM TEST CALL TO SerialTest().
'  ChDir "X:\etdot10\license\snum\vbtest1\vb5"
'  ChDrive "X:\etdot10\license\snum\vbtest1\vb5"
'  MsgBox "About to call SerialTest() ..."
'  x_spTest = "This is the original testing string."
'  iLen = Len(x_spTest)
'  retVal = SerialTest(iLen, x_spTest)
'  MsgBox "Returned from call = " & Trim$(Str$(retVal))
'
'  'PERFORM TEST CALL TO SerialTest2().
'  ChDir "X:\etdot10\license\snum\vbtest1\vb5"
'  ChDrive "X:\etdot10\license\snum\vbtest1\vb5"
'  MsgBox "About to call SerialTest2() ..."
'  x_spTest = "This is the original testing string."
'  iLen = Len(x_spTest)
'  retVal = SerialTest2(iLen)
'  MsgBox "Returned from call = " & Trim$(Str$(retVal))
'
'  'PERFORM TEST CALL TO SerialTest3().
'  ChDir "X:\etdot10\license\snum\vbtest1\vb5"
'  ChDrive "X:\etdot10\license\snum\vbtest1\vb5"
'  MsgBox "About to call SerialTest3() ..."
'  x_spTest = "This is the original testing string."
'  iLen = Len(x_spTest)
'  retVal = SerialTest3(iLen, x_spTest)
'  MsgBox "New string value = " & x_spTest & _
'      "; returned from call = " & Trim$(Str$(retVal))
'
'  'PERFORM TEST CALL TO SerialTest4().
'  ChDir "X:\etdot10\license\snum\vbtest1\vb5"
'  ChDrive "X:\etdot10\license\snum\vbtest1\vb5"
'  MsgBox "About to call SerialTest4() ..."
'  iModules(0) = 0
'  iModules(1) = 1
'  iModules(2) = 2
'  iModules(3) = 3
'  iModules(4) = 4
'  retVal = SerialTest4(iModules(0))
'  MsgBox "Returned from call = " & Trim$(Str$(retVal))
'
End Sub




'OLD STUFF:
''extern "C" int SerialTest( int iLen, char *spTest )
'Declare Function SerialTest Lib "snum.dll" _
'    (ByVal iLen As Integer, ByVal spTest As String) As Integer
'
''extern "C" int SerialTest2( int iLen )
'Declare Function SerialTest2 Lib "snum.dll" _
'    (ByVal iLen As Integer) As Integer
'
''extern "C" int SerialTest3( int iLen, char *spTest )
'Declare Function SerialTest3 Lib "snum.dll" _
'    (ByVal iLen As Integer, ByVal spTest As String) As Integer
'
''extern "C" int SerialTest4( int iModules[] )
'Declare Function SerialTest4 Lib "snum.dll" _
'    (ByRef iModules As Long) As Integer





