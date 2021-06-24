Attribute VB_Name = "PropTranMod"
'Option Explicit
''allows for *.ini calls
'Declare Function GetPrivateProfileInt% Lib "kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$)
'Declare Function GetPrivateProfileString% Lib "kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
'Declare Function WritePrivateProfileString% Lib "kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpString$, ByVal lplFileName$)
'
'Declare Function GetWindowsDirectory Lib "kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
'
''window handle functions
'Declare Function IsWindowVisible Lib "USER" (ByVal hWnd As Integer) As Integer
'Declare Function GetDesktopWindow Lib "USER" () As Integer
'Declare Function GetWindow Lib "USER" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
'Declare Function GetWindowText Lib "USER" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'Declare Function SetFocusAPI Lib "USER" Alias "SetFocus" (ByVal hWnd As Integer) As Integer
'
''PropFind record for each chemical to be transfered
'Type PropFind_Transfer_Record
'    Name As String * 50
'    VaporPressure As Double
'    IDAC As Double
'    HenrysConstant As Double
'    MolecularWeight As Double
'    NBP As Double
'    LiquidDensity As Double
'    MolarVol_OPT As Double
'    MolarVol_NBP As Double
'    RefractiveIndex As Double
'    AqueousSolubility As Double
'    OctanolWaterPartCoeff As Double
'    LiquidDiffusivity As Double
'    GasDiffusivity As Double
'
'    WaterDensity As Double
'    WaterViscosity As Double
'    WaterSurfaceTension As Double
'    AirDensity As Double
'    AirViscosity As Double
'
'End Type
'
''windows constants
'Global Const GW_CHILD = 5
'Global Const GW_OWNER = 4
'Global Const GW_HWNDNEXT = 2
'
'
''Global variables:
'Global INI_WindowsDir As String
'Global INI_ProgramType As String
'Global INI_FileName As String
'
'Global CallingProgram As String
'Global CallingAppFullname As String
'Global INI_TransferFile As String
'
'Global TransferMode As Integer
'Global propfind_hwnd As Integer
'Global TransferFilePath  As String
'
'Dim transfer_record As PropFind_Transfer_Record
'
'Function ActivateProgram(apptitle As String) As Integer
'    Dim hWnd, ret As Integer
'    Dim wtitle As String * 256
'    Dim tmpstr$
'    Dim fview_try_cnt As Integer
'    Dim fview_hwnd As Integer
'    Dim fview_title As String
''  GET THE DESKTOP WINDOW HANDLE
'hWnd = GetDesktopWindow()
'
''  GET THE DESKTOPS FIRST CHILD.  THAT WILL BE THE
''  FIRST WINDOW IN THE TASK LIST
'hWnd = GetWindow(hWnd, GW_CHILD)
'
''  LOOP UNTIL YOU FIND THE WINDOW HANDLE THAT YOU WANT
'Do While (hWnd <> 0)
'    ret = GetWindowText(hWnd, wtitle, 256)
'    tmpstr$ = Left$(wtitle, ret)
'
''  MAKING SURE WINDOW IS VISIBLE AND TOP WINDOW.
''  GETS RID OF ALL THE NON ESSENTIAL WINDOWS IN THE LIST
'    If (IsWindowVisible(hWnd) <> 0) And (GetWindow(hWnd, GW_OWNER) = 0) Then
'        If (InStr(1, tmpstr$, apptitle, 1) <> 0) Then
'            fview_try_cnt = 0
'            fview_hwnd = hWnd
'            fview_title = tmpstr$
'            ActivateProgram = True
'
'            'okay now that you found the window give it focus!
'            hWnd = SetFocusAPI(fview_hwnd)
'
'            Exit Function
'        End If
'    End If
'    hWnd = GetWindow(hWnd, GW_HWNDNEXT)
'Loop
'
'ActivateProgram = False
'
'End Function
'
'Function ExecutePropFindProgram() As Integer
'Dim response As String
'
''okay so since it does not exist find it
'response = INI_Getsetting_PropTranMod(INI_TransferFile, "PropTran", "ProgramPath")
'
''run program and store handle of program
'propfind_hwnd = Shell(response, 1)
'
'If propfind_hwnd = 0 Then
'    ExecutePropFindProgram = True
'Else
'    ExecutePropFindProgram = False
'End If
'
'End Function
'
''Function GetWindowsDir() As String
''Dim Value As Integer
''Dim storage As String * 144
''
''  Value = GetWindowsDirectory(ByVal storage, ByVal Len(storage))
''  GetWindowsDir = Trim$(Left$(storage, Value))
''
''End Function
'
'Function INI_Getsetting_PropTranMod(INI_FILE As String, INI_SpecifiedProgramType As String, INI_VariableName As String) As String
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
'  INI_Getsetting_PropTranMod = temp
'
'End Function
'
'Sub INI_PutSetting_PropTranMod(INI_VariableName As String, ini_newsetting As String)
'Dim lpApplicationName As String
'Dim lpKeyName As String
'Dim lpString As String
'Dim lpFileName As String
'
'Dim valid As Integer
'
'  lpApplicationName = INI_ProgramType
'  lpKeyName = INI_VariableName
'  lpString = ini_newsetting
'  lpFileName = INI_FileName
'
'  valid = WritePrivateProfileString(ByVal lpApplicationName, ByVal lpKeyName, ByVal lpString, ByVal lpFileName)
'
'End Sub
'
'Function ReadPropFindRecord(FileName$) As Integer
'    Dim FileNum As Integer
'    Dim i As Integer
'
'If (Not fileexists(FileName$)) Then
'    ReadPropFindRecord = False
'    Exit Function
'End If
'
'FileNum = FreeFile
'Open FileName$ For Binary Access Read As FileNum
'  For i = 1 To NumSelectedChemicals
'    Get #FileNum, 1024 * i, transfer_record(i)
'  Next i
'Close FileNum
'
'ReadPropFindRecord = True
'
'End Function
'
'Sub SaveTransferFile(FileName$)
'    Dim i As Integer
'
'ReDim transfer_record(NumSelectedChemicals)
'
'For i = 1 To NumSelectedChemicals
'
'    transfer_record(i) = PropContaminant(i)
'
'Next i
'
'Call WritePropFindRecord(FileName$)
'
'End Sub
'
'Sub WritePropFindRecord(FileName$)
'    Dim FileNum As Integer
'
'FileNum = FreeFile
'Open FileName$ For Binary Access Read As FileNum
'
'  For i = 1 To NumSelectedChemicals
'    Get #FileNum, 1024 * i, transfer_record(i)
'  Next i
'
'Close FileNum
'
'End Sub
'
