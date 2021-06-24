Attribute VB_Name = "DLL_Declarations"
Option Explicit

'Public Declare Sub AIRDENS Lib "asap1" ( _
    DG As Double, _
    TEMP As Double, _
    PRES As Double)

Public Declare Sub Go_GETCSPT Lib "asapptad" Alias "_GETCSPT@20" ( _
    CS As Double, _
    VQ As Double, _
    HC As Double, _
    CI As Double, _
    CE As Double)
''Public Declare Sub Go_GETCSPT Lib "asapptad" Alias "GETCSPT" ( _
    CS As Double, _
    VQ As Double, _
    HC As Double, _
    CI As Double, _
    CE As Double)

''''SUBROUTINE AIRDENS(DG, TEMP, PRES)
Public Declare Sub Go_AIRDENS Lib "asapptad" Alias "AIRDENS" ( _
    DG As Double, _
    TEMP As Double, _
    PRES As Double)


Public Declare Sub Go_TESTSUB1 _
    Lib "hokanson_test1" _
    Alias "TESTSUB1" ( _
    A1 As Double, _
    A2 As Double)


''''SUBROUTINE AIRDENS(DG, TEMP, PRES)
Public Declare Sub Go_Hokanson_AIRDENS _
    Lib "hokanson_test1" _
    Alias "AIRDENS" ( _
    DG As Double, _
    TEMP As Double, _
    PRES As Double)

Public Declare Sub Go_Hokanson_GETCSPT Lib _
    "hokanson_test1" Alias "_GETCSPT@20" ( _
    CS As Double, _
    VQ As Double, _
    HC As Double, _
    CI As Double, _
    CE As Double)

Public Declare Sub Go_asapptad_test3_GETCSPT Lib _
    "asapptad_test3" Alias "_GETCSPT@20" ( _
    CS As Double, _
    VQ As Double, _
    HC As Double, _
    CI As Double, _
    CE As Double)

Public Declare Sub Go_asapptad_test4_GETCSPT Lib _
    "asapptad_test4" Alias "_GETCSPT@20" ( _
    CS As Double, _
    VQ As Double, _
    HC As Double, _
    CI As Double, _
    CE As Double)


Const DLL_Declarations_declarations_end = True

