Attribute VB_Name = "DLL_Declarations"
Option Explicit

'Public Declare Sub AIRDENS Lib "asap1" ( _
    DG As Double, _
    TEMP As Double, _
    PRES As Double)

Public Declare Sub GETCSPT Lib "asapptad" Alias "_GETCSPT@20" ( _
    CS As Double, _
    VQ As Double, _
    HC As Double, _
    CI As Double, _
    CE As Double)


Const DLL_Declarations_declarations_end = True

