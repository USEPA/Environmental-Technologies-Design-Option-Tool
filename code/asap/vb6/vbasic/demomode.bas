Attribute VB_Name = "DemoModeMod"
Option Explicit

' DEMO MODE GLOBAL VARIABLE

'For Distribution on CD
    Global Const Mode_Distribution_on_CD = True

'For Demo Mode use this setting
'   Global Const DemoMode% = false

'For Complete Version (non-Demo) use this setting
   Global Const DemoMode% = False
   Global Const StudentMode% = False

Sub check_area()

If 0 <> StrComp(App.Path, "w:\winapps\asap", 1) Then
    MsgBox "This Program is not being run on the correct System!" + Chr$(13) + "Therefore it will not work."
    End
End If

End Sub

Function decrypt_string(password As String) As String
Dim newpass$
Dim keyval%, i%, length%
ReDim Key(100) As Integer

For i% = 32 To 122
    keyval% = ((i% * 3) Mod 91)
    Key(keyval%) = i%
Next i%

length% = Len(password)
newpass$ = ""
For i% = 1 To length%
    newpass$ = newpass$ + Chr$(Key(Asc(Mid$(password, i%, 1)) - 32))
Next i%

decrypt_string = newpass$
End Function

Function demomode_check_packing(packingname As String) As Integer
    Dim packing_found%
    Dim msg$

If (Not DemoMode%) Then
    demomode_check_packing = 0
    Exit Function
End If

packing_found% = 0
If packingname = "Tri-Packs_No.1" Then packing_found% = 1
If packingname = "Tri-Packs_No.2" Then packing_found% = 1


If packing_found% Then
    demomode_check_packing = 0
Else
    demomode_check_packing = 1
    msg$ = "This Program is in Demo Mode." + Chr$(13) + Chr$(13)
    msg$ = msg$ + "You may on choose from the following Packing Materials:" + Chr$(13) + Chr$(13)
    msg$ = msg$ + Chr$(9) + "Tri-Packs_No.1" + Chr$(13)
    msg$ = msg$ + Chr$(9) + "Tri-Packs_No.2" + Chr$(13)
    MsgBox msg$
End If

End Function

Function fileexists(Filename As String) As Integer
    Dim test%

On Error GoTo ErrorIndexOpen

test% = GetAttr(Filename)

fileexists = True

Exit Function

ErrorIndexOpen:
    fileexists = False
    Exit Function
Resume Next

End Function

