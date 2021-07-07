Attribute VB_Name = "DemoModeMod"
Option Explicit

'For Distribution on CD
Global Const Mode_Distribution_on_CD = True
Global Const Mode_Stapleton_Disclaimer = False

Global Const DemoMode = False
Global Const StudentMode = False
Global Const SecureDBMode = True

Type blist_save_type
    cas As Long
    Name As String * 40
End Type

Function decrypt_string(password As String) As String
Dim newpass$, i%, length%, keyval%
Dim key() As Integer

ReDim key(100) As Integer

For i% = 32 To 122
    keyval% = ((i% * 3) Mod 91)
    key(keyval%) = i%
Next i%

length% = Len(password)
newpass$ = ""
For i% = 1 To length%
    newpass$ = newpass$ + Chr$(key(Asc(Mid$(password, i%, 1)) - 32))
Next i%

Erase key

decrypt_string = newpass
End Function

'Function demo_check_chemicals(lbox As VListBox) As Integer
'    Dim demochem%, choice$, msg$, NL  As String
'
'
'If (Not DemoMode) Then
'    demo_check_chemicals = False
'    Exit Function
'End If
'
'
'choice$ = Left$(Trim$(lbox.List(0)), 6)
'
'demochem% = False
'If (0 = StrComp(choice$, "79016 ", 1)) Then demochem% = True
'If (0 = StrComp(choice$, "71432 ", 1)) Then demochem% = True
'If (0 = StrComp(choice$, "95501 ", 1)) Then demochem% = True
'
'If (demochem%) Then
'    demo_check_chemicals = False
'Else
'    NL = Chr$(10) + Chr$(13)
'    demo_check_chemicals = True
'    msg$ = "You are in  Demo Mode for this program." + NL
'    msg$ = msg$ + "You may only process one of the following 3 chemicals:" + NL + NL
'    msg$ = msg$ + "     79016 trichloroethylene" + NL
'    msg$ = msg$ + "     71432 benzene" + NL
'    msg$ = msg$ + "     95501 1,2-dichlorobenzene" + NL
'    MsgBox msg$
'End If
'
'
'End Function

Function fileexists(FileName As String) As Integer
    Dim test%

On Error GoTo ErrorIndexOpen

test% = GetAttr(FileName)

fileexists = True

Exit Function

ErrorIndexOpen:
    fileexists = False
    Exit Function
Resume Next

End Function

Sub read_blist_file()
    Dim reccnt%, i%
    Dim blist_tmp As blist_save_type

' CHECK FOR THE EXISTANCE OF THE BLIST.LST FILE.
' IF IT DOES NOT EXIST CREATE IT FROM THE DATABASE
If (Not fileexists(Database_Path + "\blist.lst")) Then
    write_blist_file
End If

contam_prop_form!contam_combo.Clear

reccnt% = 0

Open Database_Path + "\blist.lst" For Binary Access Read As #4
    Get #4, 1, reccnt%
    ReDim Preserve db_index(reccnt%)

    For i% = 1 To reccnt%
        Get #4, , blist_tmp
        db_index(i%) = blist_tmp.cas
        contam_prop_form!contam_combo.AddItem Str$(blist_tmp.cas) + "  " + Trim$(blist_tmp.Name)
    Next i%

Close #4

db_num_entries = reccnt%

End Sub

Sub write_blist_file()
    Dim reccnt%
    Dim blist_tmp As blist_save_type

'Read the database contaminant names and CAS numbers
'
' OPEN RECORDSET.
'
Set RS_Main = DB_Main.OpenRecordset( _
    "SELECT * FROM [Names (Master)] WHERE [Names (Master)].CAS = " & _
    Format$(dbinput.CasNumber, "0"))
If (RS_Main.EOF = False) Then
  RS_Main.MoveFirst
  RS_Main.MoveLast
  RS_Main.MoveFirst
End If
Set Selection = RS_Main
''''Set Selection = contam_prop_form!Data1.Recordset

reccnt% = 0
    
Open Database_Path + "\blist.lst" For Binary Access Write As #4
    Put #4, 1, reccnt%

    Selection.MoveFirst
    While Not Selection.EOF
        If Selection(0) = "On" Then
            reccnt% = reccnt% + 1
            blist_tmp.cas = Selection(1)
            blist_tmp.Name = Trim$(Selection(2))
            Put #4, , blist_tmp
        End If
        Selection.MoveNext
    Wend

' PUT THE RECORD COUNT AS THE FIRST INTEGER IN THE FILE
    Put #4, 1, reccnt%

Close #4

End Sub

