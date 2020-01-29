Attribute VB_Name = "PackSelMod"
Option Explicit

'  THIS MODULE CONTAINS ALL THE DECLARATIONS AND  FUNCTIONS THAT RELATE
' TO THE READING AND WRITING OF THE DIFFERENT PACKING DATABASES.

Type PackingDataType
     Name As String
     Material As String
     source As String
     NominalSize As Double
     PackingFactor As Double
     CriticalSurfaceTension As Double
     SpecificSurfaceArea As Double
     OndaWettedSurfaceArea As Double
     SourceDatabase As Integer
     UserInput As Integer
     ValChanged As Integer
End Type

Global Const MAXDATABASEPACKINGS = 30
Global Const MAXUSERPACKINGS = 60

Global NumPackingsInDatabase As Integer
Global NumUserPackings As Integer

Global DatabasePacking() As PackingDataType

Global UserPacking(1 To MAXUSERPACKINGS) As PackingDataType

Global PackingChanged As Integer  'Whether current packing type is modified

Global Const ORIGINALPACKINGDATABASE = 1
Global Const USERMODIFIEDPACKINGDATABASE = 2

Global PackingDatabaseSource As Integer

Global PackingValuesChanged As Integer 'Whether user has modified any of the values for a selected packing

Global ShownPackingProperties As Integer 'Whether have shown the packing values on PTADScreen1.  Will be used to set the properties only once.

Sub ReadMainPackingDB()
    Dim i%

' Load Packing Parameters from File into Array
    Dim packingname$

packingname$ = "Tri-Packs_No.2"

' DEMO MODE CHANGE ::TACK
    If DemoMode% Then packingname$ = "Tri-Packs_No.1"
' END DEMO CHANGE

frmSelectPacking.cboSelectPacking.Clear

Open App.Path & "\dbase\PACKmain.db" For Binary As #1
Get #1, 1, NumPackingsInDatabase
ReDim DatabasePacking(1 To NumPackingsInDatabase) As PackingDataType

For i% = 1 To NumPackingsInDatabase
    Call ReadPackingDataType(1, DatabasePacking(i%))

    DatabasePacking(i%).SourceDatabase = ORIGINALPACKINGDATABASE
    DatabasePacking(i%).UserInput = False
    DatabasePacking(i%).ValChanged = False
    frmSelectPacking.cboSelectPacking.AddItem DatabasePacking(i%).Name

'  Set Default Packing
    If DatabasePacking(i%).Name = packingname$ Then
        DefaultPacking = DatabasePacking(i%)
    End If
Next i%

Close #1

frmSelectPacking.mnuPackDatabase(0).Checked = True
frmSelectPacking.fraPackingDatabase.Caption = "Original Database"
PackingDatabaseSource = ORIGINALPACKINGDATABASE
frmSelectPacking.mnuPackDatabase(3).Enabled = False
ShownPackingProperties = False
     
End Sub

Sub ReadPackingDataType(fnum As Integer, buf As PackingDataType)
    Dim strsize%

Get #fnum, , buf.NominalSize
Get #fnum, , buf.PackingFactor
Get #fnum, , buf.CriticalSurfaceTension
Get #fnum, , buf.SpecificSurfaceArea
Get #fnum, , buf.OndaWettedSurfaceArea
Get #fnum, , buf.SourceDatabase
Get #fnum, , buf.UserInput
Get #fnum, , buf.ValChanged

Get #fnum, , strsize%
buf.Name = String$(strsize%, " ")
Get #fnum, , buf.Name

Get #fnum, , strsize%
buf.Material = String$(strsize%, " ")
Get #fnum, , buf.Material

Get #fnum, , strsize%
buf.source = String$(strsize%, " ")
Get #fnum, , buf.source


End Sub

Sub ReadUserPackingDB()
    Dim i%

Open App.Path & "\dbase\PACKuser.db" For Binary As #1
Get #1, 1, NumUserPackings

For i% = 1 To NumUserPackings
    Call ReadPackingDataType(1, UserPacking(i%))

    UserPacking(i).SourceDatabase = USERMODIFIEDPACKINGDATABASE
Next i%

Close #1

End Sub

Sub WriteMainPackingDB()
    Dim i%

Open App.Path & "\dbase\PACKmain.db" For Binary As #1
Put #1, 1, NumPackingsInDatabase

For i% = 1 To NumPackingsInDatabase
    Call WritePackingDataType(1, DatabasePacking(i%))
Next i%

Close #1

End Sub

Sub WritePackingDataType(fnum As Integer, buf As PackingDataType)
    Dim strsize%

Put #fnum, , buf.NominalSize
Put #fnum, , buf.PackingFactor
Put #fnum, , buf.CriticalSurfaceTension
Put #fnum, , buf.SpecificSurfaceArea
Put #fnum, , buf.OndaWettedSurfaceArea
Put #fnum, , buf.SourceDatabase
Put #fnum, , buf.UserInput
Put #fnum, , buf.ValChanged

strsize% = Len(buf.Name)
Put #fnum, , strsize%
Put #fnum, , buf.Name

strsize% = Len(buf.Material)
Put #fnum, , strsize%
Put #fnum, , buf.Material

strsize% = Len(buf.source)
Put #fnum, , strsize%
Put #fnum, , buf.source

End Sub

Sub WriteUserPackingDB()
    Dim msg$
    Dim i%, response%

msg$ = ""
msg$ = msg$ + "Would you like to PERMANENTLY update "
msg$ = msg$ + "the changes in the user-modified database "
msg$ = msg$ + "to disk?  This cannot be undone."
response% = MsgBox(msg$, MB_ICONquestion + MB_YESNO, "")
If response% = IDYES Then
    'REDUNDANT, THEREFORE REMOVED, EJO 4/9/98.
    'Response% = MsgBox("This can not be undone.", MB_OKCANCEL + MB_ICONEXCLAMATION, "Warning")
    'If Response% = IDOK Then
        Open App.Path & "\dbase\PACKuser.db" For Binary As #1
        Put #1, 1, NumUserPackings
        
        For i% = 1 To NumUserPackings
            Call WritePackingDataType(1, UserPacking(i%))
        Next i%
    
        Close #1
    'End If
End If



End Sub

