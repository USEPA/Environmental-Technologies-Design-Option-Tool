Attribute VB_Name = "Do_Database_Op"
Option Explicit

Type MergeData_Type
  CAS As String
  Name_Old As String
  Name_New As String
End Type






Const Do_Database_Op_decl_end = True


Function Scan_Isotherms_Table() _
    As Boolean
On Error GoTo err_ThisFunc
Dim fn(1 To 2) As String
Dim Db(1 To 2) As Database
Dim Rs_Iso(1 To 2) As Recordset
Dim Rs_Chem(1 To 2) As Recordset
Dim i As Integer
Dim f As Integer
Dim ThisCas_Iso(1 To 2) As Long
Dim ThisCas_Chem(1 To 2) As Long
Dim ThisCas_Iso_As_String(1 To 2) As String
Dim ThisCas_Chem_As_String(1 To 2) As String
Dim ThisName_Iso(1 To 2) As String
Dim ThisName_Chem(1 To 2) As String
Dim ThisID_Iso(1 To 2) As Long
Dim ThisID_Chem(1 To 2) As Long
  For i = 1 To 2
    fn(i) = frmMain.txtName(i)
    Set Db(i) = OpenDatabase(fn(i))
  Next i
  '
  ' OVERALL PROCEDURE:
  ' ==================
  '
  ' Open Rs_Iso(1) sorted first by [Component Number], then by [Name].
  ' Open Rs_Chem(1) sorted first by [CAS], then by [Name].
  ' For each record in Rs_Iso(1):
  ' - Set ThisCas_Iso = Rs_Iso(1).[Component Number]
  ' - Search for ThisCas_Iso in Rs_Chem(1).[CAS]
  ' - If the search failed:
  '   - Log this record as failing the search, e.g.:
  '     - 1,"Acenaphthylene","208968"
  ' - If the search succeeded:
  '   - Log this record as a hit, e.g.:
  '     - 1,"Acenaphthylene","208968",1,"Acenaphthylene","208968"
  '
  '
  ' OPEN LOG FILE.
  '
  f = FreeFile
  Open frmMain.txtName(0).Text & "\log.txt" For Append As #f
  Print #f,
  Print #f,
  Print #f,
  Call LogOutput(f, "Started Scan_Isotherms_Table().")
  '
  ' OPEN RECORDSETS.
  '
  Set Rs_Iso(1) = Db(1).OpenRecordset( _
      "select * from [Isotherms] order by " & _
      "[Component Number], [Name]")
  Set Rs_Chem(1) = Db(1).OpenRecordset( _
      "select * from [Chemicals] order by " & _
      "[CAS], [Name]")
  '
  ' MAIN SEARCH.
  '
  Rs_Iso(1).MoveFirst
  Rs_Iso(1).MoveLast
  Rs_Iso(1).MoveFirst
  Do While Not Rs_Iso(1).EOF
    ThisCas_Iso(1) = Database_Get_Long(Rs_Iso(1), "Component Number")
    ThisCas_Iso_As_String(1) = Trim$(Str$(ThisCas_Iso(1)))
    ThisName_Iso(1) = Database_Get_String(Rs_Iso(1), "Name")
    ThisID_Iso(1) = Database_Get_Long(Rs_Iso(1), "ID")
    If (False = Database_TestForExistingLong( _
        Db(1), _
        Rs_Chem(1), _
        "Chemicals", _
        "CAS", _
        ThisCas_Iso(1))) Then
      ' SEARCH FAILED.
      Call LogOutput(f, _
          "`FAILED`," & _
          Trim$(Str$(ThisID_Iso(1))) & "," & _
          "`" & ThisName_Iso(1) & "`," & _
          "`" & ThisCas_Iso_As_String(1) & "`")
    Else
      ' SEARCH SUCCEEDED.
      ThisCas_Chem(1) = Database_Get_Long(Rs_Iso(1), "Component Number")
      ThisCas_Chem_As_String(1) = Trim$(Str$(ThisCas_Iso(1)))
      ThisName_Chem(1) = Database_Get_String(Rs_Iso(1), "Name")
      ThisID_Chem(1) = Database_Get_Long(Rs_Iso(1), "ID")
      Call LogOutput(f, _
          "`Succeeded`," & _
          Trim$(Str$(ThisID_Iso(1))) & "," & _
          "`" & ThisName_Iso(1) & "`," & _
          "`" & ThisCas_Iso_As_String(1) & "`," & _
          Trim$(Str$(ThisID_Chem(1))) & "," & _
          "`" & ThisName_Chem(1) & "`," & _
          "`" & ThisCas_Chem_As_String(1) & "`")
    End If
    Rs_Iso(1).MoveNext
  Loop
  '
  ' CLOSE DATABASES.
  '
  For i = 1 To 2
    Db(i).Close
  Next i
  '
  ' CLOSE LOG FILE.
  '
  Call LogOutput(f, "Ended Scan_Isotherms_Table().")
  Close #f
exit_normally_ThisFunc:
  Scan_Isotherms_Table = True
  Exit Function
exit_err_ThisFunc:
  Scan_Isotherms_Table = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Scan_Isotherms_Table")
  Resume exit_err_ThisFunc
End Function


Function Convert_Isotherms_Table() _
    As Boolean
On Error GoTo err_ThisFunc
Dim fn(1 To 2) As String
Dim Db(1 To 2) As Database
Dim Rs_Iso(1 To 2) As Recordset
Dim Rs_Chem(1 To 2) As Recordset
Dim i As Integer
Dim f As Integer
Dim ThisCas_Iso(1 To 2) As Long
Dim ThisCas_Chem(1 To 2) As Long
Dim ThisCas_Iso_As_String(1 To 2) As String
Dim ThisCas_Chem_As_String(1 To 2) As String
Dim ThisName_Iso(1 To 2) As String
Dim ThisName_Chem(1 To 2) As String
Dim ThisID_Iso(1 To 2) As Long
Dim ThisID_Chem(1 To 2) As Long
Dim WasFound As Boolean
  '
  ' COPY DB#1 TO DB#2.
  '
  For i = 1 To 2
    fn(i) = frmMain.txtName(i)
  Next i
  FileCopy fn(1), fn(2)
  '
  ' OPEN DATABASES.
  '
  For i = 1 To 2
    Set Db(i) = OpenDatabase(fn(i))
  Next i
  '
  ' OVERALL PROCEDURE:
  ' ==================
  '
  ' - Load in the entire change file (File #3) into an array:
  '   - Read in the initial dummy line
  '   - For each line:
  '     - Read in .CAS
  '     - Read in .Name_Old
  '     - Read in .Name_New
  ' - Open Rs_Iso(2) sorted first by [Component Number], then by [Name].
  ' - For each record in Rs_Iso(2):
  '   - Set ThisCas_Iso = Rs_Iso(2).[Component Number]
  '   - Set ThisName_Iso = Rs_Iso(2).[Name]
  '   - If ThisCas_Iso and ThisName_Iso is found in the array, then:
  '     - Edit the record
  '     - Set Rs_Iso(2).[Name] = .Name_New
  '     - Log the change
  '     - Update the record
  ' - Open Rs_Chem(2) sorted first by [CAS], then by [Name].
  ' - For each record in Rs_Chem(2):
  '   - Set ThisCas_Chem = Rs_Chem(2).[CAS]
  '   - Set ThisName_Chem = Rs_Chem(2).[Name]
  '   - If ThisCas_Chem and ThisName_Chem is found in the array, then:
  '     - Edit the record
  '     - Set Rs_Chem(2).[Name] = .Name_New
  '     - Log the change
  '     - Update the record
  '
  ' OPEN LOG FILE.
  '
  f = FreeFile
  Open frmMain.txtName(0).Text & "\log.txt" For Append As #f
  Print #f,
  Print #f,
  Print #f,
  Call LogOutput(f, "Started Convert_Isotherms_Table().")
  '
  ' LOAD THE MERGE DATA.
  '
Dim UB As Integer
Dim MergeData() As MergeData_Type
Dim f2 As Integer
Dim DummyStr1 As String
Dim s(1 To 3) As String
  UB = 0
  ReDim MergeData(0 To 0)
  f2 = FreeFile
  Open frmMain.txtName(3).Text For Input As #f2
  Line Input #f2, DummyStr1
  Do While (1 = 1)
    Input #f2, s(1), s(2), s(3)
    If (s(1) = "EOF") Then Exit Do
    UB = UB + 1
    If (UB = 1) Then
      ReDim MergeData(1 To 1)
    Else
      ReDim Preserve MergeData(1 To UB)
    End If
    With MergeData(UB)
      .CAS = s(1)
      .Name_Old = s(2)
      .Name_New = s(3)
    End With
  Loop
  Close #f2
  '
  ' OPEN RECORDSETS.
  '
  Set Rs_Iso(2) = Db(2).OpenRecordset( _
      "select * from [Isotherms] order by " & _
      "[Component Number], [Name]")
  Set Rs_Chem(2) = Db(2).OpenRecordset( _
      "select * from [Chemicals] order by " & _
      "[CAS], [Name]")
  '
  ' MAIN SEARCH-AND-REPLACE OF [Isotherms].
  '
  Rs_Iso(2).MoveFirst
  Rs_Iso(2).MoveLast
  Rs_Iso(2).MoveFirst
  Do While Not Rs_Iso(2).EOF
    ThisCas_Iso(2) = Database_Get_Long(Rs_Iso(2), "Component Number")
    ThisCas_Iso_As_String(2) = Trim$(Str$(ThisCas_Iso(2)))
    ThisName_Iso(2) = Database_Get_String(Rs_Iso(2), "Name")
    ThisID_Iso(2) = Database_Get_Long(Rs_Iso(2), "ID")
    ' SEARCH FOR THIS NAME-CAS COMBINATION.
    WasFound = False
    For i = 1 To UB
      If (Trim$(MergeData(i).CAS) = Trim$(ThisCas_Iso_As_String(2))) Then
        If (Trim$(MergeData(i).Name_Old) = Trim$(ThisName_Iso(2))) Then
          WasFound = True
          Exit For
        End If
      End If
    Next i
    If (WasFound = False) Then
      ' THIS RECORD IS OKAY; SKIP IT.
    Else
      ' THIS RECORD REQUIRES CHANGING; MODIFY IT.
      Rs_Iso(2).Edit
      Rs_Iso(2).Fields("Name") = MergeData(i).Name_New
      Call LogOutput(f, _
          "[Isotherms]: Changed name of ID#" & Trim$(Str$(ThisID_Iso(2))) & _
          " from `" & ThisName_Iso(2) & "` to " & _
          "`" & MergeData(i).Name_New & "`")
      Rs_Iso(2).Update
    End If
    Rs_Iso(2).MoveNext
  Loop
  '
  ' MAIN SEARCH-AND-REPLACE OF [Chemicals].
  '
  Rs_Chem(2).MoveFirst
  Rs_Chem(2).MoveLast
  Rs_Chem(2).MoveFirst
  Do While Not Rs_Chem(2).EOF
    ThisCas_Chem(2) = Database_Get_Long(Rs_Chem(2), "CAS")
    ThisCas_Chem_As_String(2) = Trim$(Str$(ThisCas_Chem(2)))
    ThisName_Chem(2) = Database_Get_String(Rs_Chem(2), "Name")
    ThisID_Chem(2) = Database_Get_Long(Rs_Chem(2), "Compo ID")
    ' SEARCH FOR THIS NAME-CAS COMBINATION.
    WasFound = False
    For i = 1 To UB
      If (Trim$(MergeData(i).CAS) = Trim$(ThisCas_Chem_As_String(2))) Then
        If (Trim$(MergeData(i).Name_Old) = Trim$(ThisName_Chem(2))) Then
          WasFound = True
          Exit For
        End If
      End If
    Next i
    If (WasFound = False) Then
      ' THIS RECORD IS OKAY; SKIP IT.
    Else
      ' THIS RECORD REQUIRES CHANGING; MODIFY IT.
''''      Rs_Chem(2).Delete
''''      Call LogOutput(f, _
''''          "[Chemicals]: Deleted ID#" & Trim$(Str$(ThisID_Chem(2))) & _
''''          " (`" & ThisName_Chem(2) & "`)")
      Rs_Chem(2).Edit
      Rs_Chem(2).Fields("Name") = MergeData(i).Name_New
      Call LogOutput(f, _
          "[Chemicals]: Changed name of ID#" & Trim$(Str$(ThisID_Chem(2))) & _
          " from `" & ThisName_Chem(2) & "` to " & _
          "`" & MergeData(i).Name_New & "`")
      Rs_Chem(2).Update
    End If
    Rs_Chem(2).MoveNext
  Loop
  '
  ' CLOSE DATABASES.
  '
  For i = 1 To 2
    Db(i).Close
  Next i
  '
  ' CLOSE LOG FILE.
  '
  Call LogOutput(f, "Ended Convert_Isotherms_Table().")
  Close #f
exit_normally_ThisFunc:
  Convert_Isotherms_Table = True
  Exit Function
exit_err_ThisFunc:
  Convert_Isotherms_Table = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Convert_Isotherms_Table")
  Resume exit_err_ThisFunc
End Function


