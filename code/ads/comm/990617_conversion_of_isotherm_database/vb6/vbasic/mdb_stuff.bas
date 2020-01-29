Attribute VB_Name = "MDB_Stuff"
Option Explicit

Global Ws1 As Workspace
Global fn_DB_dir As String

Global fn_DB_Fuels As String




Const MDB_Stuff_declarations_end = True


Function Database_Get_String(Rs As Recordset, fn As String) As String
Dim v As Variant
  v = Rs(fn)
  If (IsNull(v)) Then
    Database_Get_String = ""
  Else
    Database_Get_String = Trim$(v)
  End If
End Function
Function Database_Get_Double(Rs As Recordset, fn As String) As Double
Dim v As Variant
  v = Rs(fn)
  If (IsNull(v)) Then
    Database_Get_Double = 0#
  Else
    Database_Get_Double = CDbl(v)
  End If
End Function
Function Database_Get_Integer(Rs As Recordset, fn As String) As Integer
Dim v As Variant
  v = Rs(fn)
  If (IsNull(v)) Then
    Database_Get_Integer = 0
  Else
    Database_Get_Integer = CInt(v)
  End If
End Function
Function Database_Get_Long(Rs As Recordset, fn As String) As Long
Dim v As Variant
  v = Rs(fn)
  If (IsNull(v)) Then
    Database_Get_Long = 0
  Else
    Database_Get_Long = CLng(v)
  End If
End Function


Function GetNewPrimaryKey0( _
    db_SourceData As Database, _
    tn As String, _
    fn As String) As Long
Dim criteria As String
Dim Rs1 As Recordset
Dim keys() As Long
Dim i As Integer
Dim j As Long
Dim keys_count As Integer
Dim found As Boolean
  '
  ' LOAD ALL PRIMARY KEYS INTO MEMORY.
  '
  criteria = "select " & fn & " from [" & tn & "]"
  Set Rs1 = _
      db_SourceData.OpenRecordset(criteria)
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  keys_count = Rs1.RecordCount
  ReDim keys(1 To keys_count)
  i = 0
  Do Until Rs1.EOF
    i = i + 1
    keys(i) = Database_Get_Long(Rs1, fn)
    Rs1.MoveNext
  Loop
  Rs1.Close
  j = 1
  Do While (1 = 1)
    found = False
    For i = 1 To keys_count
      If (keys(i) = j) Then
        found = True
        Exit For
      End If
    Next i
    If (Not found) Then
      Exit Do
    End If
    j = j + 1
  Loop
  GetNewPrimaryKey0 = j
End Function


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


Function Database_IsTableExist( _
    Db1 As Database, _
    TableName As String) As Boolean
Dim Rs1 As Recordset
  On Error GoTo err_Database_IsTableExist
  Set Rs1 = Db1.OpenRecordset(TableName)
  Database_IsTableExist = True
  Rs1.Close
exit_err_Database_IsTableExist:
  Exit Function
err_Database_IsTableExist:
  Database_IsTableExist = False
  Resume exit_err_Database_IsTableExist
End Function


Function Database_NoRecordsInRecordset( _
    Rs1 As Recordset) As Boolean
  On Error GoTo err_Database_NoRecordsInRecordset
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  Database_NoRecordsInRecordset = False
  Exit Function
exit_err_Database_NoRecordsInRecordset:
  Database_NoRecordsInRecordset = True
  Exit Function
err_Database_NoRecordsInRecordset:
  'ERROR OCCURRED, THEREFORE THERE MUST NOT BE ANY
  'VALID RECORDS WITHIN THIS RECORDSET.
  Resume exit_err_Database_NoRecordsInRecordset
End Function


Function Database_TestForCriteria( _
    Db1 As Database, _
    Rs1 As Recordset, _
    in_Criteria As String) _
    As Boolean
    'in_TableName As String, _
    in_FieldName As String, _
    in_TestForLong As Long) _
    As Boolean
Dim Current_Criteria As String
Dim NumRecords As Integer
  On Error GoTo err_ThisSub
  Current_Criteria = in_Criteria
  ''''Current_Criteria = _
      "select * from [" & in_TableName & "] " & _
      "where " & in_FieldName & "=" & Trim$(Str$(in_TestForLong))
  Set Rs1 = _
      Db1.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_ThisSub
  If (NumRecords = 0) Then
    Database_TestForCriteria = False
    Exit Function
  Else
    Database_TestForCriteria = True
    Exit Function
  End If
  Rs1.Close
exit_normally_ThisSub:
  Database_TestForCriteria = True
  Exit Function
exit_err_ThisSub:
  Database_TestForCriteria = False
  Exit Function
err_ThisSub:
  ''''Call Show_Trapped_Error("Database_TestForCriteria")
  Resume exit_err_ThisSub
End Function


Function Database_TestForExistingLong( _
    Db1 As Database, _
    Rs1 As Recordset, _
    in_TableName As String, _
    in_FieldName As String, _
    in_TestForLong As Long) _
    As Boolean
Dim Current_Criteria As String
Dim NumRecords As Integer
  On Error GoTo err_ThisSub
  Current_Criteria = _
      "select * from [" & in_TableName & "] " & _
      "where " & in_FieldName & "=" & Trim$(Str$(in_TestForLong))
  Set Rs1 = _
      Db1.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_ThisSub
  If (NumRecords = 0) Then
    Database_TestForExistingLong = False
    Exit Function
  Else
    Database_TestForExistingLong = True
    Exit Function
  End If
  Rs1.Close
exit_normally_ThisSub:
  Database_TestForExistingLong = True
  Exit Function
exit_err_ThisSub:
  Database_TestForExistingLong = False
  Exit Function
err_ThisSub:
  ''''Call Show_Trapped_Error("Database_TestForExistingLong")
  Resume exit_err_ThisSub
End Function


Function Database_TestForExistingString00( _
    Db1 As Database, _
    Rs1 As Recordset, _
    in_TableName As String, _
    in_FieldName As String, _
    in_SearchCriteria As String) _
    As Boolean
Dim Current_Criteria As String
Dim NumRecords As Integer
  On Error GoTo err_ThisSub
  Current_Criteria = in_SearchCriteria
  ' _
      "select * from [" & in_TableName & "] " & _
      "where " & in_FieldName & "='" & in_TestForStr & "'"
  Set Rs1 = _
      Db1.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_ThisSub
  If (NumRecords = 0) Then
    Database_TestForExistingString00 = False
    Exit Function
  Else
    Database_TestForExistingString00 = True
    Exit Function
  End If
  Rs1.Close
exit_normally_ThisSub:
  Database_TestForExistingString00 = True
  Exit Function
exit_err_ThisSub:
  Database_TestForExistingString00 = False
  Exit Function
err_ThisSub:
  ''''Call Show_Trapped_Error("Database_TestForExistingString")
  Resume exit_err_ThisSub
End Function


Function Database_TestForExistingString0( _
    Db1 As Database, _
    Rs1 As Recordset, _
    in_TableName As String, _
    in_FieldName As String, _
    in_TestForStr As String) _
    As Boolean
Dim Current_Criteria As String
Dim NumRecords As Integer
  On Error GoTo err_ThisSub
  Current_Criteria = _
      "select * from [" & in_TableName & "] " & _
      "where " & in_FieldName & "='" & in_TestForStr & "'"
  Set Rs1 = _
      Db1.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_ThisSub
  If (NumRecords = 0) Then
    Database_TestForExistingString0 = False
    Exit Function
  Else
    Database_TestForExistingString0 = True
    Exit Function
  End If
  Rs1.Close
exit_normally_ThisSub:
  Database_TestForExistingString0 = True
  Exit Function
exit_err_ThisSub:
  Database_TestForExistingString0 = False
  Exit Function
err_ThisSub:
  ''''Call Show_Trapped_Error("Database_TestForExistingString")
  Resume exit_err_ThisSub
End Function


Function Database_TestForExistingString( _
    Db1 As Database, _
    in_TableName As String, _
    in_FieldName As String, _
    in_TestForStr As String) _
    As Boolean
Dim Current_Criteria As String
Dim Rs1 As Recordset
Dim NumRecords As Integer
  On Error GoTo err_ThisSub
  Current_Criteria = _
      "select * from [" & in_TableName & "] " & _
      "where " & in_FieldName & "='" & in_TestForStr & "'"
  Set Rs1 = _
      Db1.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_ThisSub
  If (NumRecords = 0) Then
    Database_TestForExistingString = False
    Exit Function
  Else
    Database_TestForExistingString = True
    Exit Function
  End If
  Rs1.Close
exit_normally_ThisSub:
  Database_TestForExistingString = True
  Exit Function
exit_err_ThisSub:
  Database_TestForExistingString = False
  Exit Function
err_ThisSub:
  ''''Call Show_Trapped_Error("Database_TestForExistingString")
  Resume exit_err_ThisSub
End Function


