Attribute VB_Name = "MDB_Stuff"
Option Explicit





Const MDB_Stuff_declarations_end = True


Function Database_Get_String(rs As Recordset, fn As String) As String
Dim v As Variant
  v = rs(fn)
  If (IsNull(v)) Then
    Database_Get_String = ""
  Else
    Database_Get_String = Trim$(v)
  End If
End Function
Function Database_Get_Double(rs As Recordset, fn As String) As Double
Dim v As Variant
  v = rs(fn)
  If (IsNull(v)) Then
    Database_Get_Double = 0#
  Else
    Database_Get_Double = CDbl(v)
  End If
End Function
Function Database_Get_Integer(rs As Recordset, fn As String) As Integer
Dim v As Variant
  v = rs(fn)
  If (IsNull(v)) Then
    Database_Get_Integer = 0
  Else
    Database_Get_Integer = CInt(v)
  End If
End Function
Function Database_Get_Long(rs As Recordset, fn As String) As Long
Dim v As Variant
  v = rs(fn)
  If (IsNull(v)) Then
    Database_Get_Long = 0
  Else
    Database_Get_Long = CLng(v)
  End If
End Function


Function GetNewPrimaryKey0( _
    db_SourceData As database, _
    tn As String, _
    fn As String) As Long
Dim criteria As String
Dim RS1 As Recordset
Dim keys() As Long
Dim i As Integer
Dim J As Long
Dim keys_count As Integer
Dim Found As Boolean

  'LOAD ALL PRIMARY KEYS INTO MEMORY.
  criteria = "select " & fn & " from [" & tn & "]"
  Set RS1 = _
      db_SourceData.OpenRecordset(criteria)
  RS1.MoveFirst
  RS1.MoveLast
  RS1.MoveFirst
  keys_count = RS1.RecordCount
  ReDim keys(1 To keys_count)
  i = 0
  Do Until RS1.EOF
    i = i + 1
    keys(i) = Database_Get_Long(RS1, fn)
    RS1.MoveNext
  Loop
  RS1.Close
  J = 1
  Do While (1 = 1)
    Found = False
    For i = 1 To keys_count
      If (keys(i) = J) Then
        Found = True
        Exit For
      End If
    Next i
    If (Not Found) Then
      Exit Do
    End If
    J = J + 1
  Loop
  GetNewPrimaryKey0 = J
End Function


Function Database_IsTableExist( _
    Db1 As database, _
    TableName As String) As Boolean
Dim RS1 As Recordset
  On Error GoTo err_Database_IsTableExist
  Set RS1 = Db1.OpenRecordset(TableName)
  Database_IsTableExist = True
  RS1.Close
exit_err_Database_IsTableExist:
  Exit Function
err_Database_IsTableExist:
  Database_IsTableExist = False
  Resume exit_err_Database_IsTableExist
End Function


Function Database_NoRecordsInRecordset( _
    RS1 As Recordset) As Boolean
  On Error GoTo err_Database_NoRecordsInRecordset
  RS1.MoveFirst
  RS1.MoveLast
  RS1.MoveFirst
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



