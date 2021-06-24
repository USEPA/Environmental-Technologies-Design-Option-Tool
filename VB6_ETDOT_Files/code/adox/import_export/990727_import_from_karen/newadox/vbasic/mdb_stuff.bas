Attribute VB_Name = "MDB_Stuff"
Option Explicit

Global Ws1 As Workspace
Global fn_DB_dir As String





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
    db_SourceData As Database, _
    tn As String, _
    fn As String) As Long
Dim criteria As String
Dim Rs1 As Recordset
Dim keys() As Long
Dim i As Integer
Dim j As Long
Dim keys_count As Integer
Dim Found As Boolean
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
    Found = False
    For i = 1 To keys_count
      If (keys(i) = j) Then
        Found = True
        Exit For
      End If
    Next i
    If (Not Found) Then
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



