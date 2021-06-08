Option Strict Off
Option Explicit On
Module MDB_Stuff
	
	'UPGRADE_ISSUE: Workspace object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
    Public Ws1 As dao.Workspace
    Public DAOEngine As dao.DBEngine
    Dim rst As dao.Recordset
    Dim ws As dao.Workspace
    Public fn_DB_dir As String
    Public fn_DB_Isotherm As String
	Public fn_DB_Carbon As String
	
	'
	'DATABASE PASSWORD LISTED BELOW.  AS A DIFFERENCE FROM
	'THE WAY AdXDesignS WORKED UNDER MS VB 3.0,
	'NO USER NAME IS REQUIRED, ONLY A PASSWORD.
	'
	'NOTE: FOR SOME #$(*&@#$*# ANNOYING REASON, ONCE A
	'PASSWORD IS SUPPLIED FOR A DATABASE TO BE ACCESSED
	'FROM VISUAL BASIC 5 (IN THE WAY THIS PROGRAM ACCESSES
	'THE DATABASES), IT CANNOT BE DIRECTLY ACCESS FROM
	'MICROSOFT ACCESS USING THE SAME PASSWORD IF THE PASSWORD
	'WAS 14 CHARACTERS LONG (OR MAYBE JUST THIS EXAMPLE).
	'USING THE NEW 9 CHARACTER PASSWORD WORKS EQUALLY
	'WELL IN VB 5.0 AND MS ACCESS 8.0.
	'              Eric J. Oman
	'              9/21/98
	'
	''Global Const User_Password = "frieda4wisc836"
	'Global Const Encrypted_User_Password = "AeJ>;2atJh8m^g"
	'Global Const User_Password = "frieda836"
	Public Const Encrypted_User_Password As String = "AeJ>;2m^g"
	''Global Const User_Name = "victor t. hart"
	'Global Const Encrypted_User_Name = "qJ8k\e%kO%G2ek"
	
	
	
	
	
	
	Const MDB_Stuff_declarations_end As Boolean = True

    Function Database_Get_String(rs As dao.Recordset, fn As String) As String
        Dim v As Object
        v = rs(fn).Value
        If (IsDBNull(v)) Then
            Database_Get_String = ""
        Else
            Database_Get_String = Trim$(v)
        End If
    End Function
    Function Database_Get_Double(rs As dao.Recordset, fn As String) As Double
        Dim v As Object 'Variant
        v = rs(fn).Value
        If (IsDBNull(v)) Then
            Database_Get_Double = 0#
        Else
            Database_Get_Double = CDbl(v)
        End If
    End Function
    Function Database_Get_Integer(rs As dao.Recordset, fn As String) As Integer
        Dim v As Object 'VariantType
        v = rs(fn).Value
        If (IsDBNull(v)) Then
            Database_Get_Integer = 0
        Else
            Database_Get_Integer = CInt(v)
        End If
    End Function


    Function Database_Get_Long(ByRef rs As dao.Recordset, ByRef fn As String) As Integer
        Dim v As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object rs(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        v = rs(fn).Value
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If (IsDBNull(v)) Then
            Database_Get_Long = 0
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Database_Get_Long = CInt(v)
        End If
    End Function
	
	
    Function GetNewPrimaryKey0(ByRef db_SourceData As dao.Database, ByRef tn As String, ByRef fn As String) As Integer
        Dim criteria As String
        'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        Dim Rs1 As dao.Recordset
        Dim keys() As Integer
        Dim i As Short
        Dim j As Integer
        Dim keys_count As Short
        Dim Found As Boolean

        'LOAD ALL PRIMARY KEYS INTO MEMORY.
        criteria = "select " & fn & " from [" & tn & "]"
        'UPGRADE_WARNING: Couldn't resolve default property of object db_SourceData.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rs1 = db_SourceData.OpenRecordset(criteria)
        'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rs1.MoveFirst()
        'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rs1.MoveLast()
        'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rs1.MoveFirst()
        'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.RecordCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        keys_count = Rs1.RecordCount
        'UPGRADE_WARNING: Lower bound of array keys was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        ReDim keys(keys_count)
        i = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do Until Rs1.EOF
            i = i + 1
            keys(i) = Database_Get_Long(Rs1, fn)
            'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Rs1.MoveNext()
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rs1.Close()
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
	
	
	Function decrypt_string(ByRef password As String) As String
		Dim newpass As String
		Dim i, keyval, length As Short
		Dim Key(100) As Short
		For i = 32 To 122
			keyval = ((i * 3) Mod 91)
			Key(keyval) = i
		Next i
		length = Len(password)
		newpass = ""
		For i = 1 To length
			newpass = newpass & Chr(Key(Asc(Mid(password, i, 1)) - 32))
		Next i
		decrypt_string = newpass
	End Function
	
	


End Module