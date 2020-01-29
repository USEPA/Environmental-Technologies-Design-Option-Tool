Attribute VB_Name = "MODExtract"
Option Explicit
Global DBTXTFilePath As String
Global DBTXTFileRefPath As String
Global TemplatePath As String
'Global DBDestPath As String ' not sure we're still using this
Global files_found(2) As String
Global file_needed As String
Global PathMaster As String
Global Path801 As String
Global SourcePath As String
Global DestPath As String
Global Path911 As String
Global Path801Access As String
Global Const NULLPATH = "e"

Sub FindTemplate()

   Dim Response As Integer
        
    FRMMain.SSPanel1.Value = 0
        
   Screen.MousePointer = 11
   If FRMMain.Option1.Value = True Then
            Response = MsgBox("This process takes approximately 30 minutes.  Extract using the specified paths?", 4, "Begin Extraction")
            If Response = 6 Then
                DBTXTFilePath = SourcePath + "\src"
                DBTXTFileRefPath = SourcePath + "\refs"
                Call Extract801
                FRMMain!Label1.Caption = ""
                FRMMain.Refresh
            Else
                Screen.MousePointer = 1
                Exit Sub
            End If
    ElseIf FRMMain.Option2.Value = True Then
            Response = MsgBox("This process takes approximately 30 minutes.  Extract using the specified paths?", 4, "Begin Extraction")
            If Response = 6 Then
                Call Extract801Access
                FRMMain!Label1.Caption = ""
                FRMMain.Refresh
            Else
                Screen.MousePointer = 1
                Exit Sub
            End If
    ElseIf FRMMain.Option3.Value = True Then
            Response = MsgBox("This process takes approximately 30 minutes.  Extract using the specified paths?", 4, "Begin Extraction")
            If Response = 6 Then
                Call Extract911
                FRMMain!Label1.Caption = ""
                FRMMain.Refresh
            Else
                Screen.MousePointer = 1
                Exit Sub
            End If
   End If
   
   FRMMain!SSPanel1.Value = 100
   Beep
   MsgBox "Extraction Complete", 0, "Extraction Done"
   Screen.MousePointer = 1
   
End Sub

Public Sub ExtractRefs()
   Dim db As Database
   Dim DBTbl As Recordset
   Dim DBDef As TableDef
   Dim PrimaryIndex As Index
   Dim Field1 As Field
   Dim DNUM As String
   Dim LDNUM As String
   Dim thistext As String
   Dim text As String
   Dim linesdone As Long
   Dim size As Long
   Dim i As Integer
   Dim file As Integer
   Dim filenum As String
   Dim Start As Integer
   Dim counter1 As Long
   Dim lastpos As Integer
    
   linesdone = 0
   Call GetFileSize(size, i, 2)
   
   On Error Resume Next
   
   Set db = OpenDatabase(DestPath & "\master.mdb", False, False)
   
   Set DBDef = db.TableDefs("801REFS")
   Set PrimaryIndex = DBDef.CreateIndex("PrimaryKey")
   PrimaryIndex.Primary = False
   PrimaryIndex.Unique = False
   Set Field1 = PrimaryIndex.CreateField("REFNUM")
   PrimaryIndex.Fields.Append Field1
   DBDef.Indexes.Append PrimaryIndex
      
   Set DBTbl = db.OpenRecordset("801refs", dbOpenTable)
      
    FRMMain!Label1.Caption = "Please wait...Initializing"
    FRMMain!Label1.Refresh
   
   'Delete current records in table
   'If the table is empty, skip this step
    If Not DBTbl.EOF Then
        DBTbl.MoveFirst
        Do While Not DBTbl.EOF
            DBTbl.Delete
            DBTbl.MoveNext
        Loop
    End If
 
   file = 1

   filenum = Trim$(Str$(file))
   If file < 10 Then filenum = "0" + Trim$(Str$(file))
   Open DBTXTFileRefPath & filenum + ".dat" For Input As #1

   Start = True
    FRMMain!Label1.Caption = "Extracting 801 References..."
    FRMMain.Refresh
ReadNext:

   If EOF(1) = True Then
      
      DBTbl("REFNUM") = LDNUM
      DBTbl("REFERENCE") = thistext
      DBTbl.Update
      thistext = ""
      Close #1
      
      file = file + 1
      'check for end of data files
      If file > i - 1 Then
         DBTbl.Close
         'Turn on DIPPR801 switch
         Set db = OpenDatabase(DestPath & "\dbsave.mdb", False, False)
         Set DBTbl = db.OpenRecordset("PrefDatabases", dbOpenTable)
         DBTbl.MoveFirst
         DBTbl.Edit
         DBTbl("Available") = True
         DBTbl.Update
         DBTbl.Close
         FRMMain!Label1.Caption = ""
         'Close #2
         Exit Sub
      End If



      filenum = Trim$(Str$(file))
      If file < 10 Then filenum = "0" + Trim$(Str$(file))
      Open DBTXTFileRefPath & filenum + ".dat" For Input As #1
      Start = True
   End If

    Line Input #1, text

    linesdone = linesdone + 1
    counter1 = counter1 + 1
   'guestimate how much is done
   If Int((linesdone / size) * 8000) <= 100 Then
       FRMMain!SSPanel1.Value = Int((linesdone / size) * 8000)
   End If
    
   DNUM = Mid$(text, 1, 5)
If Trim(DNUM) = CStr(4999) Then
    counter1 = counter1 + 1
End If
    
   If Start = True Then
      LDNUM = DNUM
      Start = False
      DBTbl.AddNew
   
   End If

   If LDNUM <> DNUM Then
      DBTbl("REFNUM") = LDNUM
      DBTbl("REFERENCE") = thistext
      LDNUM = DNUM
      DBTbl.Update
      DBTbl.AddNew
      thistext = ""
   End If

Check:

    lastpos = Len(RTrim(text))
    If Mid(text, lastpos, 1) = "-" Then
        lastpos = lastpos - 1
    End If
    
    thistext = Trim(thistext & Mid$(text, 6, lastpos))
    
    GoTo ReadNext:
    

End Sub

Sub Extract911()
    Dim db As Database
    Dim db2 As Database
    Dim db3 As Database
    Dim DBTbl As Recordset
    Dim DBTbl2 As Recordset
    Dim DBTbl3 As Recordset
    Dim CNum As Double
    Dim CAS As Double
    Dim PropCode As String
    Dim DBDef As TableDef
    Dim PrimaryIndex As Index
    Dim Field1 As Field
    Dim Field2 As Field
    Dim Counter As Double
    Dim counter1 As Double
    Dim tempTemp As String
    Dim TempUnits As String
    Dim x, i As Integer
    Dim tempchar As String
    Dim tempstring As String
    
    counter1 = 0
    
    'On Error GoTo error_handler

    Counter = 0
    
    Set db = OpenDatabase(DestPath & "\master.mdb", False, False)
    
    Set DBTbl = db.OpenRecordset("DIPPR911", dbOpenTable)
        
    Set db2 = OpenDatabase(SourcePath, False, True, "Paradox 4.x;")
    Set DBTbl2 = db2.OpenRecordset("valuesf", dbOpenTable)
    Set DBTbl3 = db2.OpenRecordset("chemical", dbOpenSnapshot)
     
    'Delete current records in table
    FRMMain!Label1.Caption = "Please wait...Initializing"
    FRMMain!Label1.Refresh
    
    'If the table is empty, skip this step
    If Not DBTbl.EOF Then
        DBTbl.MoveFirst
        Do While Not DBTbl.EOF
            DBTbl.Delete
            DBTbl.MoveNext
        Loop
    End If
    
    FRMMain!Label1.Caption = "Extracting Data..."
    FRMMain.Refresh
    DBTbl2.MoveFirst
    Do While Not DBTbl2.EOF
        On Error Resume Next
        CNum = DBTbl2("C #")
        PropCode = DBTbl2("Property Code")
        DBTbl3.FindFirst "[C#] =" & CNum
        CAS = DBTbl3("CAS #")
        DBTbl.AddNew
        DBTbl("C #") = CNum
        DBTbl("CAS #") = CAS
        DBTbl("Property Code") = PropCode
        DBTbl("PEARLS Code") = GetPropCode(PropCode)
        DBTbl("Key Rating") = DBTbl2("Key Rating")
        DBTbl("Article #") = DBTbl2("Article #")
        DBTbl("Rating") = DBTbl2("Rating")
        DBTbl("Desc/Method") = DBTbl2("Desc/Method")
        DBTbl("Keywords") = DBTbl2("Keywords")
        DBTbl("Value") = DBTbl2("Value")
        DBTbl("Units") = DBTbl2("Units")
                
        ' Characters in the paradox database are not all MS
        ' compliant, so if the character is not recognized, by
        ' default we make it a "-"
        If (Len(DBTbl("Units")) <> Null) Then
            i = Len(DBTbl("Units"))
        End If
        
        For x = 1 To i
            tempchar = Right(DBTbl("Units"), x)
            If Asc(tempchar) < 32 Or (Asc(tempchar) > 126 And Asc(tempchar) < 160) Then
                tempstring = Left(DBTbl("Units"), Len(DBTbl("Units")) - Len(tempchar))
                tempstring = tempstring & "-"
                tempstring = tempstring & Right(tempchar, Len(tempchar) - 1)
                DBTbl("Units") = tempstring
            End If
        Next x
                
'        If Not (DBTbl2("Temperature") = Null) Then
            tempTemp = DBTbl2("Temperature")
            DBTbl("Temperature") = Val(tempTemp)
            DBTbl("TempUnits") = GetUnits(tempTemp)
'        End If
        
        DBTbl("Equation") = DBTbl2("Equation")
        DBTbl("Coef1") = DBTbl2("Coef1")
        DBTbl("Coef2") = DBTbl2("Coef2")
        DBTbl("Coef3") = DBTbl2("Coef3")
        DBTbl("Coef4") = DBTbl2("Coef4")
        DBTbl("Coef5") = DBTbl2("Coef5")
        DBTbl("Ref2") = DBTbl2("Ref2")
        DBTbl("Ref3") = DBTbl2("Ref3")
        DBTbl("FootN") = DBTbl2("FootN")
        DBTbl("Pressure") = DBTbl2("Pressure")
        DBTbl("Comment") = DBTbl2("Comment")
        DBTbl("Default Multiplier") = DBTbl2("Default Multiplier")
        DBTbl("Default Addend") = DBTbl2("Default Addend")
        DBTbl.Update
        'Counter = Counter + 0.007
        counter1 = counter1 + 1
        
        FRMMain.SSPanel1.Value = (DBTbl2.PercentPosition / 2)

        DBTbl2.MoveNext
        tempTemp = " "
    Loop
    On Error GoTo error_handler
    DBTbl.Close
    DBTbl2.Close
    DBTbl3.Close
    
    Set DBDef = db.TableDefs("DIPPR911")
    Set PrimaryIndex = DBDef.CreateIndex("PrimaryKey1")
    PrimaryIndex.Primary = False
    PrimaryIndex.Unique = False
    Set Field1 = PrimaryIndex.CreateField("Cas #")
    PrimaryIndex.Fields.Append Field1
    'DBDef.Indexes(1).Name = ""
        
    Set PrimaryIndex = DBDef.CreateIndex("PrimaryKey2")
    PrimaryIndex.Primary = False
    PrimaryIndex.Unique = False
    Set Field1 = PrimaryIndex.CreateField("Cas #")
    Set Field2 = PrimaryIndex.CreateField("PEARLS Code")
    PrimaryIndex.Fields.Append Field1
    PrimaryIndex.Fields.Append Field2
    'DBDef.Indexes.Append PrimaryIndex
    
    Set DBTbl = db.OpenRecordset("CITATION", dbOpenTable)
    Set DBTbl2 = db2.OpenRecordset("CITATION", dbOpenTable)
        
    'Delete current records in table
    'Skip if table is empty
    '**DBTbl is redefined from previous in this subroutine**
    If Not DBTbl.EOF Then
        DBTbl.MoveFirst
        Do While Not DBTbl.EOF
            DBTbl.Delete
            DBTbl.MoveNext
        Loop
    End If
    
    DBTbl2.MoveFirst
    Do While Not DBTbl2.EOF
        DBTbl.AddNew
        DBTbl("Paper #") = DBTbl2("Paper #")
        DBTbl("Title") = DBTbl2("Title")
        DBTbl("Author") = DBTbl2("Author")
        DBTbl("Journal") = DBTbl2("Journal")
        DBTbl("Date") = DBTbl2("Date")
        DBTbl("Volume") = DBTbl2("Volume")
        DBTbl("Number") = DBTbl2("Number")
        DBTbl("Pages") = DBTbl2("Pages")
        DBTbl("PCode1") = DBTbl2("PCode1")
        DBTbl("PCode2") = DBTbl2("PCode2")
        DBTbl("PCode3") = DBTbl2("PCode3")
        DBTbl("PCode4") = DBTbl2("PCode4")
        DBTbl("PCode5") = DBTbl2("PCode5")
        DBTbl("PCode6") = DBTbl2("PCode6")
        DBTbl("Date Assigned") = DBTbl2("Date Assigned")
        DBTbl.Update
        
        FRMMain!SSPanel1.Value = 50 + (DBTbl2.PercentPosition / 2)
        
        DBTbl2.MoveNext
    Loop
    
    DBTbl.Close
    DBTbl2.Close
    
    Set DBDef = db.TableDefs("CITATION")
    Set PrimaryIndex = DBDef.CreateIndex("PrimaryKey3")
    PrimaryIndex.Primary = False
    PrimaryIndex.Unique = False
    Set Field1 = PrimaryIndex.CreateField("Paper #")
    PrimaryIndex.Fields.Append Field1
    'DBDef.Indexes.Append PrimaryIndex
           
    'Turn on DIPPR911 switch
    Set db = OpenDatabase(DestPath & "\dbsave.mdb", False, False)
    Set DBTbl = db.OpenRecordset("PrefDatabases", dbOpenTable)
    
    DBTbl.MoveFirst
    DBTbl.MoveNext
    DBTbl.Edit
    DBTbl("Available") = True
    DBTbl.Update
    DBTbl.Close
    
    Exit Sub

error_handler:
    
    If Err = 94 Then
        DBTbl("TempUnits") = " "
        Resume Next
    Else
        MsgBox "Error #" & Err & " in DCUT", 48, " DCUT "
    End If

End Sub
Public Function GetPropCode(Code As String) As Integer

    Select Case Code
        Case "1a"
            GetPropCode = 31
        Case "1b"
            GetPropCode = 30
        Case "1cc"
            GetPropCode = 28
        Case "1cn"
            GetPropCode = 29
        Case "2a"
            GetPropCode = 35
        Case "2c"
            GetPropCode = 36
        Case "2d"
            GetPropCode = 37
        Case "3a"
            GetPropCode = 0
        Case "3b"
            GetPropCode = 1
        Case "3bt"
            GetPropCode = 2
        Case "3c"
            GetPropCode = 39
        Case "3d"
            GetPropCode = 3
        Case "3e"
            GetPropCode = 4
        Case "3f"
            GetPropCode = 5
        Case "3g"
            GetPropCode = 6
        Case "3h"
            GetPropCode = 16
        Case "3i"
            GetPropCode = 15
        Case "3j"
            GetPropCode = 19
        Case "3k"
            GetPropCode = 20
        Case "3l"
            GetPropCode = 17
        Case "3lt"
            GetPropCode = 18
        Case "3ml"
            GetPropCode = 21
        Case "3mv"
            GetPropCode = 22
        Case "3n"
            GetPropCode = 7
        Case "3o"
            GetPropCode = 8
        Case "3p"
            GetPropCode = 9
        Case "3q"
            GetPropCode = 13
        Case "3r"
            GetPropCode = 14
        Case "3s"
            GetPropCode = 38
        Case "3t"
            GetPropCode = 10
        Case "3tt"
            GetPropCode = 12
        Case "3tz"
            GetPropCode = 11
        Case "4a"
            GetPropCode = 34
        Case "4b"
            GetPropCode = 32
        Case "4c"
            GetPropCode = 33
        Case "5al"
            GetPropCode = 24
        Case "5au"
            GetPropCode = 23
        Case "5b"
            GetPropCode = 25
        Case "5c"
            GetPropCode = 26
        Case "5d"
            GetPropCode = 27
        Case "6ab"
            GetPropCode = 41
        Case "6ac"
            GetPropCode = 42
        Case "6ba"
            GetPropCode = 43
        Case "6bb"
            GetPropCode = 44
        Case "6bc"
            GetPropCode = 45
        Case "6fa"
            GetPropCode = 46
        Case "6fb"
            GetPropCode = 47
        Case "6fc"
            GetPropCode = 48
        Case "6ca"
            GetPropCode = 49
        Case "6cb"
            GetPropCode = 50
        Case "6da"
            GetPropCode = 51
        Case "6db"
            GetPropCode = 52
        Case "6hc"
            GetPropCode = 53
        Case "6i"
            GetPropCode = 54
        Case Else
            GetPropCode = -1
    End Select
    
End Function

Sub Extract801Access()
    Dim db As Database, db1 As Database, db2 As Database, db801 As Database
    Dim DBTbl As Recordset, DBTbl1 As Recordset, DBTbl2 As Recordset
    Dim Tbl801 As Recordset, Tbl801A As Recordset, Tbl801B As Recordset
    Dim Info As Variant
    Dim text As String
    Dim CurCNum As Long, NewCNum As Long
    Dim TValue As Double
    Dim TRefID As Long
    Dim Start As Integer
    Dim size As Long
    Dim i As Integer
    Dim TmpFormula As String, TmpFamily As String, TmpName As String
    Dim TmpCAS As Double
    Dim DBDef As TableDef
    Dim PrimaryIndex As Index
    Dim Field1 As Field
    Dim counter1 As Long
    Dim Counter2 As Integer
    Dim file As Integer
    Dim reltext As String
    Dim aref As String
    Dim temprate As Integer
    Dim DataT As String, SourceT As String, ErrorT As String
    Dim Rating As Integer
    Dim TEqnID As Long, TNumCoeff As Long
    Dim TA As Double, TB As Double, TC As Double, TD As Double, TE As Double
    Dim TMin As Double, TMax As Double
    Dim TDataT As String

    counter1 = 0
    Counter2 = 0
    FRMMain!Label1.Caption = "Extracting Data..."
    FRMMain.Refresh
    FRMMain!Label1.Caption = "Please wait....Initializing"
    FRMMain!Label1.Refresh
    
    On Error Resume Next
   
   Set db = OpenDatabase(DestPath & "\master.mdb", False, False)
   
    Set DBDef = db.TableDefs("DIPPR801")
   Set PrimaryIndex = DBDef.CreateIndex("PrimaryKey")
   PrimaryIndex.Primary = False
   PrimaryIndex.Unique = False
   Set Field1 = PrimaryIndex.CreateField("CAS")
   PrimaryIndex.Fields.Append Field1
'   DBDef.Indexes.Append PrimaryIndex
      
   Set DBTbl = db.OpenRecordset("DIPPR801", dbOpenTable)
   Set DBTbl1 = db.OpenRecordset("801TempDepend")
   'Delete current records in table
   'If the table is empty, skip this step
    If Not DBTbl.EOF Then
        DBTbl.MoveFirst
        Do While Not DBTbl.EOF
            DBTbl.Delete
            DBTbl.MoveNext
        Loop
    End If
    
    If Not DBTbl1.EOF Then
        DBTbl1.MoveFirst
        Do While Not DBTbl1.EOF
            DBTbl1.Delete
             DBTbl1.MoveNext
        Loop
    End If
    
    Set db801 = OpenDatabase(SourcePath & "\Dippr801db.mdb")
    
    FRMMain!Label1.Caption = "Extracting Data..."
    FRMMain.Refresh
    
    Set Tbl801 = db801.OpenRecordset( _
        "SELECT * " & _
        "FROM Chem_Info", dbOpenSnapshot)
    Tbl801.MoveLast
    Tbl801.MoveFirst
    
    size = Tbl801.RecordCount
    counter1 = 0
    
    CurCNum = -1
    
    Do While (Not Tbl801.EOF)
        NewCNum = Tbl801!ChemID
        
        If CurCNum = -1 Then
            DBTbl.AddNew
            CurCNum = NewCNum
        End If
        
        If CurCNum <> NewCNum Then
            'Get other tables data
            Set Tbl801A = db801.OpenRecordset( _
                "SELECT * FROM Const_Values " & _
                "WHERE ChemID = " & CStr(CurCNum), dbOpenSnapshot)
            Tbl801A.MoveFirst

'******************************************************************************
            Do While (Not Tbl801A.EOF)
            
              If Trim(Tbl801A!Acceptance) = "A" Then
              
                If IsNull(Tbl801A!Data_Type) Then
                    DataT = ""
                Else
                    DataT = Trim(Tbl801A!Data_Type)
                End If
                
                If IsNull(Tbl801A!Source_Type) Then
                    SourceT = ""
                Else
                    SourceT = Trim(Tbl801A!Source_Type)
                End If
                
                If IsNull(Tbl801A!Error) Then
                    ErrorT = ""
                Else
                    ErrorT = Trim(Tbl801A!Error)
                End If
                
                If IsNull(Tbl801A!Value) Then
                    TValue = 0
                Else
                    TValue = Val(Tbl801A!Value)
                End If
                
                If IsNull(Tbl801A!refid) Then
                    TRefID = 0
                Else
                    TRefID = Tbl801A!refid
                End If
                
                Rating = 10
                
                If DataT = "Experimental" Then
                    Rating = Rating + 1
                    
                    If SourceT = "Evaluated" Then
                        Rating = Rating + 1
                    End If
                    
                    If ErrorT = "< 0.2%" Then
                        Rating = Rating + 2
                    ElseIf ErrorT = "< 1%" Then
                        Rating = Rating + 2
                    ElseIf ErrorT = "< 3%" Then
                        Rating = Rating + 2
                    ElseIf ErrorT = "< 5%" Then
                        Rating = Rating + 1
                    ElseIf ErrorT = "< 10%" Then
                        Rating = Rating + 1
                    End If
                Else
                    Rating = -1
                End If
              
                If IsNull(DBTbl("MWR")) Then
                    DBTbl("MW") = TValue
                    If Rating > -1 Then
                        DBTbl("MWR") = Rating
                    End If
                    DBTbl("MWREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "MW" And DBTbl("MWR") < Rating Then
                    DBTbl("MW") = TValue
                    DBTbl("MWR") = Rating
                    DBTbl("MWREF") = TRefID
                End If

                If IsNull(DBTbl("TCR")) Then
                    DBTbl("TC") = TValue
                    If Rating > -1 Then
                        DBTbl("TCR") = Rating
                    End If
                    DBTbl("TCREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "TC" And DBTbl("TCR") < Rating Then
                    DBTbl("TC") = TValue
                    DBTbl("TCR") = Rating
                    DBTbl("TCREF") = TRefID
                End If

                If IsNull(DBTbl("PCR")) Then
                    DBTbl("PC") = TValue
                    If Rating > -1 Then
                        DBTbl("PCR") = Rating
                    End If
                    DBTbl("PCREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "PC" And DBTbl("PCR") < Rating Then
                    DBTbl("PC") = TValue
                    DBTbl("PCR") = Rating
                    DBTbl("PCREF") = TRefID
                End If

                If IsNull(DBTbl("VCR")) Then
                    DBTbl("VC") = TValue
                    If Rating > -1 Then
                        DBTbl("VCR") = Rating
                    End If
                    DBTbl("VCREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "VC" And DBTbl("VCR") < Rating Then
                    DBTbl("VC") = TValue
                    DBTbl("VCR") = Rating
                    DBTbl("VCREF") = TRefID
                End If

                If IsNull(DBTbl("ZCR")) Then
                    DBTbl("ZC") = TValue
                    If Rating > -1 Then
                        DBTbl("ZCR") = Rating
                    End If
                    DBTbl("ZCREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "ZC" And DBTbl("ZCR") < Rating Then
                    DBTbl("ZC") = TValue
                    DBTbl("ZCR") = Rating
                    DBTbl("ZCREF") = TRefID
                End If

                If IsNull(DBTbl("MPR")) Then
                    DBTbl("MP") = TValue
                    If Rating > -1 Then
                        DBTbl("MPR") = Rating
                    End If
                    DBTbl("MPREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "MP" And DBTbl("MPR") < Rating Then
                    DBTbl("MP") = TValue
                    DBTbl("MPR") = Rating
                    DBTbl("MPREF") = TRefID
                End If
                
                If IsNull(DBTbl("TPTR")) Then
                    DBTbl("TPT") = TValue
                    If Rating > -1 Then
                        DBTbl("TPTR") = Rating
                    End If
                    DBTbl("TPTREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "TPT" And DBTbl("TPTR") < Rating Then
                    DBTbl("TPT") = TValue
                    DBTbl("TPTR") = Rating
                    DBTbl("TPTREF") = TRefID
                End If

                If IsNull(DBTbl("TPPR")) Then
                    DBTbl("TPP") = TValue
                    If Rating > -1 Then
                        DBTbl("TPPR") = Rating
                    End If
                    DBTbl("TPPREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "TPP" And DBTbl("TPPR") < Rating Then
                    DBTbl("TPP") = TValue
                    DBTbl("TPPR") = Rating
                    DBTbl("TPPREF") = TRefID
                End If

                If IsNull(DBTbl("NBPR")) Then
                    DBTbl("NBP") = TValue
                    If Rating > -1 Then
                        DBTbl("NBPR") = Rating
                    End If
                    DBTbl("NBPREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "NBP" And DBTbl("NBPR") < Rating Then
                    DBTbl("NBP") = TValue
                    DBTbl("NBPR") = Rating
                    DBTbl("NBPREF") = TRefID
                End If

                If IsNull(DBTbl("LVOLR")) Then
                    DBTbl("LVOL") = TValue
                    If Rating > -1 Then
                        DBTbl("LVOLR") = Rating
                    End If
                    DBTbl("LVOLREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "LVOL" And DBTbl("LVOLR") < Rating Then
                    DBTbl("LVOL") = TValue
                    DBTbl("LVOLR") = Rating
                    DBTbl("LVOLREF") = TRefID
                End If

                If IsNull(DBTbl("HFORR")) Then
                    DBTbl("HFOR") = TValue
                    If Rating > -1 Then
                        DBTbl("HFORR") = Rating
                    End If
                    DBTbl("HFORREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "HFOR" And DBTbl("HFORR") < Rating Then
                    DBTbl("HFOR") = TValue
                    DBTbl("HFORR") = Rating
                    DBTbl("HFORREF") = TRefID
                End If

                If IsNull(DBTbl("GFORR")) Then
                    DBTbl("GFOR") = TValue
                    If Rating > -1 Then
                        DBTbl("GFORR") = Rating
                    End If
                    DBTbl("GFORREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "GFOR" And DBTbl("GFORR") < Rating Then
                    DBTbl("GFOR") = TValue
                    DBTbl("GFORR") = Rating
                    DBTbl("GFORREF") = TRefID
                End If

                If IsNull(DBTbl("ENTR")) Then
                    DBTbl("ENT") = TValue
                    If Rating > -1 Then
                        DBTbl("ENTR") = Rating
                    End If
                    DBTbl("ENTREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "ENT" And DBTbl("ENTR") < Rating Then
                    DBTbl("ENT") = TValue
                    DBTbl("ENTR") = Rating
                    DBTbl("ENTREF") = TRefID
                End If

                If IsNull(DBTbl("HFUSR")) Then
                    DBTbl("HFUS") = TValue
                    If Rating > -1 Then
                        DBTbl("HFUSR") = Rating
                    End If
                    DBTbl("HFUSREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "HFUS" And DBTbl("HFUSR") < Rating Then
                    DBTbl("HFUS") = TValue
                    DBTbl("HFUSR") = Rating
                    DBTbl("HFUSREF") = TRefID
                End If

                If IsNull(DBTbl("HCOMR")) Then
                    DBTbl("HCOM") = TValue
                    If Rating > -1 Then
                        DBTbl("HCOMR") = Rating
                    End If
                    DBTbl("HCOMREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "HCOM" And DBTbl("HCOMR") < Rating Then
                    DBTbl("HCOM") = TValue
                    DBTbl("HCOMR") = Rating
                    DBTbl("HCOMREF") = TRefID
                End If

                If IsNull(DBTbl("ACENR")) Then
                    DBTbl("ACEN") = TValue
                    If Rating > -1 Then
                        DBTbl("ACENR") = Rating
                    End If
                    DBTbl("ACENREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "ACEN" And DBTbl("ACENR") < Rating Then
                    DBTbl("ACEN") = TValue
                    DBTbl("ACENR") = Rating
                    DBTbl("ACENREF") = TRefID
                End If
    
                If IsNull(DBTbl("RGR")) Then
                    DBTbl("RG") = TValue
                    If Rating > -1 Then
                        DBTbl("RGR") = Rating
                    End If
                    DBTbl("RGREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "RG" And DBTbl("RGR") < Rating Then
                    DBTbl("RG") = TValue
                    DBTbl("RGR") = Rating
                    DBTbl("RGREF") = TRefID
                End If

                If IsNull(DBTbl("SOLPR")) Then
                    DBTbl("SOLP") = TValue
                    If Rating > -1 Then
                        DBTbl("SOLPR") = Rating
                    End If
                    DBTbl("SOLPREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "SOLP" And DBTbl("SOLPR") < Rating Then
                    DBTbl("SOLP") = TValue
                    DBTbl("SOLPR") = Rating
                    DBTbl("SOLPREF") = TRefID
                End If
    
                If IsNull(DBTbl("DMR")) Then
                    DBTbl("DM") = TValue
                    If Rating > -1 Then
                        DBTbl("DMR") = Rating
                    End If
                    DBTbl("DMREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "DM" And DBTbl("DMR") < Rating Then
                    DBTbl("DM") = TValue
                    DBTbl("DMR") = Rating
                    DBTbl("DMREF") = TRefID
                End If
    
                If IsNull(DBTbl("VDWVR")) Then
                    DBTbl("VDWV") = TValue
                    If Rating > -1 Then
                        DBTbl("VDWVR") = Rating
                    End If
                    DBTbl("VDWVREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "VDWV" And DBTbl("VDWVR") < Rating Then
                    DBTbl("VDWV") = TValue
                    DBTbl("VDWVR") = Rating
                    DBTbl("VDWVREF") = TRefID
                End If

                If IsNull(DBTbl("VDWAR")) Then
                    DBTbl("VDWA") = TValue
                    If Rating > -1 Then
                        DBTbl("VDWAR") = Rating
                    End If
                    DBTbl("VDWAREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "VDWA" And DBTbl("VDWAR") < Rating Then
                    DBTbl("VDWA") = TValue
                    DBTbl("VDWAR") = Rating
                    DBTbl("VDWAREF") = TRefID
                End If
    
                If IsNull(DBTbl("RIR")) Then
                    DBTbl("RI") = TValue
                    If Rating > -1 Then
                        DBTbl("RIR") = Rating
                    End If
                    DBTbl("RIREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "RI" And DBTbl("RIR") < Rating Then
                    DBTbl("RI") = TValue
                    DBTbl("RIR") = Rating
                    DBTbl("RIREF") = TRefID
                End If
    
                If IsNull(DBTbl("FPR")) Then
                    DBTbl("FP") = TValue
                    If Rating > -1 Then
                        DBTbl("FPR") = Rating
                    End If
                    DBTbl("FPREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "FP" And DBTbl("FPR") < Rating Then
                    DBTbl("FP") = TValue
                    DBTbl("FPR") = Rating
                    DBTbl("FPREF") = TRefID
                End If
    
                If IsNull(DBTbl("FLMLR")) Then
                    DBTbl("FLML") = TValue
                    If Rating > -1 Then
                        DBTbl("FLMLR") = Rating
                    End If
                    DBTbl("FLMLREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "FLML" And DBTbl("FLMLR") < Rating Then
                    DBTbl("FLML") = TValue
                    DBTbl("FLMLR") = Rating
                    DBTbl("FLMLREF") = TRefID
                End If

                If IsNull(DBTbl("FLMUR")) Then
                    DBTbl("FLMU") = TValue
                    If Rating > -1 Then
                        DBTbl("FLMUR") = Rating
                    End If
                    DBTbl("FLMUREF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "FLMU" And DBTbl("FLMUR") < Rating Then
                    DBTbl("FLMU") = TValue
                    DBTbl("FLMUR") = Rating
                    DBTbl("FLMUREF") = TRefID
                End If
    
                If IsNull(DBTbl("AITR")) Then
                    DBTbl("AIT") = TValue
                    If Rating > -1 Then
                        DBTbl("AITR") = Rating
                    End If
                    DBTbl("AITEF") = TRefID
                ElseIf Trim(Tbl801A!PropertyID) = "AIT" And DBTbl("AITR") < Rating Then
                    DBTbl("AIT") = TValue
                    DBTbl("AITR") = Rating
                    DBTbl("AITREF") = TRefID
                End If
              End If
              
              Tbl801A.MoveNext
            Loop
            Tbl801A.Close
            
            Set Tbl801B = db801.OpenRecordset( _
                "SELECT * FROM Tdep_Coefficients " & _
                "WHERE ChemID = " & CStr(CurCNum), dbOpenSnapshot)
            Tbl801B.MoveFirst
    
            Do While (Not Tbl801B.EOF)
                If IsNull(Tbl801B!EqnID) Then
                    TEqnID = 0
                Else
                    TEqnID = Tbl801B!EqnID
                End If
                
                If IsNull(Tbl801B!Num_Coeffs) Then
                    TNumCoeff = 0
                Else
                    TNumCoeff = Tbl801B!Num_Coeffs
                End If
                
                If IsNull(Tbl801B!A) Then
                    TA = 0
                Else
                    TA = Tbl801B!A
                End If
                
                If IsNull(Tbl801B!B) Then
                    TB = 0
                Else
                    TB = Tbl801B!B
                End If
                
                If IsNull(Tbl801B!C) Then
                    TC = 0
                Else
                    TC = Tbl801B!C
                End If
                
                If IsNull(Tbl801B!D) Then
                    TD = 0
                Else
                    TD = Tbl801B!D
                End If
                
                If IsNull(Tbl801B!E) Then
                    TE = 0
                Else
                    TE = Tbl801B!E
                End If
                
                If IsNull(Tbl801B!Min_T) Then
                    TMin = 0
                Else
                    TMin = Tbl801B!Min_T
                End If
                
                If IsNull(Tbl801B!Max_T) Then
                    TMax = 0
                Else
                    TMax = Tbl801B!Max_T
                End If
                
                If IsNull(Tbl801B!Data_Type) Then
                    TDataT = "T"
                Else
                    TDataT = Tbl801B!Data_Type
                End If
                
                If Trim(Tbl801B!PropertyID) = "SDN" Then
                    DBTbl("SDNEQN") = TEqnID
                    DBTbl("SDNNUM") = TNumCoeff
                    DBTbl("SDNA") = TA
                    DBTbl("SDNB") = TB
                    DBTbl("SDNC") = TC
                    DBTbl("SDND") = TD
                    DBTbl("SDNE") = TE
                    DBTbl("SDNTMIN") = TMin
                    DBTbl("SDNTMAX") = TMax
                    DBTbl("SDREL") = TDataT
                End If
    
                If Trim(Tbl801B!PropertyID) = "LDN" Then
                    DBTbl("LDNEQN") = TEqnID
                    DBTbl("LDNNUM") = TNumCoeff
                    DBTbl("LDNA") = TA
                    DBTbl("LDNB") = TB
                    DBTbl("LDNC") = TC
                    DBTbl("LDND") = TD
                    DBTbl("LDNE") = TE
                    DBTbl("LDNTMIN") = TMin
                    DBTbl("LDNTMAX") = TMax
                    DBTbl("LDNREL") = TDataT
                End If
                
                If Trim(Tbl801B!PropertyID) = "VP" Then
                    DBTbl("VPEQN") = TEqnID
                    DBTbl("VPNUM") = TNumCoeff
                    DBTbl("VPA") = TA
                    DBTbl("VPB") = TB
                    DBTbl("VPC") = TC
                    DBTbl("VPD") = TD
                    DBTbl("VPE") = TE
                    DBTbl("VPTMIN") = TMin
                    DBTbl("VPTMAX") = TMax
                    DBTbl("VPREL") = TDataT
                End If
    
                If Trim(Tbl801B!PropertyID) = "HVP" Then
                    DBTbl("HVPEQN") = TEqnID
                    DBTbl("HVPNUM") = TNumCoeff
                    DBTbl("HVPA") = TA
                    DBTbl("HVPB") = TB
                    DBTbl("HVPC") = TC
                    DBTbl("HVPD") = TD
                    DBTbl("HVPE") = TE
                    DBTbl("HVPTMIN") = TMin
                    DBTbl("HVPTMAX") = TMax
                    DBTbl("HVPREL") = TDataT
                End If
    
                If Trim(Tbl801B!PropertyID) = "SCP" Then
                    DBTbl("SCPEQN") = TEqnID
                    DBTbl("SCPNUM") = TNumCoeff
                    DBTbl("SCPA") = TA
                    DBTbl("SCPB") = TB
                    DBTbl("SCPC") = TC
                    DBTbl("SCPD") = TD
                    DBTbl("SCPE") = TE
                    DBTbl("SCPTMIN") = TMin
                    DBTbl("SCPTMAX") = TMax
                    DBTbl("SCPREL") = TDataT
               End If
                
                If Trim(Tbl801B!PropertyID) = "LCP" Then
                    DBTbl("LCPEQN") = TEqnID
                    DBTbl("LCPNUM") = TNumCoeff
                    DBTbl("LCPA") = TA
                    DBTbl("LCPB") = TB
                    DBTbl("LCPC") = TC
                    DBTbl("LCPD") = TD
                    DBTbl("LCPE") = TE
                    DBTbl("LCPTMIN") = TMin
                    DBTbl("LCPTMAX") = TMax
                    DBTbl("LCPREL") = TDataT
                End If
                
                If Trim(Tbl801B!PropertyID) = "ICP" Then
                    DBTbl("ICPEQN") = TEqnID
                    DBTbl("ICPNUM") = TNumCoeff
                    DBTbl("ICPA") = TA
                    DBTbl("ICPB") = TB
                    DBTbl("ICPC") = TC
                    DBTbl("ICPD") = TD
                    DBTbl("ICPE") = TE
                    DBTbl("ICPTMIN") = TMin
                    DBTbl("ICPTMAX") = TMax
                    DBTbl("ICPREL") = TDataT
               End If
                
                If Trim(Tbl801B!PropertyID) = "SVR" Then
                    DBTbl("SVREQN") = TEqnID
                    DBTbl("SVRNUM") = TNumCoeff
                    DBTbl("SVRA") = TA
                    DBTbl("SVRB") = TB
                    DBTbl("SVRC") = TC
                    DBTbl("SVRD") = TD
                    DBTbl("SVRE") = TE
                    DBTbl("SVRTMIN") = TMin
                    DBTbl("SVRTMAX") = TMax
                    DBTbl("SVRREL") = TDataT
                End If
                
                If Trim(Tbl801B!PropertyID) = "LVS" Then
                    DBTbl("LVSEQN") = TEqnID
                    DBTbl("LVSNUM") = TNumCoeff
                    DBTbl("LVSA") = TA
                    DBTbl("LVSB") = TB
                    DBTbl("LVSC") = TC
                    DBTbl("LVSD") = TD
                    DBTbl("LVSE") = TE
                    DBTbl("LVSTMIN") = TMin
                    DBTbl("LVSTMAX") = TMax
                    DBTbl("LVSREL") = TDataT
                End If
    
                If Trim(Tbl801B!PropertyID) = "VVS" Then
                    DBTbl("VVSEQN") = TEqnID
                    DBTbl("VVSNUM") = TNumCoeff
                    DBTbl("VVSA") = TA
                    DBTbl("VVSB") = TB
                    DBTbl("VVSC") = TC
                    DBTbl("VVSD") = TD
                    DBTbl("VVSE") = TE
                    DBTbl("VVSTMIN") = TMin
                    DBTbl("VVSTMAX") = TMax
                    DBTbl("VVSREL") = TDataT
                End If
                
                If Trim(Tbl801B!PropertyID) = "LTC" Then
                    DBTbl("LTCEQN") = TEqnID
                    DBTbl("LTCNUM") = TNumCoeff
                    DBTbl("LTCA") = TA
                    DBTbl("LTCB") = TB
                    DBTbl("LTCC") = TC
                    DBTbl("LTCD") = TD
                    DBTbl("LTCE") = TE
                    DBTbl("LTCTMIN") = TMin
                    DBTbl("LTCTMAX") = TMax
                    DBTbl("LTCREL") = TDataT
                End If
        
                If Trim(Tbl801B!PropertyID) = "VTC" Then
                    DBTbl("VTCEQN") = TEqnID
                    DBTbl("VTCNUM") = TNumCoeff
                    DBTbl("VTCA") = TA
                    DBTbl("VTCB") = TB
                    DBTbl("VTCC") = TC
                    DBTbl("VTCD") = TD
                    DBTbl("VTCE") = TE
                    DBTbl("VTCTMIN") = TMin
                    DBTbl("VTCTMAX") = TMax
                    DBTbl("VTCREL") = TDataT
                End If
        
                If Trim(Tbl801B!PropertyID) = "ST" Then
                    DBTbl("STEQN") = TEqnID
                    DBTbl("STNUM") = TNumCoeff
                    DBTbl("STA") = TA
                    DBTbl("STB") = TB
                    DBTbl("STC") = TC
                    DBTbl("STD") = TD
                    DBTbl("STE") = TE
                    DBTbl("STTMIN") = TMin
                    DBTbl("STTMAX") = TMax
                    DBTbl("STREL") = TDataT
                End If
                
                Tbl801B.MoveNext
            Loop
            Tbl801B.Close
            
            CurCNum = NewCNum
            DBTbl.Update
            DBTbl.AddNew
        End If
    
        If Trim(Tbl801!PropertyID) = "FORM" Then
            DBTbl("FORM") = Trim(Tbl801!Value)
        End If
        
        If Trim(Tbl801!PropertyID) = "NAME" Then
            DBTbl("NAME") = Trim(Tbl801!Value)
        End If
        
        If Trim(Tbl801!PropertyID) = "FAM" Then
            DBTbl("Family") = Trim(Tbl801!PropertyID)
        End If
    
        If Trim(Tbl801!PropertyID) = "CNAM" Then
            DBTbl("CNAM") = Trim(Tbl801!Value)
        End If
        
        If Trim(Tbl801!PropertyID) = "INAM" Then
            DBTbl("INAM") = Trim(Tbl801!Value)
        End If
                        
        If Trim(Tbl801!PropertyID) = "CASN" Then
            DBTbl("CAS") = Val(Tbl801!Value)

            DBTbl1.AddNew
            DBTbl1("CAS") = Val(Trim(Tbl801!Value))
            DBTbl1("Ref") = "PS "
            DBTbl1("Temp") = 0
            DBTbl1("Rating") = 10
            DBTbl1("Value") = 0
            DBTbl1.Update
        End If
                
        If Trim(Tbl801!PropertyID) = "STRU" Then
            DBTbl("STRU") = Trim(Tbl801!Value)
        End If
        
        Tbl801.MoveNext
        counter1 = counter1 + 1
        DisplayProgress counter1, size
    Loop
    Tbl801.Close
    
    Set db1 = OpenDatabase(DestPath & "\dbsave.mdb", False, False)
    Set DBTbl = db1.OpenRecordset("PrefDatabases", dbOpenTable)
    DBTbl.MoveFirst
    DBTbl.Edit
    DBTbl("Available") = True
    DBTbl.Update
    DBTbl.Close
    FRMMain!Label1.Caption = ""
End Sub

Sub DisplayProgress(NumRec As Long, DBSize As Long)
   'guestimate how much is done
   If (CDbl(NumRec) / DBSize) * 100 <= 100 Then
       FRMMain!SSPanel1.Value = (CDbl(NumRec) / DBSize) * 100
   End If
End Sub

Sub Extract801()
   
   Dim db As Database
   Dim db2 As Database
   Dim DBTbl As Recordset
   Dim DBTbl1 As Recordset
   Dim DBTbl2 As Recordset
   Dim Info As Variant
   Dim text As String
   Dim DNUM As String
   Dim ONUM As String
   Dim LDNUM As String
   Dim Start As Integer
   Dim size As Long
   Dim linesdone As Long
   Dim filenum As String
   Dim i As Integer
   Dim TmpFormula As String
   Dim TmpFamily As String
   Dim TmpName As String
   Dim TmpCAS As Double
   Dim DBDef As TableDef
   Dim PrimaryIndex As Index
   Dim Field1 As Field
   Dim counter1 As Long
   Dim Counter2 As Integer
   Dim file As Integer
   Dim reltext As String
   Dim aref As String
   Dim temprate As Integer
   Dim Rating As Integer
   'Open DBDestPath + "db801.lst" For Output As #2
   
   Call ExtractRefs
   counter1 = 0
   Counter2 = 0
   FRMMain!Label1.Caption = "Extracting Data..."
   FRMMain.Refresh
       
   'Select Case FRMMain!combo1.List(FRMMain!combo1.ListIndex)
   '   Case "DIPPR801 Database"
   '      On Error Resume Next
   '      Kill DBDestPath + "db801.mdb"
   '      On Error GoTo 0
   '      FileCopy App.Path + "\temp801.mdb", DBDestPath + "db801.mdb"
   '      TemplatePath = DBDestPath + "db801.mdb"
   'End Select
    FRMMain!Label1.Caption = "Please wait...Initializing"
    FRMMain!Label1.Refresh
    
   linesdone = 0
   Call GetFileSize(size, i, 1)
   
   On Error Resume Next
   
   Set db = OpenDatabase(DestPath & "\master.mdb", False, False)
   
   Set DBDef = db.TableDefs("DIPPR801")
   Set PrimaryIndex = DBDef.CreateIndex("PrimaryKey")
   PrimaryIndex.Primary = False
   PrimaryIndex.Unique = False
   Set Field1 = PrimaryIndex.CreateField("CAS")
   PrimaryIndex.Fields.Append Field1
'   DBDef.Indexes.Append PrimaryIndex
      
   Set DBTbl = db.OpenRecordset("DIPPR801", dbOpenTable)
   Set DBTbl1 = db.OpenRecordset("801TempDepend")
   'Delete current records in table
   'If the table is empty, skip this step
    If Not DBTbl.EOF Then
        DBTbl.MoveFirst
        Do While Not DBTbl.EOF
            DBTbl.Delete
            DBTbl.MoveNext
        Loop
    End If
    
    If Not DBTbl1.EOF Then
        DBTbl1.MoveFirst
        Do While Not DBTbl1.EOF
            DBTbl1.Delete
             DBTbl1.MoveNext
        Loop
    End If
 
 
   On Error GoTo 0
               
   file = 1

   filenum = Trim$(Str$(file))
   If file < 10 Then filenum = "0" + Trim$(Str$(file))
   Open DBTXTFilePath & filenum + ".dat" For Input As #1

   Start = True
    FRMMain!Label1.Caption = "Extracting Data..."
    FRMMain.Refresh
ReadNext:

   If EOF(1) = True Then
      'Write #2, TmpCAS, TmpName, "DIPPR801", TmpFamily, TmpFormula, ""
      DBTbl.Update
      Close #1
      file = file + 1
      If file > i - 1 Then
         DBTbl.Close
         'Turn on DIPPR801 switch
         Set db = OpenDatabase(DestPath & "\dbsave.mdb", False, False)
         Set DBTbl = db.OpenRecordset("PrefDatabases", dbOpenTable)
         DBTbl.MoveFirst
         DBTbl.Edit
         DBTbl("Available") = True
         DBTbl.Update
         DBTbl.Close
         FRMMain!Label1.Caption = ""
         'Close #2
         Exit Sub
      End If
      filenum = Trim$(Str$(file))
      If file < 10 Then filenum = "0" + Trim$(Str$(file))
      Open DBTXTFilePath & filenum + ".dat" For Input As #1
      Start = True
   End If

    Line Input #1, text

    linesdone = linesdone + 1
   
   'guestimate how much is done
   If Int((linesdone / size) * 8000) <= 100 Then
       FRMMain!SSPanel1.Value = Int((linesdone / size) * 8000)
   End If
   
    counter1 = counter1 + 1

   DNUM = Mid$(text, 1, 4)
         
   If Start = True Then
      LDNUM = DNUM
      Start = False
      DBTbl.AddNew
   
   End If

   If LDNUM <> DNUM Then
      LDNUM = DNUM
      'Write #2, TmpCAS, TmpName, "DIPPR801", TmpFamily, TmpFormula, ""
      DBTbl.Update
      DBTbl.AddNew
   End If

Check:

   If Mid$(text, 5, 4) = "FORM" Then
      Info = Mid$(text, 11, 60)
      DBTbl("FORM") = Info
      TmpFormula = Info
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "NAME" Then
      Info = Mid$(text, 9, 2)
      DBTbl("Family") = Info
      TmpFamily = Info
      Info = Mid$(text, 11, 60)
      DBTbl("NAME") = Info
      TmpName = Info
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "CNAM" Then
      Info = Mid$(text, 11, 60)
      DBTbl("CNAM") = Info
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "INAM" Then
      Info = Mid$(text, 11, 60)
      DBTbl("INAM") = Info
      GoTo ReadNext:
   End If
   
   If Mid$(text, 5, 4) = "CASN" Then
      Info = Mid$(text, 11, 60)
      DBTbl("CAS") = Val(Info)
      TmpCAS = Val(Info)
      
        'The following is the adding of the reference for the
        'f(t) properties. Since ALL f(t) properties are derived
        ' all f(t)propteries have a reference of "PS".
        ' In the future it would be better to just change the pearls
        ' data base but as for now we just change the entries.
        'Its done here cus this was where the cas is pulled which
        'is all we really need to make th entery
        DBTbl1.AddNew
        DBTbl1("CAS") = TmpCAS
        DBTbl1("Ref") = "PS "
        DBTbl1("Temp") = 0
        DBTbl1("Rating") = 10
        DBTbl1("Value") = 0
        DBTbl1.Update

      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "STRU" Then
      Info = Mid$(text, 11, 60)
      DBTbl("STRU") = Info
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "MW  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("MWR")) And Rating = -1 Then
            DBTbl("MW") = Val(Info)
            DBTbl("MWREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("MWR")) And Rating > -1 Then
            DBTbl("MW") = Val(Info)
            DBTbl("MWR") = Rating
            DBTbl("MWREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("MWR") < Rating Then
            DBTbl("MW") = Val(Info)
            DBTbl("MWR") = Rating
            DBTbl("MWREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "TC  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("TCR")) And Rating = -1 Then
            DBTbl("TC") = Val(Info)
            DBTbl("TCREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("TCR")) And Rating > -1 Then
            DBTbl("TC") = Val(Info)
            DBTbl("TCR") = Rating
            DBTbl("TCREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("TCR") < Rating Then
            DBTbl("TC") = Val(Info)
            DBTbl("TCR") = Rating
            DBTbl("TCREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If
   
   If Mid$(text, 5, 4) = "PC  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("PCR")) And Rating = -1 Then
            DBTbl("PC") = Val(Info)
            DBTbl("PCREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("PCR")) And Rating > -1 Then
            DBTbl("PC") = Val(Info)
            DBTbl("PCR") = Rating
            DBTbl("PCREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("PCR") < Rating Then
            DBTbl("PC") = Val(Info)
            DBTbl("PCR") = Rating
            DBTbl("PCREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "VC  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("VCR")) And Rating = -1 Then
            DBTbl("VC") = Val(Info)
            DBTbl("VCREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("VCR")) And Rating > -1 Then
            DBTbl("VC") = Val(Info)
            DBTbl("VCR") = Rating
            DBTbl("VCREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("VCR") < Rating Then
            DBTbl("VC") = Val(Info)
            DBTbl("VCR") = Rating
            DBTbl("VCREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "ZC  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("ZCR")) And Rating = -1 Then
            DBTbl("ZC") = Val(Info)
            DBTbl("ZCREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("ZCR")) And Rating > -1 Then
            DBTbl("ZC") = Val(Info)
            DBTbl("ZCR") = Rating
            DBTbl("ZCREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("ZCR") < Rating Then
            DBTbl("ZC") = Val(Info)
            DBTbl("ZCR") = Rating
            DBTbl("ZCREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "MP  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("MPR")) And Rating = -1 Then
            DBTbl("MP") = Val(Info)
            DBTbl("MPREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("MPR")) And Rating > -1 Then
            DBTbl("MP") = Val(Info)
            DBTbl("MPR") = Rating
            DBTbl("MPREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("MPR") < Rating Then
            DBTbl("MP") = Val(Info)
            DBTbl("MPR") = Rating
            DBTbl("MPREF") = Mid$(text, 24, 4)
         End If
    End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "TPT " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("TPTR")) And Rating = -1 Then
            DBTbl("TPT") = Val(Info)
            DBTbl("TPTREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("TPTR")) And Rating > -1 Then
            DBTbl("TPT") = Val(Info)
            DBTbl("TPTR") = Rating
            DBTbl("TPTREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("TPTR") < Rating Then
            DBTbl("TPT") = Val(Info)
            DBTbl("TPTR") = Rating
            DBTbl("TPTREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "TPP " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("TPPR")) And Rating = -1 Then
            DBTbl("TPP") = Val(Info)
            DBTbl("TPPREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("TPPR")) And Rating > -1 Then
            DBTbl("TPP") = Val(Info)
            DBTbl("TPPR") = Rating
            DBTbl("TPPREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("TPPR") < Rating Then
            DBTbl("TPP") = Val(Info)
            DBTbl("TPPR") = Rating
            DBTbl("TPPREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "NBP " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("NBPR")) And Rating = -1 Then
            DBTbl("NBP") = Val(Info)
            DBTbl("NBPREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("NBPR")) And Rating > -1 Then
            DBTbl("NBP") = Val(Info)
            DBTbl("NBPR") = Rating
            DBTbl("NBPREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("NBPR") < Rating Then
            DBTbl("NBP") = Val(Info)
            DBTbl("NBPR") = Rating
            DBTbl("NBPREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "LVOL" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("LVOLR")) And Rating = -1 Then
            DBTbl("LVOL") = Val(Info)
            DBTbl("LVOLREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("LVOLR")) And Rating > -1 Then
            DBTbl("LVOL") = Val(Info)
            DBTbl("LVOLR") = Rating
            DBTbl("LVOLREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("MWR") < Rating Then
            DBTbl("LVOL") = Val(Info)
            DBTbl("LVOLR") = Rating
            DBTbl("LVOLREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If
   
   If Mid$(text, 5, 4) = "HFOR" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("HFORR")) And Rating = -1 Then
            DBTbl("HFOR") = Val(Info)
            DBTbl("HFORREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("HFORR")) And Rating > -1 Then
            DBTbl("HFOR") = Val(Info)
            DBTbl("HFORR") = Rating
            DBTbl("HFORREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("HFORR") < Rating Then
            DBTbl("HFOR") = Val(Info)
            DBTbl("HFORR") = Rating
            DBTbl("HFORREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "GFOR" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("GFORR")) And Rating = -1 Then
            DBTbl("GFOR") = Val(Info)
            DBTbl("GFORREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("GFORR")) And Rating > -1 Then
            DBTbl("GFOR") = Val(Info)
            DBTbl("GFORR") = Rating
            DBTbl("GFORREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("GFORR") < Rating Then
            DBTbl("GFOR") = Val(Info)
            DBTbl("GFORR") = Rating
            DBTbl("GFORREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "ENT " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("ENTR")) And Rating = -1 Then
            DBTbl("ENT") = Val(Info)
            DBTbl("ENTREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("ENTR")) And Rating > -1 Then
            DBTbl("ENT") = Val(Info)
            DBTbl("ENTR") = Rating
            DBTbl("ENTREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("ENTR") < Rating Then
            DBTbl("ENT") = Val(Info)
            DBTbl("ENTR") = Rating
            DBTbl("ENTREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "HFUS" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("HFUSR")) And Rating = -1 Then
            DBTbl("HFUS") = Val(Info)
            DBTbl("HFUSREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("HFUSR")) And Rating > -1 Then
            DBTbl("HFUS") = Val(Info)
            DBTbl("HFUSR") = Rating
            DBTbl("HFUSREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("HFUSR") < Rating Then
            DBTbl("HFUS") = Val(Info)
            DBTbl("HFUSR") = Rating
            DBTbl("HFUSREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If
   
   If Mid$(text, 5, 4) = "HCOM" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("HCOMR")) And Rating = -1 Then
            DBTbl("HCOM") = Val(Info)
            DBTbl("HCOMREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("HCOMR")) And Rating > -1 Then
            DBTbl("HCOM") = Val(Info)
            DBTbl("HCOMR") = Rating
            DBTbl("HCOMREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("HCOMR") < Rating Then
            DBTbl("HCOM") = Val(Info)
            DBTbl("HCOMR") = Rating
            DBTbl("HCOMREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "ACEN" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("ACENR")) And Rating = -1 Then
            DBTbl("ACEN") = Val(Info)
            DBTbl("ACENREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("ACENR")) And Rating > -1 Then
            DBTbl("ACEN") = Val(Info)
            DBTbl("ACENR") = Rating
            DBTbl("ACENREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("ACENR") < Rating Then
            DBTbl("ACEN") = Val(Info)
            DBTbl("ACENR") = Rating
            DBTbl("ACENREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "RG  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("RGR")) And Rating = -1 Then
            DBTbl("RG") = Val(Info)
            DBTbl("RGREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("RGR")) And Rating > -1 Then
            DBTbl("RG") = Val(Info)
            DBTbl("RGR") = Rating
            DBTbl("RGREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("RGR") < Rating Then
            DBTbl("RG") = Val(Info)
            DBTbl("RGR") = Rating
            DBTbl("RGREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "SOLP" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("SOLPR")) And Rating = -1 Then
            DBTbl("SOLP") = Val(Info)
            DBTbl("SOLPREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("SOLPR")) And Rating > -1 Then
            DBTbl("SOLP") = Val(Info)
            DBTbl("SOLPR") = Rating
            DBTbl("SOLPREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("SOLPR") < Rating Then
            DBTbl("SOLP") = Val(Info)
            DBTbl("SOLPR") = Rating
            DBTbl("SOLPREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "DM  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("DMR")) And Rating = -1 Then
            DBTbl("DM") = Val(Info)
            DBTbl("DMREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("DMR")) And Rating > -1 Then
            DBTbl("DM") = Val(Info)
            DBTbl("DMR") = Rating
            DBTbl("DMREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("DMR") < Rating Then
            DBTbl("DM") = Val(Info)
            DBTbl("DMR") = Rating
            DBTbl("DMREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "VDWV" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("VDWVR")) And Rating = -1 Then
            DBTbl("VDWV") = Val(Info)
            DBTbl("VDWVREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("VDWVR")) And Rating > -1 Then
            DBTbl("VDWV") = Val(Info)
            DBTbl("VDWVR") = Rating
            DBTbl("VDWVREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("VDWVR") < Rating Then
            DBTbl("VDWV") = Val(Info)
            DBTbl("VDWVR") = Rating
            DBTbl("VDWVREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "VDWA" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("VDWAR")) And Rating = -1 Then
            DBTbl("VDWA") = Val(Info)
            DBTbl("VDWAREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("VDWAR")) And Rating > -1 Then
            DBTbl("VDWA") = Val(Info)
            DBTbl("VDWAR") = Rating
            DBTbl("VDWAREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("VDWAR") < Rating Then
            DBTbl("VDWA") = Val(Info)
            DBTbl("VDWAR") = Rating
            DBTbl("VDWAREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "RI  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("RIR")) And Rating = -1 Then
            DBTbl("RI") = Val(Info)
            DBTbl("RIREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("RIR")) And Rating > -1 Then
            DBTbl("RI") = Val(Info)
            DBTbl("RIR") = Rating
            DBTbl("RIREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("RIR") < Rating Then
            DBTbl("RI") = Val(Info)
            DBTbl("RIR") = Rating
            DBTbl("RIREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If
   
   If Mid$(text, 5, 4) = "FP  " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("FPR")) And Rating = -1 Then
            DBTbl("FP") = Val(Info)
            DBTbl("FPREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("FPR")) And Rating > -1 Then
            DBTbl("FP") = Val(Info)
            DBTbl("FPR") = Rating
            DBTbl("FPREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("FPR") < Rating Then
            DBTbl("FP") = Val(Info)
            DBTbl("FPR") = Rating
            DBTbl("FPREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "FLML" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("FLMLR")) And Rating = -1 Then
            DBTbl("FLML") = Val(Info)
            DBTbl("FLMLREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("FLMLR")) And Rating > -1 Then
            DBTbl("FLML") = Val(Info)
            DBTbl("FLMLR") = Rating
            DBTbl("FLMLREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("FLMLR") < Rating Then
            DBTbl("FLML") = Val(Info)
            DBTbl("FLMLR") = Rating
            DBTbl("FLMLREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "FLMU" Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("FLMUR")) And Rating = -1 Then
            DBTbl("FLMU") = Val(Info)
            DBTbl("FLMUREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("FLMUR")) And Rating > -1 Then
            DBTbl("FLMU") = Val(Info)
            DBTbl("FLMUR") = Rating
            DBTbl("FLMUREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("FLMUR") < Rating Then
            DBTbl("FLMU") = Val(Info)
            DBTbl("FLMUR") = Rating
            DBTbl("FLMUREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "AIT " Then
      Info = Mid$(text, 11, 12)
      If Mid$(text, 44, 1) = "A" Then
         Rating = calcrate(text)
         
         If IsNull(DBTbl("AITR")) And Rating = -1 Then
            DBTbl("AIT") = Val(Info)
            DBTbl("AITREF") = Mid$(text, 24, 4)
         ElseIf IsNull(DBTbl("AITR")) And Rating > -1 Then
            DBTbl("AIT") = Val(Info)
            DBTbl("AITR") = Rating
            DBTbl("AITREF") = Mid$(text, 24, 4)
         ElseIf DBTbl("AITR") < Rating Then
            DBTbl("AIT") = Val(Info)
            DBTbl("AITR") = Rating
            DBTbl("AITREF") = Mid$(text, 24, 4)
         End If
      End If
      GoTo ReadNext:
   End If
   
   If Mid$(text, 5, 4) = "SDNC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("SDNEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("SDNNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("SDNA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("SDNB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("SDNC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("SDND") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("SDNE") = Val(Info)
      Line Input #1, text   'now snag second line of data
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "SDNC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("SDNTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("SDNTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("SDREL") = Val(reltext)
            
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "SDNC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail
      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "LDNC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("LDNEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("LDNNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("LDNA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("LDNB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("LDNC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("LDND") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("LDNE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "LDNC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("LDNTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("LDNTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("LDNREL") = reltext
        
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "LDNC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
            GoTo Check:            'no more entries so bail
         
      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "VP C" Then
      Info = Mid$(text, 12, 3)
      DBTbl("VPEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("VPNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("VPA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("VPB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("VPC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("VPD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("VPE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "VP C" Then
         Info = Mid$(text, 10, 7)
         DBTbl("VPTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("VPTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("VPREL") = reltext
         
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "VP C"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail
 
      
      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "HVPC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("HVPEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("HVPNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("HVPA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("HVPB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("HVPC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("HVPD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("HVPE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "HVPC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("HVPTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("HVPTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("HVPREL") = reltext
            
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "HVPC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail
         
      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "SCPC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("SCPEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("SCPNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("SCPA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("SCPB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("SCPC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("SCPD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("SCPE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "SCPC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("SCPTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("SCPTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("SCPREL") = reltext
            
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "SCPC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail
    
      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "LCPC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("LCPEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("LCPNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("LCPA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("LCPB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("LCPC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("LCPD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("LCPE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "LCPC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("LCPTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("LCPTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("LCPREL") = reltext
            
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "LCPC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail
         
     Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "ICPC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("ICPEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("ICPNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("ICPA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("ICPB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("ICPC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("ICPD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("ICPE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "ICPC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("ICPTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("ICPTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("ICPREL") = reltext
            
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "ICPC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail
         

      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "SVRC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("SVREQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("SVRNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("SVRA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("SVRB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("SVRC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("SVRD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("SVRE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "SVRC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("SVRTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("SVRTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("SVRREL") = reltext

         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "SVRC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail

      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If
   
   If Mid$(text, 5, 4) = "LVSC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("LVSEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("LVSNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("LVSA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("LVSB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("LVSC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("LVSD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("LVSE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "LVSC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("LVSTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("LVSTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("LVSREL") = reltext
         
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "LVSC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail
      
      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "VVSC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("VVSEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("VVSNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("VVSA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("VVSB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("VVSC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("VVSD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("VVSE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "VVSC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("VVSTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("VVSTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("VVSREL") = reltext
            
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "VVSC"
            Line Input #1, text   'remove  extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail

      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "LTCC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("LTCEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("LTCNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("LTCA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("LTCB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("LTCC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("LTCD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("LTCE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "LTCC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("LTCTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("LTCTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("LTCREL") = reltext
        
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "LCTC"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail
         
      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   If Mid$(text, 5, 4) = "VTCC" Then
      Info = Mid$(text, 12, 3)
      DBTbl("VTCEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("VTCNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("VTCA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("VTCB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("VTCC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("VTCD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("VTCE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "VTCC" Then
         Info = Mid$(text, 10, 7)
         DBTbl("VTCTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("VTCTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("VTCREL") = reltext
        
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "VTCC"
            Line Input #1, text   'remove  extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail

      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If
   
   If Mid$(text, 5, 4) = "ST C" Then
      Info = Mid$(text, 12, 3)
      DBTbl("STEQN") = Val(Info)
      Info = Mid$(text, 16, 2)
      DBTbl("STNUM") = Val(Info)
      Info = Mid$(text, 20, 11)
      DBTbl("STA") = Val(Info)
      Info = Mid$(text, 32, 11)
      DBTbl("STB") = Val(Info)
      Info = Mid$(text, 44, 11)
      DBTbl("STC") = Val(Info)
      Info = Mid$(text, 56, 11)
      DBTbl("STD") = Val(Info)
      Info = Mid$(text, 68, 11)
      DBTbl("STE") = Val(Info)
      Line Input #1, text
      linesdone = linesdone + 1
      If Mid$(text, 5, 4) = "ST C" Then
         Info = Mid$(text, 10, 7)
         DBTbl("STTMIN") = Val(Info)
         Info = Mid$(text, 30, 7)
         DBTbl("STTMAX") = Val(Info)
         reltext = Mid$(text, 76, 3)
         DBTbl("STREL") = reltext
            
         Line Input #1, text   'now see if more data
         linesdone = linesdone + 1
         While Mid$(text, 5, 4) = "ST C"
            Line Input #1, text   'remove 2nd half of extra data
            linesdone = linesdone + 1
         Wend
         GoTo Check:            'no more entries so bail

      Else
         GoTo Check:
      End If
      GoTo ReadNext:
   End If

   GoTo ReadNext:
   
End Sub
Sub GetFileSize(size As Long, i As Integer, which As Integer)
    'which = 1 means 801 source (src) files
    'which = 2 2 means 801 reference (ref) files
    
    Dim filenum As String
    i = 1
GetFileSize:
    On Error GoTo StartExt:
    filenum = Trim$(Str$(i))
    If i < 10 Then filenum = "0" + Trim$(Str$(i))
    If which = 1 Then
        size = size + FileLen(DBTXTFilePath + filenum + ".dat")
    Else
        size = size + FileLen(DBTXTFileRefPath + filenum + ".dat")
    End If
    i = i + 1
    GoTo GetFileSize:
   
StartExt:
    
End Sub


Public Function GetUnits(unit As String) As String

    Dim size As Integer
    Dim last_char As String
    Dim result As Integer
    Dim valid As Integer
    valid = 0
    
    last_char = Right$(unit, 1)
    If StrComp(last_char, "K") = 0 Then
        GetUnits = "K"
        valid = 1
            
    End If
    If StrComp(last_char, ")") = 0 Then
        GetUnits = "C"
        valid = 1
    End If
                
    If StrComp(last_char, "C") = 0 Then
        GetUnits = "C"
        valid = 1
    End If
    If valid = 0 Then
        GetUnits = " "
    End If
    
End Function



Public Function calcrate(text As String)
    Dim calc As Integer
         calc = 10
        If Mid$(text, 39, 1) <> "P" And Mid$(text, 39, 1) <> "D" And Mid$(text, 39, 1) <> "S" And Mid$(text, 39, 1) <> " " Then
            calc = calc + 1
            
            If Mid$(text, 40, 1) <> "U" And Mid$(text, 40, 1) <> " " Then
                calc = calc + 1
            End If
            
            If Mid$(text, 41, 1) <> " " Then
                If CInt(Mid$(text, 41, 1)) < 6 And CInt(Mid$(text, 41, 1)) <> 0 Then
                    calc = calc + 1
                End If
            End If
        Else
            calc = -1
        End If
        
    calcrate = calc
End Function

Public Function calcratetemp(text As String) As Integer

    Dim calc As Integer
         calc = 10
         If Mid$(text, 76, 1) <> "P" And Mid$(text, 76, 1) <> "D" And Mid$(text, 76, 1) <> "S" And Mid$(text, 76, 1) <> " " Then
            calc = calc + 1
        End If
        If Mid$(text, 77, 1) <> "U" And Mid$(text, 77, 1) <> " " Then
            calc = calc + 1
        End If
        If Mid$(text, 78, 1) <> " " Then
            If CInt(Mid$(text, 78, 1)) < 6 And CInt(Mid$(text, 78, 1)) <> 0 Then
                calc = calc + 1
            End If
        End If
    calcratetemp = calc

End Function
