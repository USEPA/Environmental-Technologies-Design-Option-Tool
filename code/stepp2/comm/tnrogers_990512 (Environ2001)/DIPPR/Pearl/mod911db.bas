Attribute VB_Name = "mod911DB"

Sub Create911DBInfoForm(property_number As Integer, dbtype As Integer, CurrentUnits As String)

    Dim NewFRM911DBInfo As Form
    Dim i As Integer
    Dim existing As Integer
    
    existing = -1
    
    If dbtype = 9 Then
    
        For i = 0 To Forms.count - 1
            If Forms(i).caption = "911 Database Information: " & GetPropName(property_number) Then
                existing = i
                Exit For
            End If
        Next i
    Else
        For i = 0 To Forms.count - 1
            If Forms(i).caption = "801 Database Information: " & GetPropName(property_number) Then
                existing = i
                Exit For
            End If
        Next i
    End If
    
    If existing = -1 Then
        Set NewFRM911DBInfo = New frm911DBInfo
        Screen.MousePointer = 11
        If dbtype = 9 Then
            NewFRM911DBInfo.caption = "911 Database Information: " & GetPropName(property_number)
        Else
            NewFRM911DBInfo.caption = "801 Database Information: " & GetPropName(property_number)
        End If
        Call LoadFRM911DBFromInfo(NewFRM911DBInfo, property_number, dbtype, CurrentUnits)
        Screen.MousePointer = 1
        Call CenterForm(NewFRM911DBInfo)
        NewFRM911DBInfo.Show 1
        Exit Sub
    End If
        
            
    Screen.MousePointer = 11
    Call LoadFRM911DBFromInfo(Forms(existing), property_number, dbtype, CurrentUnits)
    Screen.MousePointer = 1
    Forms(existing).Show 1

End Sub

Sub LoadFRM911DBFromInfo(FRM As Form, PROPERTY_CODE As Integer, dbtype As Integer, CurrentUnits)

    Dim ArticleNum As Long
    Dim Title As String
    Dim Author As String
    Dim Journal As String
    Dim JDate As String
    Dim Volume As String
    Dim Number As String
    Dim Pages As String
    Dim DBTbl As Recordset
    Dim DBTbl1 As Recordset
    Dim rsTempDepend As Recordset
    Dim Ret As String
    Dim mintemp As Double
    Dim maxtemp As Double
    Dim TempDepend As Integer
    Dim reliable As String
    Dim tableref As Recordset
    Dim rating As Integer
    Dim pvalue As String
    Dim ref As String
    
    'these dim are the query to get the temperatures for temp depend props
  '  Dim Qd As QueryDef, Rs As Recordset
    
    
    
    Set tableref = DBJetMaster.OpenRecordset("801refs", dbOpenTable)
            
    Ret = Chr$(13) & Chr$(10)
    
    
    On Error Resume Next
    
    If dbtype = 9 Then
    
        Set DBTbl = DBJetMaster.OpenRecordset("DIPPR911", dbOpenTable)
        DBTbl.Index = "PrimaryKey2"
        DBTbl.Seek "=", Cur_Info.CAS, PROPERTY_CODE
    
        If DBTbl.NoMatch Then
            DBTbl.Close
            Exit Sub
        End If
    Else
        Set DBTbl = DBJetMaster.OpenRecordset("DIPPR801", dbOpenTable)
        DBTbl.Index = "PrimaryKey"
        DBTbl.Seek "=", Cur_Info.CAS
    
        If DBTbl.NoMatch Then
            DBTbl.Close
            Exit Sub
        End If
    End If
    
    If dbtype = 9 Then
        FRM!caption = "911 Database Information " & GetPropName(PROPERTY_CODE)
    Else
        FRM!caption = "801 Database Information " & GetPropName(PROPERTY_CODE)
    
    End If
    
    FRM!TXTCAS.Text = ""
    FRM!TXTChemName.Text = ""
    FRM!TXTProperty.Text = ""
    FRM!TXTValue.Text = ""
    FRM!TXTRating.Text = ""
    FRM!TXTTemperature.Text = ""
    FRM!TXTPressure.Text = ""
    FRM!TXT801Code.Text = ""
    FRM!TXTComment.Text = ""
    FRM!TXTCitations.Text = ""
    
    FRM!TXTCAS.Text = Trim(Cur_Info.CAS)
    FRM!TXTChemName.Text = Trim(Cur_Info.name)
    FRM!TXTProperty.Text = Trim(GetPropName(PROPERTY_CODE))
    'dbtype 8=801 9=911 so subtract 7 to get correct value from correct source
    FRM!TXTValue.Text = FormatVal(InfoMethod(PROPERTY_CODE).value(dbtype - 7))
    
    'set equation text which was extracted from  and set into global EquationText
    FRM!TXTequations = EquationText
        
    FRM!ValUnits = CurrentUnits
    FRM!PressUnits = Cur_Info.OpPUnit
    FRM!TempUnits = Cur_Info.OpTUnit
    If dbtype = 9 Then
        FRM!TXTTemperature.Visible = True
        FRM!LBLTemperature.Visible = True
        FRM!LBLPressure.Visible = True
        FRM!Label2.Visible = False
        FRM!LBLTempRange.Visible = False
        FRM!TXTMinTemp.Visible = False
        FRM!TXTMaxTemp.Visible = False
        FRM!TXTTemperature.Text.Visible = True
        FRM!TXTRating.Visible = True
        FRM!TXTPressure.Visible = True
        FRM!LBLPressure.Visible = True
        FRM!LBLRating.Visible = True
        FRM!TXTRating.Text = Trim(DBTbl("Rating"))
        FRM!TXTTemperature.Text = Trim(DBTbl("Temperature"))
        FRM!TXTPressure.Text = Trim(DBTbl("Pressure"))
        FRM!TXT801Code.Text = Trim(DBTbl("Desc/Method"))
        FRM!TXTComment.Text = Trim(DBTbl("Comment"))
        ArticleNum = Trim(DBTbl("Article #"))
    Else
        
        Call get801temps(DBTbl, mintemp, maxtemp, TempDepend, rating, pvalue, ref, reliable, GetPropName(PROPERTY_CODE))
        
        If TempDepend = 1 Then
            
            FRM!TXTTemperature.Visible = False
            FRM!LBLTemperature.Visible = False
            FRM!LBLTempRange.Visible = True
            FRM!Label2.Visible = True
            FRM!TXTMinTemp.Visible = True
            FRM!TXTMaxTemp.Visible = True
            FRM!TXTTemperature.Visible = False
            FRM!TXTRating.Visible = False
            FRM!TXTPressure.Visible = False
            FRM!LBLPressure.Visible = False
            FRM!LBLRating.Visible = False
            FRM!TXTMinTemp.Text = mintemp
            FRM!TXTMaxTemp.Text = maxtemp
    'more new stuff for tempdepend extraction
            
   '    SQLstr = "SELECT DISTINCTROW [801TempDepend].CAS, [801TempDepend].Temp FROM 801TempDepend WHERE (([801TempDepend].CAS=108883))"
    '        Set rsTempDepend = Db.OpenRecordset(SQL, dbOpenSnapshot)
     '       Data2.Recordset = rsTempDepend
            'Data2.Recordset = "801TempDepend"
            'Data2.Refresh
            
            'Set Qd = Data2.Database.QueryDefs("qryTempDepend")
            'Qd.Parameter("CasNum") = "test"
           
          ' Set Rs = Qd.OpenRecordset(Qd, dbOpenDynaset)
          ' Set Data2.Recordset = Rs
           
       Else
        
        Set DBTbl1 = DBJetMaster.OpenRecordset("801refs", dbOpenTable)
        DBTbl1.Index = "RefNum"
        DBTbl1.Seek "=", ref
    
        If DBTbl.NoMatch Then
            MsgBox ("error in references")
            DBTbl.Close
            Exit Sub
        End If
            
            FRM!TXTTemperature.Visible = True
            FRM!TXTMinTemp.Visible = False
            FRM!TXTMaxTemp.Visible = False
            FRM!TXTTemperature.Text = "N/A"
            FRM!TXTRating.Text = rating
            FRM!TXTPressure.Text = "" 'pvalue
            
            FRM!TXTCitations.Text = DBTbl1("Reference")
            
       End If
       
        'what the hell is this?
        'FRM!TXTPressure.Text = Trim(DBTbl("Pressure"))
        
        'this is the xu pu stuff another DCUT mod
        FRM!TXT801Code.Text = reliable
        
        
    
    End If
    
    'FRM.TXTequations.Text = Trim(DBTbl("Comment"))
    
    DBTbl.Close
    
    If dbtype = 9 Then
        Set DBTbl = DBJetMaster.OpenRecordset("CITATION", dbOpenTable)
    
        DBTbl.Index = "PrimaryKey"
        DBTbl.Seek "=", ArticleNum
    
        If DBTbl.NoMatch Then
            DBTbl.Close
            Exit Sub
        End If
    
        
        Title = Trim(DBTbl("Title"))
        Author = Trim(DBTbl("Author"))
        Journal = Trim(DBTbl("Journal"))
        JDate = Trim(DBTbl("Date"))
        Volume = Trim(DBTbl("Volume"))
        Number = Trim(DBTbl("Number"))
        Pages = Trim(DBTbl("Pages"))
        DBTbl.Close
    
        FRM!TXTCitations.Text = Author & ", " & Ret & Title & ", " & Ret & Journal & ", " & JDate & ", " & Volume & ", " & Number & ", " & Pages
    Else
        Set DBTbl = DBJetMaster.OpenRecordset("801REFS", dbOpenTable)

        'since any ref should occur only once either
        'findfirst will get it or its not there
        DBTbl.FindFirst "REFERENCE = '" & ref & "'"
        
        'Title = DBTbl("REFERENCE")
        If DBTbl.NoMatch Then
            DBTbl.Close
            Exit Sub
        End If
        Title = Trim(DBTbl("reference"))
        
    End If
End Sub


