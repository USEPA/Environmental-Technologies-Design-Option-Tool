VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpreferences 
   Caption         =   "Environment Preferences"
   ClientHeight    =   4380
   ClientLeft      =   765
   ClientTop       =   1995
   ClientWidth     =   8880
   ControlBox      =   0   'False
   Icon            =   "frmprefe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   8880
   Begin MSFlexGridLib.MSFlexGrid GRDDefaultUnits 
      Height          =   1215
      Left            =   4560
      TabIndex        =   30
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2143
      _Version        =   65541
      Rows            =   57
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fire and Explosion Methods"
      Height          =   1935
      Left            =   4560
      TabIndex        =   21
      Top             =   120
      Width           =   4155
      Begin VB.ComboBox aitcmbx 
         Height          =   315
         Left            =   1620
         TabIndex        =   25
         Text            =   "Combo4"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox fpcmbx 
         Height          =   315
         Left            =   1620
         TabIndex        =   24
         Text            =   "Combo3"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox lflcmbx 
         Height          =   315
         Left            =   1620
         TabIndex        =   23
         Text            =   "Combo2"
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox uflcmbx 
         Height          =   315
         Left            =   1620
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label aitlbl 
         Caption         =   "AutoIgnition Temp"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1500
         Width           =   1395
      End
      Begin VB.Label fplbl 
         Caption         =   "Flash Point"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Label lfllbl 
         Caption         =   "Lower Fl Limit"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label ufllbl 
         Caption         =   "Upper Fl Limit"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   420
         Width           =   1395
      End
   End
   Begin VB.CommandButton CMDRestore 
      Caption         =   "Restore Defaults"
      Height          =   375
      Left            =   3660
      TabIndex        =   19
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame FRMBIPHierarchy 
      Caption         =   "BIP Hierarchy"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   4275
      Begin VB.ComboBox CMB2ndChoice 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Text            =   "CMB2ndChoice"
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox CMB3rdChoice 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Text            =   "CMB3rdChoice"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox CMB1stChoice 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Text            =   "CMB1stChoice"
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox CMBProperty 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Text            =   "CMBProperty"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label LBL3rdChoice 
         Caption         =   "3rd Choice"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label LBL2ndChoice 
         Caption         =   "2nd Choice"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   990
         Width           =   855
      End
      Begin VB.Label LBL1stChoice 
         Caption         =   "1st Choice"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   855
      End
      Begin VB.Label LBLProperty 
         Caption         =   "Property"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.CommandButton CMDCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5340
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton CMDSavePref 
      Caption         =   "Save and Exit"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame FRMSigFigs 
      Caption         =   "Significant Figures Displayed"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4275
      Begin VB.TextBox TXTGeneral 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Text            =   "TXTGeneral"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TXTLT001 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Text            =   "TXTLT001"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TXTGT1000 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Text            =   "TXTGT1000"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label LBLGeneralNum 
         Caption         =   "All other numbers"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   1110
         Width           =   1935
      End
      Begin VB.Label LBLLT001 
         Caption         =   "Numbers less than 0.001"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   750
         Width           =   1935
      End
      Begin VB.Label LBLGT1000 
         Caption         =   "Numbers greater than 1000"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   390
         Width           =   1935
      End
   End
   Begin VB.Label mthdlbl 
      Caption         =   "Method Preferences"
      Height          =   315
      Left            =   4680
      TabIndex        =   20
      Top             =   420
      Width           =   1935
   End
   Begin VB.Label LBLDefaultUnits 
      Caption         =   "Default Units"
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   2160
      Width           =   4215
   End
End
Attribute VB_Name = "frmpreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMB1stChoice_Click()

    Dim i As Integer
    Dim TempIndex As Integer
    
    If PrefStartup = True Then Exit Sub
    
    TempIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 1)
    
    BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 1) = FRMPreferences!CMB1stChoice.ListIndex + 1
    
    For i = 1 To 4
        If i <> 1 Then
            If BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, i) = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 1) Then
                BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, i) = TempIndex
            End If
        End If
    Next i
    
    PrefStartup = True
    FRMPreferences!CMB2ndChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 2) - 1
    FRMPreferences!CMB3rdChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 3) - 1
    PrefStartup = False

End Sub

Private Sub CMB2ndChoice_Click()

    Dim i As Integer
    Dim TempIndex As Integer
    
    If PrefStartup = True Then Exit Sub
    
    TempIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 2)
    
    BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 2) = FRMPreferences!CMB2ndChoice.ListIndex + 1
    
    For i = 1 To 4
        If i <> 2 Then
            If BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, i) = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 2) Then
                BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, i) = TempIndex
            End If
        End If
    Next i
    
    PrefStartup = True
    FRMPreferences!CMB1stChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 1) - 1
    FRMPreferences!CMB3rdChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 3) - 1
    PrefStartup = False

End Sub

Private Sub CMB3rdChoice_Click()

    Dim i As Integer
    Dim TempIndex As Integer
    
    If PrefStartup = True Then Exit Sub
    
    TempIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 3)
    
    BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 3) = FRMPreferences!CMB3rdChoice.ListIndex + 1
    
    For i = 1 To 4
        If i <> 3 Then
            If BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, i) = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 3) Then
                BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, i) = TempIndex
            End If
        End If
    Next i
    
    PrefStartup = True
    FRMPreferences!CMB1stChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 1) - 1
    FRMPreferences!CMB2ndChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 2) - 1
    PrefStartup = False

End Sub


Private Sub CMBProperty_Click()

    FRMPreferences!CMB1stChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 1) - 1
    FRMPreferences!CMB2ndChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 2) - 1
    FRMPreferences!CMB3rdChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 3) - 1

End Sub

Private Sub CMDCancel_Click()
    
    SetDefaultUnit = False
    
    Unload Me
    
End Sub


Private Sub CMDRestore_Click()

    Dim i As Integer
    
    'Set the modified flag
    WorkModified = True
    FRMMain.caption = "PEARLS:  " & SaveFileName & " modified"
    
    TXTGT1000.Text = "0.00E+00"
    TXTLT001.Text = "0.00E+00"
    TXTGeneral.Text = "0.00"
    
    For i = 1 To 4
        BIPHierarchy(1, i) = i
        BIPHierarchy(2, i) = i
        BIPHierarchy(3, i) = i
    Next i
    
    FRMPreferences!CMBProperty.ListIndex = 0
    FRMPreferences!CMB1stChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 1) - 1
    FRMPreferences!CMB2ndChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 2) - 1
    FRMPreferences!CMB3rdChoice.ListIndex = BIPHierarchy(FRMPreferences!CMBProperty.ListIndex + 1, 3) - 1
     
    For i = 0 To NumProperties
        FRMPreferences!GRDDefaultUnits.Row = i + 1
        FRMPreferences!GRDDefaultUnits.Col = 0
        FRMPreferences!GRDDefaultUnits.Text = GetPropName(i)
        FRMPreferences!GRDDefaultUnits.Col = 1
        FRMPreferences!GRDDefaultUnits.Text = Get_DefaultUnit(i)
    Next i
    
    FRMPreferences!GRDDefaultUnits.Row = i + 1
    FRMPreferences!GRDDefaultUnits.Col = 0
    FRMPreferences!GRDDefaultUnits.Text = "f(T) Temperatures"
    FRMPreferences!GRDDefaultUnits.Col = 1
    FRMPreferences!GRDDefaultUnits.Text = "K"

    ' the block 5 preferences
    ' remember that we've indexed the array from 0 but the
    ' actual preference methods and properties are indexed from 1
    For i = 0 To 3
        For J = 0 To NumMethods - 1
            B5Preference(i, J) = J + 1
        Next J
    Next i
    FRMPreferences!uflcmbx.Clear
    FRMPreferences!lflcmbx.Clear
    FRMPreferences!fpcmbx.Clear
    FRMPreferences!aitcmbx.Clear
    For i = 0 To NumMethods - 1
        prefstring = get_B5_method_name(1, i)
        If Len(prefstring) > 3 Then
            FRMPreferences!uflcmbx.AddItem prefstring
        End If
        prefstring = get_B5_method_name(2, i)
        If Len(prefstring) > 3 Then
            FRMPreferences!lflcmbx.AddItem prefstring
        End If
        prefstring = get_B5_method_name(3, i)
        If Len(prefstring) > 3 Then
            FRMPreferences!fpcmbx.AddItem prefstring
        End If
        prefstring = get_B5_method_name(4, i)
        If Len(prefstring) > 3 Then
            FRMPreferences!aitcmbx.AddItem prefstring
        End If
    Next i
    FRMPreferences!uflcmbx.ListIndex = 0
    FRMPreferences!lflcmbx.ListIndex = 0
    FRMPreferences!fpcmbx.ListIndex = 0
    FRMPreferences!aitcmbx.ListIndex = 0
End Sub

Private Sub CMDSavePref_Click()

    Dim DBTbl As Recordset
    Dim TempUnit As String
    Dim TempTFTUnit As String
    Dim property_name As String
    Screen.MousePointer = 11
    SetDefaultUnit = False
    
    'Set the modified flag
    WorkModified = True
    FRMMain.caption = "PEARLS:  " & SaveFileName & " modified"
    
    FormatGT1000 = Trim(TXTGT1000.Text)
    FormatLT001 = Trim(TXTLT001.Text)
    FormatGeneral = Trim(TXTGeneral.Text)
    
    Set DBTbl = DBJetUser.OpenRecordset("PrefBIPHierarchy", dbOpenTable)
    
    DBTbl.MoveFirst
    DBTbl.Edit
    DBTbl("BIP 1") = BIPHierarchy(1, 1)
    DBTbl("BIP 2") = BIPHierarchy(1, 2)
    DBTbl("BIP 3") = BIPHierarchy(1, 3)
    DBTbl("BIP 4") = BIPHierarchy(1, 4)
    DBTbl.Update
    DBTbl.MoveNext
    DBTbl.Edit
    DBTbl("BIP 1") = BIPHierarchy(2, 1)
    DBTbl("BIP 2") = BIPHierarchy(2, 2)
    DBTbl("BIP 3") = BIPHierarchy(2, 3)
    DBTbl("BIP 4") = BIPHierarchy(2, 4)
    DBTbl.Update
    DBTbl.MoveNext
    DBTbl.Edit
    DBTbl("BIP 1") = BIPHierarchy(3, 1)
    DBTbl("BIP 2") = BIPHierarchy(3, 2)
    DBTbl("BIP 3") = BIPHierarchy(3, 3)
    DBTbl("BIP 4") = BIPHierarchy(3, 4)
    DBTbl.Update
    DBTbl.Close
    
    Set DBTbl = DBJetUser.OpenRecordset("PrefFormatting", dbOpenTable)
    
    DBTbl.MoveFirst
    DBTbl.Edit
    If TXTGT1000.Text <> "" Then
        DBTbl("Setting") = FormatGT1000
    End If
    DBTbl.Update
    DBTbl.MoveNext
    DBTbl.Edit
    If TXTLT001.Text <> "" Then
        DBTbl("Setting") = FormatLT001
    End If
    DBTbl.Update
    DBTbl.MoveNext
    DBTbl.Edit
    If TXTGeneral.Text <> "" Then
        DBTbl("Setting") = FormatGeneral
    End If
    DBTbl.Update
    DBTbl.Close
    
    Set DBTbl = DBJetUser.OpenRecordset("PrefDefaultUnits", dbOpenTable)
    
    DBTbl.MoveFirst
    For i = 0 To NumProperties
        DBTbl.Edit
        DBTbl("Default Unit") = DefaultUnit(i)
        DBTbl.Update
        DBTbl.MoveNext
    Next i
        
    DBTbl.Edit
    DBTbl("Default Unit") = DefaultTFTUnit
    DBTbl.Update
    DBTbl.Close
    
    PrefStartup = False
    
    'Convert units to specified defaults
    For i = 0 To NumProperties
        CurProp = i
        TempUnit = InfoMethod(i).Unit
        TempTFTUnit = InfoMethod(i).TFTUnit
        If Trim(TempUnit) <> "" Then
            Call ConvertUnits(TempUnit, DefaultUnit(i))
        End If
        
        If Trim(TempTFTUnit) <> "" Then
            Call ConvertTFTUnits(TempTFTUnit, DefaultTFTUnit)
        End If
        
    Next i
    
    ' DENISE add block 5 stuff here
        Call update_B5_preferences
        On Error GoTo error_block_5_closed
        Set DBTbl = DBJetUser.OpenRecordset("Block5pref", dbOpenTable)
        On Error GoTo error_block_5_open
   
For prop_index = 1 To 4
DBTbl.Index = "PrimaryKey"
If prop_index = 1 Then
    property_name = "UFL"
ElseIf prop_index = 2 Then
    property_name = "LFL"
ElseIf prop_index = 3 Then
    property_name = "FP"
ElseIf prop_index = 4 Then
    property_name = "AIT"
End If
DBTbl.Seek "=", property_name
If DBTbl.NoMatch Then
    
        DBTbl.AddNew
        DBTbl("Property") = property_name
        DBTbl("Method1") = B5Preference(prop_index - 1, 0)
        DBTbl("Method2") = B5Preference(prop_index - 1, 1)
        DBTbl("Method3") = B5Preference(prop_index - 1, 2)
        DBTbl("Method4") = B5Preference(prop_index - 1, 3)
        DBTbl("Method5") = B5Preference(prop_index - 1, 4)
        DBTbl("Method6") = B5Preference(prop_index - 1, 5)
        DBTbl("Method7") = B5Preference(prop_index - 1, 6)
        DBTbl.Update
    
Else
        DBTbl.Edit
        DBTbl("Method1") = B5Preference(prop_index - 1, 0)
        DBTbl("Method2") = B5Preference(prop_index - 1, 1)
        DBTbl("Method3") = B5Preference(prop_index - 1, 2)
        DBTbl("Method4") = B5Preference(prop_index - 1, 3)
        DBTbl("Method5") = B5Preference(prop_index - 1, 4)
        DBTbl("Method6") = B5Preference(prop_index - 1, 5)
        DBTbl("Method7") = B5Preference(prop_index - 1, 6)
        DBTbl.Update
        
    
End If
Next prop_index
error_block_5_open:
    DBTbl.Close
    GoTo finish_sub
error_block_5_closed:
       ' Set DBDef = db.TableDefs("Block5Pref")
        'Set PrimaryIndex = DBDef.CreateIndex("PrimaryKey")
        'PrimaryIndex.Primary = False
        'PrimaryIndex.Unique = False
       ' Set Field1 = PrimaryIndex.CreateField("CAS")
       ' PrimaryIndex.fields.Append Field1
       ' DBDef.Indexes.Append PrimaryIndex
       ' Resume Next
        
finish_sub:
    If Cur_Info.CAS <> 0 Then
        Call DisplayProps
    End If
    ' Clear the Block 5 preference lists
    uflcmbx.Clear
    lflcmbx.Clear
    fpcmbx.Clear
    aitcmbx.Clear
    ' reset the mousepointer to arrow
    Screen.MousePointer = 1
    Unload Me
    
End Sub

Private Sub Form_Load()

   
    PrefStartup = True
    
    CenterForm Me
            
End Sub


Private Sub GRDDefaultUnits_Click()

    Dim i As Integer
       
    CurProp = GRDDefaultUnits.Row - 1
    
    If CurProp = 55 Then CurProp = -1
    
    FRMUnits!CMBUnits.Clear

    FRMUnits!CMBUnits.AddItem Get_DefaultUnit(CurProp)

    i = 1
    Do While Unit1(i) <> "End"
        If Trim(Unit1(i)) = Trim(Get_DefaultUnit(CurProp)) Then
            FRMUnits!CMBUnits.AddItem Unit2(i)
        ElseIf Trim(Unit2(i)) = Trim(Get_DefaultUnit(CurProp)) Then
            FRMUnits!CMBUnits.AddItem Unit1(i)
        End If
        i = i + 1
    Loop
    
    If CurProp = -1 Then 'aka f(t)
        FRMUnits!CMBUnits.Text = DefaultTFTUnit
    Else
        FRMUnits!CMBUnits.Text = DefaultUnit(CurProp)
    End If
    
    FRMUnits.Show 1
    
End Sub

