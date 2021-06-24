VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmprint 
   Caption         =   "Print Settings"
   ClientHeight    =   3750
   ClientLeft      =   1950
   ClientTop       =   1920
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3750
   ScaleWidth      =   5895
   Begin VB.CommandButton previewcmd 
      Caption         =   "Preview"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CheckBox allchemck 
      Caption         =   "Select All Chemicals"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CheckBox allpropck 
      Caption         =   "Select All Properties"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ListBox chemlst 
      Height          =   1425
      Left            =   3000
      MultiSelect     =   1  'Simple
      TabIndex        =   8
      Top             =   360
      Width           =   2775
   End
   Begin VB.ListBox proplst 
      Height          =   1425
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton donecmd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton savecmd 
      Caption         =   "Save Settings"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton printcmd 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox SSFrame1 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   5595
      TabIndex        =   10
      Top             =   2160
      Width           =   5655
      Begin VB.OptionButton optgrchem 
         Caption         =   "Chemical"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optgrprop 
         Caption         =   "Property"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5160
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
   End
   Begin VB.Label proptxt 
      Caption         =   "Select Properties to Print"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label chemtxt 
      Caption         =   "Select Chemicals to Print"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim chem_count As Integer
Dim prop_count As Integer

Dim sel_cas() As String
Dim saveflag As Boolean






Private Sub allchemck_Click()

    Dim i As Integer
    Dim selected_flag As Boolean
        ' this code is backwards but for some reason it works
        ' flags seem to be working backwards here
    If allchemck.value = 0 Then
        selected_flag = False
    Else
        selected_flag = True
    End If
    
    For i = 0 To chemlst.ListCount - 1
        chemlst.Selected(i) = selected_flag
    Next i
    
End Sub

Private Sub allpropck_Click()

    Dim i As Integer
    Dim selected_flag As Boolean
        ' this code is backwards but for some reason it works
        ' flags seem to be working backwards here
    If allpropck.value = 0 Then
        selected_flag = False
    Else
        selected_flag = True
    End If
    For i = 0 To proplst.ListCount - 1
        proplst.Selected(i) = selected_flag
    Next i
End Sub

Private Sub donecmd_Click()

    saveflag = False
    Unload Me
    FRMMain.Show
    
End Sub


Private Sub Form_Load()

    Dim FNum As Integer
    Dim filename As String
    Dim i As Integer
    Dim temp As String
    
    'CenterForm Me
    
    saveflag = False
    Screen.MousePointer = 11
    Call fill_prop
    If fill_chem = False Then
        MsgBox ("Return to Main Form to Select chemical(s)")
        Exit Sub
    End If
    On Error GoTo no_file
    FNum = FreeFile
    filename = "printset.txt"
    Open filename For Input As #FNum
    Input #FNum, prop_count
    
    ' need to fix this
    For i = 0 To (prop_count - 1) And EOF(FNum) = False
        'temp = proplst.List(I)
        Input #FNum, temp
        proplst.Selected(temp) = True
    Next i
no_file:
    
    'For I = 0 To (chem_count - 1)
    '    chemlst.Selected(I) = True
    'Next I
    
    FRMPrint.Refresh
    CenterForm Me
    Screen.MousePointer = 1
End Sub



Private Sub optgrchem_Click()

    'ascck.Value = 2
    'descck.Value = 2
    
End Sub



Private Sub fill_prop()

    proplst.Clear
    proplst.AddItem "Molecular Weight"
    proplst.AddItem "Liquid Density @25"
    proplst.AddItem "Liquid Density as f(t)"
    proplst.AddItem "Melting Point"
    proplst.AddItem "Normal Boiling Point"
    proplst.AddItem "Vapor Pressure @25"
    proplst.AddItem "Vapor Pressure as f(t)"
    proplst.AddItem "Heat of Formation"
    proplst.AddItem "Liquid Heat Capacity"
    proplst.AddItem "Vapor Heat Capacity"
    proplst.AddItem "Heat of Vaporization @25"
    proplst.AddItem "Heat of Vaporization NBP"
    proplst.AddItem "Heat of Vaporization as f(t)"
    proplst.AddItem "Critical Temperature"
    proplst.AddItem "Critical Pressure"
    proplst.AddItem "Diffusivity of Chemical in Water"
    proplst.AddItem "Diffusivity of Chemical in Air"
    proplst.AddItem "Surface Tension @25"
    proplst.AddItem "Surface Tension as f(t)"
    proplst.AddItem "Vapor Viscosity as f(t)"
    proplst.AddItem "Liquid Viscosity as f(t)"
    proplst.AddItem "Liquid Thermal Conductivity as f(t)"
    proplst.AddItem "Vapor Thermal Conductivity as f(t)"
    proplst.AddItem "Upper Flammability Limit"
    proplst.AddItem "Lower Flammability Limit"
    proplst.AddItem "Flash Point"
    proplst.AddItem "Autoignition Temperature"
    proplst.AddItem "Heat of Combustion"
    proplst.AddItem "Carbonaceous ThOD"
    proplst.AddItem "Combined ThOD"
    proplst.AddItem "Chemical Oxygen Demand"
    proplst.AddItem "Biochemical Oxygen Demand"
    proplst.AddItem "Infinite Dilution Act. Coeff Water in Chem"
    proplst.AddItem "Henry's Constant"
    proplst.AddItem "Infinite Dilution Act. Coeff Chem in Water"
    proplst.AddItem "Log KOW"
    proplst.AddItem "Log KOC"
    proplst.AddItem "Bioconcentration Factor"
    proplst.AddItem "Critical Volume"
    proplst.AddItem "Solubility Limit Chemical in Water"
    proplst.AddItem "Solubility Limit Water in Chemical"
    proplst.AddItem "Antoine Coefficients"
   
End Sub


Private Sub optgrprop_Click()

    'ascck.Value = 0
    'descck.Value = 0
    
End Sub

Private Sub previewcmd_Click()

    Dim i As Integer
    Dim N As Integer

'mrt - 11/5/98
    Dim NumProps As Integer
    
    ReDim sel_prop(proplst.SelCount) As Integer
    ReDim sel_chem(chemlst.SelCount) As String
    ReDim sel_cas(chemlst.SelCount) As String
    Dim start_counter As Integer
    Dim hold_first As Integer
    Dim was_calculated As Boolean
    Dim cur_cas As String
    Dim FormulaString As String
    
    chem_count = chemlst.SelCount
    prop_count = 0
        ' first check that the user has selected at least one property and chemical
    If proplst.SelCount = 0 Then
        MsgBox ("Select property(s) to print")
        Exit Sub
    End If
    If chemlst.SelCount = 0 Then
        MsgBox ("Select chemical(s) to print")
        Exit Sub
    End If
        ' set the mouse to hourglass while we do all this stuff
    Screen.MousePointer = 11
    
        ' clear the database table in case the user has changed the CurMethod anywhere
    Call Clear_Print_Table
    
        ' next store (in sel_prop) the list of properties the user wants to print

'mrt - NumProperties wasn't working, so I just iterated to end of list - 11/5/98
    NumProps = proplst.ListCount
    print_antoine = False
    
    For i = 0 To NumProps - 1
        If proplst.Selected(i) = True And Not (i = NumProps - 1) Then
            sel_prop(prop_count) = i
            prop_count = prop_count + 1
        ElseIf proplst.Selected(i) = True And i = NumProps - 1 Then
            print_antoine = True
        End If
    Next i
    
        ' make sure if cur_info is selected it's printed first to save recalculation time
        ' this needs more work, commented out for now
        ' this is a bit strange but we store that index in hold_first, deselect that
        ' chemical, and then later reselect it using hold_first
    
    start_counter = 0
    'was_calculated = False
    'For I = 0 To chemlst.ListCount - 1
    '    If chemlst.Selected(I) = True Then
    '        If CStr(Cur_Info.CAS) = Trim(Right(chemlst.List(I), 9)) Then
     '           sel_cas(0) = CStr(Cur_Info.CAS)
     '           chemlst.Selected(I) = False
     '           hold_first = I
     '           was_calculated = True
      '          start_counter = 1
      '          Exit For
      '      End If
            
        'End If
    'Next I
        ' next store (in sel_cas) the list of chemicals (by cas) the user has selected
        ' if the first cas in the array was already set, start_counter will be 1 instead of 0
    For i = 0 To chemlst.ListCount - 1
        If chemlst.Selected(i) = True Then
            sel_cas(start_counter) = Trim(Right(chemlst.List(i), 9))
            start_counter = start_counter + 1
        End If
    Next i
    'If was_calculated = True Then
     '   chemlst.Selected(hold_first) = True
    'End If
        ' next call the function that will take care of recalculating each chemical
        ' and exporting the data to a table
        
    If export_custom_info(sel_cas, chem_count) = False Then
        Screen.MousePointer = 1
        Exit Sub
    End If
    saveflag = False
    
        ' now we set the parameters for the report itself
   
    CrystalReport1.DataFiles(0) = UserDBName
    CrystalReport1.Destination = 0
        ' initialize the selection formula string for the report
    FormulaString = " "
        ' the first case is if the user wants to print all properties grouped by chemical for chemical(s)
    If (allpropck.value = True Or proplst.SelCount = NumProps) And optgrprop.value = False Then
        CrystalReport1.ReportFileName = PathReport & "\chemone.rpt"
        'FormulaString = "("
        'Call select_chems(sel_cas, FormulaString)
        'FormulaString = FormulaString + ")"
        'CrystalReport1.SelectionFormula = FormulaString
        For N = 0 To chemlst.SelCount - 1
            CrystalReport1.SelectionFormula = "{PrintTable2.CAS} = " & Chr(34) & sel_cas(N) & Chr(34)
            CrystalReport1.Action = 1
        Next N
        Screen.MousePointer = 1
        Exit Sub
    End If
        ' the second is group by chemical, selected properties (this covers the case of all chemicals)
    If optgrchem.value = True Then
        CrystalReport1.ReportFileName = PathReport & "\chem.rpt"
        FormulaString = "("
        Call select_chems(sel_cas, FormulaString)
        FormulaString = FormulaString + ") And ("
        Call select_props(sel_prop, FormulaString)
        FormulaString = FormulaString + ")"
        CrystalReport1.SelectionFormula = FormulaString
        
        ' the third is group by property, selected properties
    ElseIf optgrprop.value = True Then
        
        CrystalReport1.ReportFileName = PathReport & "\prop.rpt"
        FormulaString = "("
        Call select_chems(sel_cas, FormulaString)
        FormulaString = FormulaString + ")"
        If proplst.SelCount <> proplst.ListCount Then
            FormulaString = FormulaString + " And ("
            Call select_props(sel_prop, FormulaString)
            FormulaString = FormulaString + ")"
        End If
        CrystalReport1.SelectionFormula = FormulaString
            ' finally, if the user didn't select a group by option, prompt them to and exit
    Else
        Screen.MousePointer = 1
        MsgBox ("Select 'group by' option")
        Exit Sub
    End If
        
    Screen.MousePointer = 1
    CrystalReport1.Action = 1
   
End Sub

Private Sub printcmd_Click()
    Dim i As Integer
        
'mrt - 11/5/98
    Dim NumProps As Integer
        
    Dim N As Integer
    ReDim sel_prop(proplst.SelCount) As Integer
    ReDim sel_chem(chemlst.SelCount) As String
    ReDim sel_cas(chemlst.SelCount) As String
    Dim start_counter As Integer
    Dim hold_first As Integer
    Dim cur_cas As String
    Dim FormulaString As String
    
    chem_count = chemlst.SelCount
    prop_count = 0
        ' first check that the user has selected at least one property and chemical
    If proplst.SelCount = 0 Then
        MsgBox ("Select property(s) to print")
        Exit Sub
    End If
    If chemlst.SelCount = 0 Then
        MsgBox ("Select chemical(s) to print")
        Exit Sub
    End If
        ' set the mouse to hourglass while we do all this stuff
    Screen.MousePointer = 11
    
        ' next store (in sel_prop) the list of properties the user wants to print

'mrt - NumProperties wasn't working, so I just iterated to end of list - 11/5/98
    NumProps = proplst.ListCount
    print_antoine = False
    
    For i = 0 To NumProps - 1
        If proplst.Selected(i) = True And Not (i = NumProps - 1) Then
            sel_prop(prop_count) = i
            prop_count = prop_count + 1
        ElseIf proplst.Selected(i) = True And i = NumProps - 1 Then
            print_antoine = True
        End If
    Next i
    
        ' make sure if cur_info is selected it's printed first to save recalculation time
        ' this needs more work, commented out for now
        ' this is a bit strange but we store that index in hold_first, deselect that
        ' chemical, and then later reselect it using hold_first
    start_counter = 0
    For i = 0 To chemlst.ListCount - 1
        If CStr(Cur_Info.CAS) = Trim(Right(chemlst.List(i), 9)) Then
            sel_cas(0) = CStr(Cur_Info.CAS)
            chemlst.Selected(i) = False
            hold_first = i
            start_counter = 1
            Exit For
        End If
    Next i
        ' next store (in sel_cas) the list of chemicals (by cas) the user has selected
        ' if the first cas in the array was already set, start_counter will be 1 instead of 0
    For i = 0 To chemlst.ListCount - 1
        If chemlst.Selected(i) = True Then
            sel_cas(start_counter) = Trim(Right(chemlst.List(i), 9))
            start_counter = start_counter + 1
        End If
    Next i
    chemlst.Selected(hold_first) = True
        ' next call the function that will take care of recalculating each chemical
        ' and exporting the data to a table
        
    If export_custom_info(sel_cas, chem_count) = False Then
        Screen.MousePointer = 1
        Exit Sub
    End If
    saveflag = False
    
        ' now we set the parameters for the report itself
   
    CrystalReport1.DataFiles(0) = UserDBName
    CrystalReport1.Destination = 1
        ' initialize the selection formula string for the report
    FormulaString = " "
        ' the first case is if the user wants to print all properties grouped by chemical for chemical(s)
    If (allpropck.value = True Or proplst.SelCount = NumProps) And optgrprop.value = False Then
        CrystalReport1.ReportFileName = PathReport & "\chemone.rpt"
        'FormulaString = "("
        'Call select_chems(sel_cas, FormulaString)
        'FormulaString = FormulaString + ")"
        'CrystalReport1.SelectionFormula = FormulaString
        For N = 0 To chemlst.SelCount - 1
            CrystalReport1.SelectionFormula = "{PrintTable2.CAS} = " & Chr(34) & sel_cas(N) & Chr(34)
            CrystalReport1.Action = 1
        Next N
        Screen.MousePointer = 1
        Exit Sub
    End If
        ' the second is group by chemical, selected properties (this covers the case of all chemicals)
    If optgrchem.value = True Then
        CrystalReport1.ReportFileName = PathReport & "\chem.rpt"
        FormulaString = "("
        Call select_chems(sel_cas, FormulaString)
        FormulaString = FormulaString & ") And ("
        Call select_props(sel_prop, FormulaString)
        FormulaString = FormulaString & ")"
        CrystalReport1.SelectionFormula = FormulaString
        
        ' the third is group by property, selected properties
    ElseIf optgrprop.value = True Then
        
        CrystalReport1.ReportFileName = PathReport & "\prop.rpt"
        FormulaString = "("
        Call select_chems(sel_cas, FormulaString)
        FormulaString = FormulaString & ")"
        If proplst.SelCount <> proplst.ListCount Then
            FormulaString = FormulaString & " And ("
            Call select_props(sel_prop, FormulaString)
            FormulaString = FormulaString & ")"
        End If
        CrystalReport1.SelectionFormula = FormulaString
            ' finally, if the user didn't select a group by option, prompt them to and exit
    Else
        Screen.MousePointer = 1
        MsgBox ("Select 'group by' option")
        Exit Sub
    End If
        
    Screen.MousePointer = 1
    CrystalReport1.Action = 1
   
    
   
End Sub

Private Sub savecmd_Click()

    Dim i As Integer
    
'mrt - 11/5/98
    Dim NumProps As Integer
    
    Dim FNum As Integer
    Dim filename As String
    ReDim sel_prop(proplst.ListCount) As Integer
    On Error GoTo not_open
    Screen.MousePointer = 11
    prop_count = 0
    
'mrt - NumProperties wasn't working, so I just iterated to end of list - 11/5/98
    NumProps = proplst.ListCount
    
    For i = 0 To (NumProps - 1)
        If proplst.Selected(i) = True Then
            sel_prop(prop_count) = i
            prop_count = prop_count + 1
        End If
    Next i
    FNum = FreeFile
    filename = "printset.txt"
    
save_settings:
    Open filename For Output As #FNum
    On Error GoTo open_file
    Write #FNum, prop_count
    For i = 0 To (prop_count - 1)
        Write #FNum, sel_prop(i)
    Next i
    Close #FNum
    saveflag = True
    Screen.MousePointer = 1
    Exit Sub
open_file:
    On Error GoTo not_open
    Screen.MousePointer = 1
    MsgBox ("error: unable to save settings")
    Close #FNum
    Exit Sub
not_open:
    Screen.MousePointer = 1
    MsgBox ("error: unable to save settings")
    Exit Sub

End Sub





Private Sub select_props(sel_prop() As Integer, FormulaString As String)

    Dim i As Integer
    Dim cur_prop As Integer
    
     For i = 0 To prop_count - 1
        
        If i <> 0 Then
            FormulaString = FormulaString + " Or "
        End If
        cur_prop = sel_prop(i)
        Select Case cur_prop
            Case 0:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 0"
            Case 1:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 1"
            Case 2:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 2"
            Case 3:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 3"
            Case 4:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 4"
            Case 5:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 5"
            Case 6:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 6"
            Case 7:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 7"
            Case 8:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 8"
            Case 9:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 9"
            Case 10:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 10"
            Case 11:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 11"
            Case 12:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 12"
            Case 13:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 13"
            Case 14:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 14"
            Case 15:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 15"
            Case 16:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 16"
            Case 17:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 17"
            Case 18:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 18"
            Case 19:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 19"
            Case 20:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 20"
            Case 21:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 21"
            Case 22:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 22"
            Case 23:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 23"
            Case 24:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 24"
            Case 25:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 25"
            Case 26:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 26"
            Case 27:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 27"
            Case 28:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 28"
            Case 29:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 29"
            Case 30:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 30"
            Case 31:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 31"
            Case 32:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 32"
            Case 33:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 33"
            Case 34:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 34"
            Case 35:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 35"
            Case 36:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 36"
            Case 37:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 37"
            Case 38:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 38"
            Case 39:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 39"
            Case 40:
                FormulaString = FormulaString + "{PrintTable2.Property Number} = 40"
            Case Else:
                ' do nothing
            End Select
    Next i
    
'mrt- because of the awkward numbering system for properties, antoine ended up being
'       number 55. This causes problems in the export_custom_info function and
'       should be fixed:(
    If i <> 0 Then
        FormulaString = FormulaString + " Or "
    End If
    If print_antoine = True Then
        FormulaString = FormulaString + "{PrintTable2.Property Number} = 55"
    End If
    
End Sub

Private Function fill_chem() As Boolean

    Dim i As Integer
    Dim DBTbl As Recordset
    
    ' first clear the list
    chemlst.Clear
    
    On Error GoTo DB_Closed_Error
    Set DBTbl = DBJetUser.OpenRecordset("User List", dbOpenTable)
    On Error GoTo DB_Open_Error
    DBTbl.MoveFirst
    While DBTbl.EOF = False
             ' the extra spaces are to hide the cas no
            chemlst.AddItem Trim(DBTbl("Name")) & "                                                 " & CStr(DBTbl("CAS"))
           
        DBTbl.MoveNext
    Wend
    DBTbl.Close
    fill_chem = True
    Exit Function
    
DB_Open_Error:
    MsgBox ("Error reading user database")
    DBTbl.Close
    Exit Function
DB_Closed_Error:
    MsgBox ("Can't find user database.  Use " & Chr(34) & "file preferences" & Chr(34) & " to set paths to user database.")
    Exit Function
End Function

Private Sub select_chems(sel_cas() As String, FormulaString As String)
    Dim i As Integer
    Dim cur_cas As String
    
    For i = 0 To chem_count - 1
        If i <> 0 Then
            FormulaString = FormulaString + " Or "
        End If
        cur_cas = Trim(sel_cas(i))
        FormulaString = FormulaString & "{PrintTable2.CAS} = " & Chr(34) & cur_cas & Chr(34)
    Next i
End Sub
