VERSION 5.00
Begin VB.Form frmShredDisp 
   Caption         =   "Structure Dissasembly Software"
   ClientHeight    =   7830
   ClientLeft      =   885
   ClientTop       =   2385
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7830
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2025
      TabIndex        =   76
      Top             =   7380
      Width           =   1815
   End
   Begin VB.Frame frusergr 
      Caption         =   "Groups Found"
      Height          =   2805
      Left            =   315
      TabIndex        =   9
      Top             =   3150
      Width           =   7395
      Begin VB.Line Line1 
         X1              =   3690
         X2              =   3690
         Y1              =   180
         Y2              =   2745
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   21
         Left            =   4410
         TabIndex        =   75
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   21
         Left            =   6615
         TabIndex        =   74
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   21
         Left            =   3825
         TabIndex        =   73
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   20
         Left            =   4410
         TabIndex        =   72
         Top             =   2295
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   20
         Left            =   6615
         TabIndex        =   71
         Top             =   2295
         Width           =   645
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   20
         Left            =   3825
         TabIndex        =   70
         Top             =   2295
         Width           =   495
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   19
         Left            =   4410
         TabIndex        =   69
         Top             =   2070
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   19
         Left            =   6615
         TabIndex        =   68
         Top             =   2070
         Width           =   645
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   19
         Left            =   3825
         TabIndex        =   67
         Top             =   2070
         Width           =   495
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   18
         Left            =   4410
         TabIndex        =   66
         Top             =   1845
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   18
         Left            =   6615
         TabIndex        =   65
         Top             =   1845
         Width           =   645
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   18
         Left            =   3825
         TabIndex        =   64
         Top             =   1845
         Width           =   495
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   17
         Left            =   4410
         TabIndex        =   63
         Top             =   1620
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   17
         Left            =   6615
         TabIndex        =   62
         Top             =   1620
         Width           =   645
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   17
         Left            =   3825
         TabIndex        =   61
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   16
         Left            =   4410
         TabIndex        =   60
         Top             =   1395
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   16
         Left            =   6615
         TabIndex        =   59
         Top             =   1395
         Width           =   645
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   16
         Left            =   3825
         TabIndex        =   58
         Top             =   1395
         Width           =   495
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   57
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   56
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   55
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   54
         Top             =   945
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   53
         Top             =   1170
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   52
         Top             =   1395
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   51
         Top             =   1620
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   50
         Top             =   1845
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   8
         Left            =   720
         TabIndex        =   49
         Top             =   2070
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   9
         Left            =   720
         TabIndex        =   48
         Top             =   2295
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   10
         Left            =   720
         TabIndex        =   47
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   11
         Left            =   4410
         TabIndex        =   46
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   12
         Left            =   4410
         TabIndex        =   45
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   13
         Left            =   4410
         TabIndex        =   44
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   0
         Left            =   2925
         TabIndex        =   43
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   1
         Left            =   2925
         TabIndex        =   42
         Top             =   495
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   2
         Left            =   2925
         TabIndex        =   41
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   3
         Left            =   2925
         TabIndex        =   40
         Top             =   945
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   4
         Left            =   2925
         TabIndex        =   39
         Top             =   1170
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   5
         Left            =   2925
         TabIndex        =   38
         Top             =   1395
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   6
         Left            =   2925
         TabIndex        =   37
         Top             =   1620
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   7
         Left            =   2925
         TabIndex        =   36
         Top             =   1845
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   8
         Left            =   2925
         TabIndex        =   35
         Top             =   2070
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   9
         Left            =   2925
         TabIndex        =   34
         Top             =   2295
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   10
         Left            =   2925
         TabIndex        =   33
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   11
         Left            =   6615
         TabIndex        =   32
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   12
         Left            =   6615
         TabIndex        =   31
         Top             =   495
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   13
         Left            =   6615
         TabIndex        =   30
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   14
         Left            =   4410
         TabIndex        =   29
         Top             =   945
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "(smiles string)"
         Height          =   255
         Index           =   15
         Left            =   4410
         TabIndex        =   28
         Top             =   1170
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   14
         Left            =   6615
         TabIndex        =   27
         Top             =   945
         Width           =   645
      End
      Begin VB.Label lblselno 
         Caption         =   "(quantity)"
         Height          =   255
         Index           =   15
         Left            =   6615
         TabIndex        =   26
         Top             =   1170
         Width           =   645
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   270
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   495
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   23
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   22
         Top             =   945
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   21
         Top             =   1170
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   20
         Top             =   1395
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   19
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   18
         Top             =   1845
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   17
         Top             =   2070
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   9
         Left            =   135
         TabIndex        =   16
         Top             =   2295
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   10
         Left            =   135
         TabIndex        =   15
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   11
         Left            =   3825
         TabIndex        =   14
         Top             =   270
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   12
         Left            =   3825
         TabIndex        =   13
         Top             =   495
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   13
         Left            =   3825
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   14
         Left            =   3825
         TabIndex        =   11
         Top             =   945
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "(group)"
         Height          =   255
         Index           =   15
         Left            =   3825
         TabIndex        =   10
         Top             =   1170
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Done"
      Height          =   375
      Left            =   4185
      TabIndex        =   3
      Top             =   7380
      Width           =   1815
   End
   Begin VB.Frame frselection 
      Caption         =   "Disassembly Settings"
      Height          =   2865
      Left            =   315
      TabIndex        =   0
      Top             =   135
      Width           =   7395
      Begin VB.ComboBox cboMethods 
         Height          =   315
         Left            =   2340
         TabIndex        =   78
         Top             =   2385
         Width           =   3525
      End
      Begin VB.ComboBox cboProperty 
         Height          =   315
         Left            =   3420
         TabIndex        =   77
         Top             =   1980
         Width           =   3525
      End
      Begin VB.TextBox tbxname 
         Height          =   330
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   5385
      End
      Begin VB.ComboBox cboschtype 
         Height          =   315
         Left            =   1710
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   1170
         Width           =   3525
      End
      Begin VB.TextBox tbxsmiles 
         Height          =   330
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   765
         Width           =   5385
      End
      Begin VB.Label Label1 
         Caption         =   "Please pick one of the following properties:"
         Height          =   285
         Left            =   270
         TabIndex        =   80
         Top             =   2025
         Width           =   3210
      End
      Begin VB.Label Label2 
         Caption         =   "Using the following method:"
         Height          =   285
         Left            =   270
         TabIndex        =   79
         Top             =   2430
         Width           =   3210
      End
      Begin VB.Label lblschtype 
         Caption         =   "Search Type:"
         Height          =   285
         Left            =   270
         TabIndex        =   5
         Top             =   1215
         Width           =   1305
      End
      Begin VB.Label lblname 
         Caption         =   "Chemical Name:"
         Height          =   285
         Left            =   270
         TabIndex        =   4
         Top             =   405
         Width           =   1305
      End
      Begin VB.Label lblsmiles 
         Caption         =   "SMILES Notation:"
         Height          =   285
         Left            =   270
         TabIndex        =   1
         Top             =   810
         Width           =   1305
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Status of chemical dissasembly:"
      Height          =   240
      Left            =   1665
      TabIndex        =   88
      Top             =   6075
      Width           =   2310
   End
   Begin VB.Label lblUnits 
      Height          =   240
      Left            =   6300
      TabIndex        =   87
      Top             =   7020
      Width           =   1500
   End
   Begin VB.Label lblProperty 
      Height          =   240
      Left            =   4050
      TabIndex        =   86
      Top             =   6390
      Width           =   3750
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   4050
      TabIndex        =   85
      Top             =   7020
      Width           =   2175
   End
   Begin VB.Label lblMethod 
      Height          =   240
      Left            =   4050
      TabIndex        =   84
      Top             =   6705
      Width           =   3750
   End
   Begin VB.Label Label5 
      Caption         =   "Results of first group found:"
      Height          =   240
      Left            =   1665
      TabIndex        =   83
      Top             =   7020
      Width           =   2310
   End
   Begin VB.Label Label4 
      Caption         =   "Method used for Calculation:"
      Height          =   240
      Left            =   1665
      TabIndex        =   82
      Top             =   6705
      Width           =   2310
   End
   Begin VB.Label Label3 
      Caption         =   "Property used for Calculation:"
      Height          =   240
      Left            =   1665
      TabIndex        =   81
      Top             =   6390
      Width           =   2310
   End
   Begin VB.Label lblstatus 
      Height          =   255
      Left            =   4050
      TabIndex        =   8
      Top             =   6075
      Width           =   3765
   End
   Begin VB.Menu mnuchem 
      Caption         =   "&Chemical"
      Begin VB.Menu mnubrowse 
         Caption         =   "&browse"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuedgr 
         Caption         =   "&groups"
      End
   End
End
Attribute VB_Name = "frmShredDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    tbxname.Text = selected_name
    tbxsmiles.Text = selected_smiles
'    Call clear_struct_groups
    
    lblstatus.Caption = ""
End Sub

Private Sub Form_Load()
    Dim i As Integer

    cboProperty.Clear
    cboProperty.AddItem " ", 0
    cboProperty.AddItem "Activity coefficient of chemical in water", 1
'    cboProperty.AddItem "Activity coefficient of water in chemical", 2
    cboProperty.AddItem "Aqueous Solubility", 2
    cboProperty.AddItem "Auto-Ignition T (AIT)", 3
'    cboProperty.AddItem "Biodegradability", 5
    cboProperty.AddItem "Critical Pressure", 4
    cboProperty.AddItem "Critical Temperature", 5
    cboProperty.AddItem "Critical Volume", 6
    cboProperty.AddItem "Freezing/Melting Point", 7
    cboProperty.AddItem "Henry's Constant", 8
    cboProperty.AddItem "Log10 Kow", 9
    cboProperty.AddItem "Molar Volume", 10
    cboProperty.AddItem "Normal Boiling Point", 11
    cboProperty.AddItem "Vapor Viscosit (Reichenberg Correlation)", 12
    
    cboschtype.Clear
    cboschtype.AddItem "Sequential, Non-Truncating"
    cboschtype.AddItem "Sequential, Truncating"
    cboschtype.AddItem "Combinatorial, Truncating"
    cboschtype.ListIndex = 2
    
    Call Set_struct_groups
End Sub

Private Sub mnubrowse_Click()
   
    Call load_chem_browse_info
    frmchembrowse.Show 1
    
'    cboProperty.ListIndex = 0
End Sub

Private Sub mnuedgr_Click()
    Screen.MousePointer = 11
    Call load_edit_groups_form
    Screen.MousePointer = 1
    frmeditgr.Show 1
End Sub

Private Sub cboMethods_Click()
    global_method = cboMethods.Text
    global_method_file = ""
    If cboMethods.ListIndex <> 0 Then
        cmdAccept.Enabled = True
        Select Case cboMethods.Text
            Case "Ambrose"
                global_method_file = "Ambrose.dat"
            Case "Boethling"
                global_method_file = "Boethling.dat"
            Case "Joback"
                global_method_file = "Joback.dat"
            Case "Fedors"
                global_method_file = "Fedors.dat"
            Case "Hine & Mookerjee"
                global_method_file = "Hine&Moo.dat"
            Case "LeBas"
                global_method_file = "LeBas.dat"
            Case "Lyderson"
                global_method_file = "Lyderson.dat"
            Case "MTU Logarithmic Groups (Pintar)", "MTU Linear Groups (Pintar)"
                global_method_file = "Pintar.dat"
            Case "Reichenberg"
                global_method_file = "Reichenberg.dat"
            Case "UNIFAC"
                global_method_file = "Unifac.dat"
        End Select
    End If
End Sub

Private Sub cboProperty_Click()
    cboMethods.Clear
    cmdAccept.Enabled = False
    If cboProperty.ListIndex <> 0 Then
        cboMethods.Clear
        cboMethods.AddItem " ", 0
        Select Case cboProperty.Text
            Case "Activity coefficient of chemical in water"
                cboMethods.AddItem "UNIFAC", 1
            Case "Aqueous Solubility"
                cboMethods.AddItem "UNIFAC", 1
            Case "Auto-Ignition T (AIT)"
                cboMethods.AddItem "MTU Logarithmic Groups (Pintar)", 1
                cboMethods.AddItem "MTU Linear Groups (Pintar)", 2
            Case "Biodegradability"
                cboMethods.AddItem "Boethling", 1
            Case "Critical Temperature"
                cboMethods.AddItem "Ambrose", 1
                cboMethods.AddItem "Joback", 2
                cboMethods.AddItem "Fedors", 3
                cboMethods.AddItem "Lyderson", 4
            Case "Critical Pressure"
                cboMethods.AddItem "Ambrose", 1
                cboMethods.AddItem "Joback", 2
                cboMethods.AddItem "Lyderson", 3
            Case "Critical Volume"
                cboMethods.AddItem "Ambrose", 1
                cboMethods.AddItem "Joback", 2
                cboMethods.AddItem "Lyderson", 3
            Case "Freezing/Melting Point"
                cboMethods.AddItem "Joback", 1
            Case "Henry's Constant"
                cboMethods.AddItem "Hine & Mookerjee", 1
                cboMethods.AddItem "UNIFAC", 2
            Case "Log10 Kow"
                cboMethods.AddItem "UNIFAC", 1
            Case "Molar Volume"
                cboMethods.AddItem "LeBas", 1
            Case "Normal Boiling Point"
                cboMethods.AddItem "Joback", 1
            Case "Normal Freezing Point"
                cboMethods.AddItem "Joback", 1
            Case "Vapor Viscosit (Reichenberg Correlation)"
                cboMethods.AddItem "Reichenberg", 1
        End Select
    End If
    If cboMethods.ListCount = 2 Then
        cboMethods.ListIndex = 1
    End If
End Sub


Private Sub cmdaccept_Click()

Dim OperatingTemp As Double
Dim CalcNumber As Double
Dim Mosdap_Result As Byte
Dim Cur_Method As String
Dim Cur_Property As String
Dim Units As String
Dim i As Integer
    If Trim(cboProperty.Text) = "" Then
        MsgBox "You must select a valid property.", vbExclamation, "DBManager"
        cboProperty.SetFocus
        Exit Sub
    ElseIf Trim(cboMethods.Text) = "" Then
        MsgBox "You must select a valid method.", vbExclamation, "DBManager"
        cboMethods.SetFocus
        Exit Sub
    End If
    If cboMethods.Text = "Ambrose" Or cboMethods.Text = "Boethling" Or cboMethods.Text = "Joback" _
                                    Or cboMethods.Text = "Fedors" Or cboMethods.Text = "Lyderson" _
                                    Or cboMethods.Text = "Reichenberg" Then
        MsgBox "This method (" & cboMethods.Text & ") doesn't shred properly with Mosdap.dll"
        Exit Sub
    End If
    Units = ""
    Mosdap_Result = Run_Mosdap(selected_smiles, global_method_file, cboschtype.ListIndex)
    Call Set_struct_groups
    If Mosdap_Result = 0 Then
        lblResult.Caption = "Error --- Error --- Error          "
        lblUnits.Caption = ""
        GoTo Mosdap_Error
    End If
    
    Cur_Method = cboMethods.Text
    Cur_Property = cboProperty.Text

    Select Case Cur_Method
        Case "Ambrose"
            If Calc_Ambrose(CalcNumber, Units, Cur_Property) = False Then
                GoTo Error_Function
            End If
        Case "Boethling"
        
        Case "Joback"
            If Calc_Joback(CalcNumber, Units, Cur_Property) = False Then
                GoTo Error_Function
            End If
        Case "Fedors"
            If Calc_Fedors(CalcNumber, Units, Cur_Property) = False Then
                GoTo Error_Function
            End If
        Case "Hine & Mookerjee"
            If Calc_Hine_and_Mookerjee(CalcNumber, Units, Cur_Property) = False Then
                GoTo Error_Function
            End If
        Case "LeBas"
            If Calc_Lebas(CalcNumber, Units, Cur_Property) = False Then
                GoTo Error_Function
            End If
        Case "Lyderson"
            If Calc_Lyderson(CalcNumber, Units, Cur_Property) = False Then
                GoTo Error_Function
            End If
            
        Case "MTU Logarithmic Groups (Pintar)", "MTU Linear Groups (Pintar)"
            If Calc_Pintar(CalcNumber, Units, Cur_Property, Cur_Method) = False Then
                GoTo Error_Function
            End If
        Case "Reichenberg"
            If Calc_Reichenberg(CalcNumber, Units, Cur_Property) = False Then
                GoTo Error_Function
            End If
        Case "UNIFAC"
            OperatingTemp = 25
            If Calc_Unifac(CalcNumber, Units, OperatingTemp, Cur_Property) = False Then
                GoTo Error_Function
            End If
        Case Else
            MsgBox "This method has not been developed", vbCritical + vbSystemModal, "Error"
    End Select
    
    lblResult.Caption = CalcNumber
    lblUnits.Caption = Units
'    MsgBox "The Calculated Value is: " & CalcNumber & " " & Units

Mosdap_Error:

    lblProperty.Caption = Cur_Property
    lblMethod.Caption = Cur_Method
    
    Select Case Mosdap_Result
        Case 0
            lblstatus.Caption = "Unable to disassemble " & selected_smiles
        Case 1
            lblstatus.Caption = "Successfully disassembled"
        Case 2
            lblstatus.Caption = "Partially disassembled"
        Case Else
            lblstatus.Caption = "An error occurred in the code while disassembling"
    End Select
    
Exit Sub

Error_Function:
    MsgBox "There was a problem with calculating method '" & Cur_Method & "'", vbCritical + vbSystemModal, "Error"
    lblResult.Caption = "Error --- Error --- Error          "
    lblUnits.Caption = ""
    GoTo Mosdap_Error
End Sub

Public Sub Set_struct_groups()
    Dim i As Integer
    
    For i = 0 To 21
        If cur_chem_groups(i) > 0 Then
            lblselindex(i).Caption = cur_chem_groups(i) & "."
            lblsel(i).Caption = group_smiles(cur_chem_groups(i) - 1)
            lblselno(i).Caption = num_cur_chem_groups(i)
        Else
            lblselindex(i).Caption = ""
            lblsel(i).Caption = ""
            lblselno(i).Caption = ""
        End If
    Next i
End Sub

