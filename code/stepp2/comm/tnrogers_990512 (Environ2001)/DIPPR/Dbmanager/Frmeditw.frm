VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmeditwizard 
   Caption         =   "Edit Wizard"
   ClientHeight    =   6105
   ClientLeft      =   1005
   ClientTop       =   825
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmethods 
      Caption         =   "Available Methods"
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   5535
      Begin VB.OptionButton optmethod 
         Caption         =   "Option1"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   22
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optmethod 
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton optmethod 
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optmethod 
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optmethod 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton optmethod 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmddone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save to Database"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdcalc 
      Caption         =   "&Calculate"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Frame frproperties 
      Caption         =   "Input Properties"
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   7455
      Begin VB.TextBox tbxinputprop 
         Height          =   375
         Index           =   5
         Left            =   5280
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox tbxinputprop 
         Height          =   375
         Index           =   4
         Left            =   5280
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.OptionButton optinorganic 
         Caption         =   "inorganic"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton optorganic 
         Caption         =   "organic"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox tbxinputprop 
         Height          =   375
         Index           =   3
         Left            =   5280
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox tbxinputprop 
         Height          =   375
         Index           =   2
         Left            =   5280
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox tbxinputprop 
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox tbxinputprop 
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblinputcalc 
         Caption         =   "Label1"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label lblinputcalc 
         Caption         =   "Label1"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label lblinputcalc 
         Caption         =   "Label1"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblinputprop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   26
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label lblinputprop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   25
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label lblinputprop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   10
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label lblinputprop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   9
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lblinputprop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label lblinputprop 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame frelements 
      Caption         =   "Chemical Elements"
      Height          =   1815
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      Begin MSFlexGridLib.MSFlexGrid grdelements 
         Height          =   1320
         Left            =   135
         TabIndex        =   30
         Top             =   360
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   2328
         _Version        =   393216
         Rows            =   10
         Cols            =   3
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.Frame frgroups 
      Caption         =   "Chemical Groups"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin MSFlexGridLib.MSFlexGrid grdgroups 
         Height          =   1320
         Left            =   135
         TabIndex        =   31
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2328
         _Version        =   393216
         Rows            =   17
         Cols            =   4
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnueditgroups 
         Caption         =   "&groups"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnushredder 
         Caption         =   "&shredder"
         Begin VB.Menu mnushredunifac 
            Caption         =   "UNIFAC"
         End
         Begin VB.Menu mnushredpintar 
            Caption         =   "Pintar"
         End
         Begin VB.Menu mnushredbenson 
            Caption         =   "Benson"
            Visible         =   0   'False
         End
         Begin VB.Menu mnushredlydersen 
            Caption         =   "Lydersen"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuhinemookerjee 
            Caption         =   "Hine Mookerjee"
         End
      End
      Begin VB.Menu mnuelfind 
         Caption         =   "&element finder"
      End
   End
   Begin VB.Menu mnuprop 
      Caption         =   "&Properties"
      Begin VB.Menu mnuselprop 
         Caption         =   "&select property"
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Options"
      Begin VB.Menu mnusettings 
         Caption         =   "settings"
      End
   End
End
Attribute VB_Name = "frmeditwizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcalc_Click()

    Dim i As Integer
    If global_cur_property < 0 Or global_cur_property > MAX_PROPERTIES Then
        MsgBox ("select a property to calculate")
        Exit Sub
    End If
    ' first clear out the stuff on the viewing form
    For i = 0 To MAX_METHODS_EACH - 1
        frmviewcalc!ckmethod(i).Visible = False
        frmviewcalc!lblMethod(i).Caption = ""
        frmviewcalc!lblMethod(i).Visible = False
        frmviewcalc!lblvalue(i).Caption = ""
        frmviewcalc!lblvalue(i).Visible = False
        frmviewcalc!lblUnits(i).Caption = ""
        frmviewcalc!lblUnits(i).Visible = False
    Next i
    Select Case global_cur_property
        Case FP
            Call start_do_fp
        Case LFL
            Call start_do_lfl
        Case UFL
            Call start_do_ufl
        Case AIT
            Call start_do_ait
        Case MW
            Call start_do_MW
        Case LD
            Call start_do_LD
        Case CV
            MsgBox ("Unable to find Lydersen groups needed for Critical Volume")
            'Call start_do_CV
        Case HC
            Call start_do_HC
        Case Vp
            Call start_do_vp
        Case Schem
            Call start_do_schem
        Case Swater
            Call start_do_swater
        Case ACchem
            Call start_do_acchem
        Case ACwater
            Call start_do_acwater
        Case ThODcarb
            Call start_do_ThODcarb
        Case ThODcomb
            Call start_do_ThODcomb
        Case BCF
            Call start_do_BCF
    End Select
    
End Sub


Private Sub cmddone_Click()

    frmeditwizard.Hide
    Unload Me
End Sub


Private Sub lblgroupnum_Click(Index As Integer)

End Sub

Private Sub lblnumgroup_Click(Index As Integer)

End Sub

Private Sub lblelement_Click(Index As Integer)
End Sub

Private Sub lblinputcalc_Click(Index As Integer)

    ' this allows the user to change the calculation preferences
    frmsettings.Show 1
    
End Sub

Private Sub mnueditgroups_Click()

    Screen.MousePointer = 11
    If load_edit_groups_form = True Then
        Screen.MousePointer = 1
        frmeditgr.Show 1
    End If
    Screen.MousePointer = 1
End Sub

Private Sub mnuelfind_Click()

If do_element_finder = True Then
    input_enabled(CONST_ELEMENTS) = True
Else
    input_enabled(CONST_ELEMENTS) = False
End If

End Sub

Private Sub mnuhinemookerjee_Click()

    If global_grouptype <> "Hine and Mookerjee" Then
        Call clear_group_array
        global_grouptype = "Hine and Mookerjee"
    End If
    global_groupfile = "Hine&moo.dat"
    If Trim(selected_smiles) <> "" Then
        If do_dbman_shredder > -1 Then
            input_enabled(CONST_HM_GROUPS) = True
            frmeditwizard!frgroups.Caption = "Hine and Mookerjee Groups"
        Else
            input_enabled(CONST_HM_GROUPS) = False
        End If
    Else
        input_enabled(CONST_HM_GROUPS) = False
    End If
    input_enabled(CONST_P_GROUPS) = False
    input_enabled(CONST_B_GROUPS) = False
    input_enabled(CONST_L_GROUPS) = False
    input_enabled(CONST_U_GROUPS) = False
            
End Sub

Private Sub mnuselprop_Click()

    Dim i As Integer
    Call load_frmselprop_info
    frmselprop.Show 1
    
End Sub

Private Sub mnusettings_Click()

    frmsettings.Show 1
    
End Sub

Private Sub mnushredbenson_Click()

    If global_grouptype <> "Benson" Then
        Call clear_group_array
        global_grouptype = "Benson"
    End If
    global_groupfile = "Benson.dat"
    If Trim(selected_smiles) <> "" Then
        If do_dbman_shredder > -1 Then
            input_enabled(CONST_B_GROUPS) = True
            frmeditwizard!frgroups.Caption = "Benson Groups"
        Else
            input_enabled(CONST_B_GROUPS) = False
        End If
    Else
        input_enabled(CONST_B_GROUPS) = False
    End If
    input_enabled(CONST_P_GROUPS) = False
    input_enabled(CONST_HM_GROUPS) = False
    input_enabled(CONST_L_GROUPS) = False
    input_enabled(CONST_U_GROUPS) = False
            
End Sub

Private Sub mnushredlydersen_Click()

    If global_grouptype <> "Lydersen" Then
        Call clear_group_array
        global_grouptype = "Lydersen"
    End If
    global_groupfile = "Lydersen.dat"
    If Trim(selected_smiles) <> "" Then
        If do_dbman_shredder > -1 Then
            input_enabled(CONST_L_GROUPS) = True
            frmeditwizard!frgroups.Caption = "Lydersen Groups"
        Else
            input_enabled(CONST_L_GROUPS) = False
        End If
    Else
        input_enabled(CONST_L_GROUPS) = False
    End If
    input_enabled(CONST_L_GROUPS) = True
    input_enabled(CONST_P_GROUPS) = False
    input_enabled(CONST_HM_GROUPS) = False
    input_enabled(CONST_U_GROUPS) = False
    input_enabled(CONST_B_GROUPS) = False
            
End Sub

Private Sub mnushredpintar_Click()
    
    If global_grouptype <> "Pintar" Then
        Call clear_group_array
        global_grouptype = "Pintar"
    End If
    global_groupfile = "Pintar.dat"
    If Trim(selected_smiles) <> "" Then
        If do_dbman_shredder > -1 Then
            input_enabled(CONST_P_GROUPS) = True
            frmeditwizard!frgroups.Caption = "Pintar Groups"
        Else
            input_enabled(CONST_P_GROUPS) = False
        End If
    Else
        input_enabled(CONST_P_GROUPS) = False
    End If
    input_enabled(CONST_U_GROUPS) = False
    input_enabled(CONST_HM_GROUPS) = False
    input_enabled(CONST_L_GROUPS) = False
    input_enabled(CONST_B_GROUPS) = False
            
End Sub

Private Sub mnushredunifac_Click()
     
    If global_grouptype <> "UNIFAC" Then
        Call clear_group_array
        global_grouptype = "UNIFAC"
    End If
    global_groupfile = "Unifac.dat"
    
    If Trim(selected_smiles) <> "" Then
        If do_dbman_shredder > -1 Then
            input_enabled(CONST_U_GROUPS) = True
            frmeditwizard!frgroups.Caption = "UNIFAC Groups"
        Else
            input_enabled(CONST_U_GROUPS) = False
        End If
    Else
        input_enabled(CONST_U_GROUPS) = False
    End If
    input_enabled(CONST_P_GROUPS) = False
    input_enabled(CONST_HM_GROUPS) = False
    input_enabled(CONST_L_GROUPS) = False
    input_enabled(CONST_B_GROUPS) = False
            
End Sub



Private Sub optmethod_Click(Index As Integer)

    ' this function needs to update the inputs listed when
    ' the method selected changes
    Call update_edwiz_input_info(Index)
End Sub



