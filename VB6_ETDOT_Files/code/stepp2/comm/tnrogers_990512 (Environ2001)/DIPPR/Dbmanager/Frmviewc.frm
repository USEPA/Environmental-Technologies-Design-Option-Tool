VERSION 5.00
Begin VB.Form frmviewcalc 
   Caption         =   "View/Accept"
   ClientHeight    =   3675
   ClientLeft      =   690
   ClientTop       =   1695
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Done"
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdaccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame frresults 
      Caption         =   "Calculated Values"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.CheckBox ckmethod 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox ckmethod 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox ckmethod 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox ckmethod 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox ckmethod 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox ckmethod 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblunits 
         Caption         =   "Label1"
         Height          =   255
         Index           =   5
         Left            =   5880
         TabIndex        =   26
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblvalue 
         Caption         =   "lblvalue"
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   25
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblmethod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   24
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label lblunits 
         Caption         =   "Label1"
         Height          =   255
         Index           =   4
         Left            =   5880
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblunits 
         Caption         =   "Label1"
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblunits 
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblunits 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblunits 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblvalue 
         Caption         =   "lblvalue"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblvalue 
         Caption         =   "lblvalue"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblvalue 
         Caption         =   "lblvalue"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblvalue 
         Caption         =   "lblvalue"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblvalue 
         Caption         =   "lblvalue"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblmethod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   5
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label lblmethod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   4
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label lblmethod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblmethod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblmethod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmviewcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaccept_Click()

    ' this needs to take the data calculated and add
    ' it to the database if appropriate
    ' update fields:
    '   Cas #
    '   Property Code
    '   PEARLS Code
    '   Key Rating  (indicate user input) = -2
    '   Desc/Method
    '   Value
    '   Units
    '   Temperature
    '   TempUnits
    '   Comment (indicate user input) = "user input data"
    
    Dim localtable As Recordset
    Dim I As Integer
    Dim J As Integer
    Dim did_update As Boolean
    Dim accepted_temperature As Double
    Dim accepted_temp_units As String
    Dim fieldcas As String
    Dim fieldmethod As String
    Dim Criteria As String
    Dim method_value As Double
    Dim method_units As String
    Dim method_method As String
    Dim local_propcode As String
    ' get the temperature and units of calculation
    'On Error GoTo error_handler
    did_update = False
    If Right(Trim(input_name(global_cur_property)), 4) = "f(t)" Then
        If selected_temperature <> -1 Then
            accepted_temperature = selected_temperature
        Else
            accepted_temperature = 25
            selected_temperature = 25
        End If
        If selected_temp_units <> "" Then
            accepted_temp_units = selected_temp_units
        Else
            accepted_temp_units = "C"
            selected_temp_units = "C"
        End If
    Else
        accepted_temperature = 25
        accepted_temp_units = "C"
    End If
    ' first find the first method selected
    fieldcas = "Cas #"
    fieldmethod = "Desc/Method"
    local_propcode = get_prop_code(global_cur_property)
    Set localtable = chembrowsedb.OpenRecordset("DIPPR911", dbOpenDynaset)
    For I = 0 To MAX_METHODS_EACH - 1
        If frmviewcalc!ckmethod(I).Visible = False Then
            Exit For
        ElseIf frmviewcalc!ckmethod(I).value = UNCHECKED Then
            GoTo next_iteration
        End If
        did_update = True
        Criteria = "[" & fieldcas & "] = " & Val(selected_cas) & " And [" & fieldmethod & "] = '" & Trim(lblmethod(I).Caption) & "'"
        localtable.FindFirst Criteria
        
        If Not localtable.NoMatch Then
edit_existing:
            
            method_value = CDbl(frmviewcalc!lblvalue(I))
            method_units = frmviewcalc!lblunits(I)
            method_method = frmviewcalc!lblmethod(I)
            chembrowsedb.BeginTrans
            
            'If localtable("CAS #") <> selected_cas Then
            '    GoTo add_new
            'End If
            localtable.Edit
            localtable("Property Code") = local_propcode
            localtable("PEARLS Code") = global_cur_property
            localtable("Key Rating") = -2
            localtable("Rating") = -2
            localtable("Desc/Method") = method_method
            localtable("Value") = method_value
            localtable("Units") = method_units
            localtable("Temperature") = CStr(accepted_temperature)
            localtable("TempUnits") = accepted_temp_units
            localtable("Comment") = "user input data"
            localtable.Update
            chembrowsedb.CommitTrans
        Else
add_new:
            method_value = CDbl(frmviewcalc!lblvalue(I))
            method_units = frmviewcalc!lblunits(I)
            method_method = frmviewcalc!lblmethod(I)
            chembrowsedb.BeginTrans
            localtable.AddNew
            localtable("Cas #") = selected_cas
            localtable("Property Code") = local_propcode
            localtable("PEARLS Code") = global_cur_property
            localtable("Key Rating") = -2
            localtable("Rating") = -2
            localtable("Desc/Method") = method_method
            localtable("Value") = method_value
            localtable("Units") = method_units
            localtable("Temperature") = CStr(accepted_temperature)
            localtable("TempUnits") = accepted_temp_units
            localtable("Comment") = "user input data"
            localtable.Update
             chembrowsedb.CommitTrans
        End If
next_iteration:
    Next I
        localtable.Close
        chembrowsedb.Close
        Set chembrowsedb = OpenDatabase(curpath & curname, False, False)
        If did_update = True Then
            MsgBox ("database successfully updated")
        End If
        Exit Sub
error_handler:
        localtable.Close
        MsgBox ("an error occurred updating database")
End Sub


Private Sub cmdcancel_Click()
    frmviewcalc.Hide
    Unload Me
End Sub
