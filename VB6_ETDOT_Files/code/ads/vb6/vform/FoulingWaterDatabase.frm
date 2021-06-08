VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmFoulingWaterDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Water Fouling Correlation Database"
   ClientHeight    =   6435
   ClientLeft      =   1710
   ClientTop       =   2655
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5760
   Begin Threed.SSFrame SSFrame1 
      Height          =   2625
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   4630
      _StockProps     =   14
      Caption         =   "Select a Water Type:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstCorrelations 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   5235
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2835
      Left            =   120
      TabIndex        =   1
      Top             =   2850
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   5001
      _StockProps     =   14
      Caption         =   "Empirical Kinetic Constants for:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtCoeff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1980
         TabIndex        =   10
         Text            =   "txtCoeff(4)"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtCoeff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1980
         TabIndex        =   9
         Text            =   "txtCoeff(3)"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCoeff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1980
         TabIndex        =   8
         Text            =   "txtCoeff(2)"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtCoeff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   7
         Text            =   "txtCoeff(1)"
         Top             =   720
         Width           =   1455
      End
      Begin Threed.SSCommand cmdRecord 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2310
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&New"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdRecord 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2310
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Edit"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdRecord 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2310
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Delete"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "txtName"
         Top             =   300
         Width           =   5235
      End
      Begin Threed.SSCommand cmdRecord 
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2310
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Save"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdRecord 
         Height          =   375
         Index           =   4
         Left            =   3960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2310
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Cancel Edit"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   18
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   17
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   16
         Top             =   1470
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   15
         Top             =   1830
         Width           =   1155
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   14
         Top             =   750
         Width           =   600
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1/min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   13
         Top             =   1110
         Width           =   600
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   12
         Top             =   1470
         Width           =   600
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1/min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   11
         Top             =   1830
         Width           =   600
      End
   End
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save any changes you have made to this database"
      Top             =   5820
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   0
      Left            =   3300
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes you have made to this database"
      Top             =   5820
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmFoulingWaterDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim New_Correlation As Integer

Dim FORM_MODE As Integer
Const FORM_MODE_VIEW = 1
Const FORM_MODE_EDIT = 2
Const FORM_MODE_ADDNEW = 3

Dim HALT_LSTCORRELATIONS As Boolean

'//////// COMMUNICATIONS WITH frmFoulingWaterDatabase: /////////////////////////////////////////////////
Private Type frmFoulingWaterDatabase_Record_Type
  K1 As Double
  K2 As Double
  K3 As Double
  K4 As Double
  Name As String
End Type
Dim Local_Record As frmFoulingWaterDatabase_Record_Type






Const frmFoulingWaterDatabase_declarations_end = True



Sub frmFoulingWaterDatabase_Edit()
  frmFoulingWaterDatabase.Show 1
End Sub


Sub Populate_lstCorrelations()
Dim SAVE_INDEX As Integer
Dim i As Integer
  If (lstCorrelations.ListIndex >= 0) Then
    SAVE_INDEX = lstCorrelations.ListIndex
  Else
    SAVE_INDEX = 0
  End If
  HALT_LSTCORRELATIONS = True
  lstCorrelations.Clear
  For i = 1 To Number_Water_Correlations
    lstCorrelations.AddItem Correlations_For_Water(i).Name
  Next i
  HALT_LSTCORRELATIONS = False
  If (SAVE_INDEX > lstCorrelations.ListCount - 1) Then
    SAVE_INDEX = lstCorrelations.ListCount - 1
  End If
  If (SAVE_INDEX >= 0) And (SAVE_INDEX <= lstCorrelations.ListCount - 1) Then
    lstCorrelations.ListIndex = SAVE_INDEX
  End If
End Sub
Sub frmFoulingWaterDatabase_Repopulate_Values()
Dim Frm As Form
Set Frm = frmFoulingWaterDatabase
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtCoeff(1), Local_Record.K1)
  Call unitsys_set_number_in_base_units( _
      Frm.txtCoeff(2), Local_Record.K2)
  Call unitsys_set_number_in_base_units( _
      Frm.txtCoeff(3), Local_Record.K3)
  Call unitsys_set_number_in_base_units( _
      Frm.txtCoeff(4), Local_Record.K4)
  'TEXT DATA.
  Frm.txtName = Trim$(Local_Record.Name)
End Sub
Sub frmFoulingWaterDatabase_Refresh()
Dim Frm As Form
Set Frm = frmFoulingWaterDatabase
Dim TextLocked As Boolean
  'REPOPULATE VALUES.
  Call frmFoulingWaterDatabase_Repopulate_Values
  'LOCK/UNLOCK TEXTBOXES AND LISTBOX.
  TextLocked = (FORM_MODE = FORM_MODE_VIEW)
  txtCoeff(1).Locked = TextLocked
  txtCoeff(2).Locked = TextLocked
  txtCoeff(3).Locked = TextLocked
  txtCoeff(4).Locked = TextLocked
  txtName.Locked = TextLocked
  lstCorrelations.Enabled = TextLocked
  'DISABLE/ENABLE BUTTONS.
  Select Case FORM_MODE
    Case FORM_MODE_VIEW:
      If Frm.lstCorrelations.ListCount = 0 Then
        Frm.cmdRecord(0).Enabled = True       'NEW.
        Frm.cmdRecord(1).Enabled = False      'EDIT.
        Frm.cmdRecord(2).Enabled = False      'DELETE.
      Else
        If Frm.lstCorrelations.ListCount >= Max_Number_Water_Correlations Then
          Frm.cmdRecord(0).Enabled = False    'NEW.
          Frm.cmdRecord(1).Enabled = True     'EDIT.
          Frm.cmdRecord(2).Enabled = True     'DELETE.
        Else
          Frm.cmdRecord(0).Enabled = True     'NEW.
          Frm.cmdRecord(1).Enabled = True     'EDIT.
          Frm.cmdRecord(2).Enabled = True     'DELETE.
        End If
      End If
      Frm.cmdRecord(3).Enabled = False        'SAVE.
      Frm.cmdRecord(4).Enabled = False        'CANCEL EDIT.
      Frm.cmdCancelOK(0).Enabled = True       'CANCEL.
      Frm.cmdCancelOK(1).Enabled = True       'OK.
    Case FORM_MODE_EDIT, FORM_MODE_ADDNEW:
      Frm.cmdRecord(0).Enabled = False        'NEW.
      Frm.cmdRecord(1).Enabled = False        'EDIT.
      Frm.cmdRecord(2).Enabled = False        'DELETE.
      Frm.cmdRecord(3).Enabled = True         'SAVE.
      Frm.cmdRecord(4).Enabled = True         'CANCEL EDIT.
      Frm.cmdCancelOK(0).Enabled = False      'CANCEL.
      Frm.cmdCancelOK(1).Enabled = False      'OK.
  End Select
End Sub
Sub frmFoulingWaterDatabase_PopulateUnits()
  Call unitsys_register(frmFoulingWaterDatabase, lblDesc(1), _
      txtCoeff(1), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmFoulingWaterDatabase, lblDesc(2), _
      txtCoeff(2), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmFoulingWaterDatabase, lblDesc(3), _
      txtCoeff(3), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmFoulingWaterDatabase, lblDesc(4), _
      txtCoeff(4), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Private Sub Load_Water_Correlations(flag As Integer)
Dim f As Integer, N As Integer, i As Integer
On Error GoTo Error_In_Reading_Corr
  f = FreeFile
  Open Database_Path & "\water_co.txt" For Input As f
  Input #f, N
  If N > Max_Number_Water_Correlations Then
    flag = True
    Close (f)
    Call Show_Error("Too many correlations in the file.")
    Exit Sub
  End If
  For i = 1 To N
  Input #f, Correlations_For_Water(i).Name, _
      Correlations_For_Water(i).Coeff(1), _
      Correlations_For_Water(i).Coeff(2), _
      Correlations_For_Water(i).Coeff(3), _
      Correlations_For_Water(i).Coeff(4)
  Next i
  Close (f)
  Number_Water_Correlations = N
  flag = False
  Exit Sub
Error_In_Reading_Corr:
  Call Show_Error("Error while reading the file containing correlations.")
  flag = True
  Resume Exit_Corr_Compound
Exit_Corr_Compound:
End Sub
Sub Store_Water_Correlations()
Dim f As Integer
Dim i As Integer
  On Error GoTo Error_In_Writing_File
  f = FreeFile
  Open Database_Path & "\water_co.txt" For Output As f
  Write #f, Number_Water_Correlations
  For i = 1 To Number_Water_Correlations
  Write #f, Correlations_For_Water(i).Name, _
      Correlations_For_Water(i).Coeff(1), _
      Correlations_For_Water(i).Coeff(2), _
      Correlations_For_Water(i).Coeff(3), _
      Correlations_For_Water(i).Coeff(4)
  Next i
  Close (f)
  Exit Sub
Error_In_Writing_File:
  Call Show_Error("Error writing to file.")
  Close (f)
  Resume Exit_Here
Exit_Here:
End Sub
Sub Load_Local_Record(RecNum As Integer)
  Local_Record.Name = _
    Correlations_For_Water(RecNum).Name
  Local_Record.K1 = _
    Correlations_For_Water(RecNum).Coeff(1)
  Local_Record.K2 = _
    Correlations_For_Water(RecNum).Coeff(2)
  Local_Record.K3 = _
    Correlations_For_Water(RecNum).Coeff(3)
  Local_Record.K4 = _
    Correlations_For_Water(RecNum).Coeff(4)
End Sub
Sub Store_Local_Record(RecNum As Integer)
  Correlations_For_Water(RecNum).Name = _
    Local_Record.Name
  Correlations_For_Water(RecNum).Coeff(1) = _
    Local_Record.K1
  Correlations_For_Water(RecNum).Coeff(2) = _
    Local_Record.K2
  Correlations_For_Water(RecNum).Coeff(3) = _
    Local_Record.K3
  Correlations_For_Water(RecNum).Coeff(4) = _
    Local_Record.K4
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
Dim resp As Integer, f As Integer, k As Integer, j As Integer
Dim RetVal As Integer
  Select Case Index
    Case 0:   'CANCEL.
      RetVal = MsgBox("Are you sure you want to exit without " & _
        "saving the database ?", vbQuestion + vbYesNo, _
        AppName_For_Display_Short & " : Exit Without Saving Database ?")
      If (RetVal = vbNo) Then Exit Sub
      Call Load_Water_Correlations(i)
      If i Then Exit Sub
      Unload Me
    Case 1:   'OK.
      RetVal = MsgBox("Are you sure you want to " & _
        "save the database ?", vbQuestion + vbYesNo, _
        AppName_For_Display_Short & " : Save Database ?")
      If (RetVal = vbNo) Then Exit Sub
      Call Store_Water_Correlations
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub cmdRecord_Click(Index As Integer)
Dim RetVal As Integer
Dim New_Rec_Index As Integer
Dim Del_Rec_Index As Integer
Dim Edit_Rec_Index As Integer
Dim i As Integer
  Select Case Index
    Case 0:   'NEW. ///////////////////////////////////////////////////////////////////////
      If (FORM_MODE <> FORM_MODE_VIEW) Then Exit Sub
      If (lstCorrelations.ListCount >= Max_Number_Water_Correlations) Then
        Exit Sub
      End If
      FORM_MODE = FORM_MODE_ADDNEW
      'SET DEFAULT SETTINGS FOR NEW RECORD.
      Local_Record.Name = "New Water Correlation"
      Local_Record.K1 = 1#
      Local_Record.K2 = 0#
      Local_Record.K3 = 0#
      Local_Record.K4 = 0#
      'REFRESH WINDOW.
      Call frmFoulingWaterDatabase_Refresh
    Case 1:   'EDIT. //////////////////////////////////////////////////////////////////////
      If (FORM_MODE <> FORM_MODE_VIEW) Then Exit Sub
      Edit_Rec_Index = lstCorrelations.ListIndex + 1
      If (Edit_Rec_Index < 1) Or (Edit_Rec_Index > Number_Water_Correlations) Then
        Call Show_Error("You must first select a correlation.")
        Exit Sub
      End If
      FORM_MODE = FORM_MODE_EDIT
      'REFRESH WINDOW.
      Call frmFoulingWaterDatabase_Refresh
    Case 2:   'DELETE. ////////////////////////////////////////////////////////////////////
      If (FORM_MODE <> FORM_MODE_VIEW) Then Exit Sub
      Del_Rec_Index = lstCorrelations.ListIndex + 1
      If (Del_Rec_Index < 1) Or (Del_Rec_Index > Number_Water_Correlations) Then
        Call Show_Error("You must first select a correlation.")
        Exit Sub
      End If
      For i = Del_Rec_Index To Number_Water_Correlations - 1
        Correlations_For_Water(i) = Correlations_For_Water(i + 1)
      Next i
      Number_Water_Correlations = Number_Water_Correlations - 1
      'REPOPULATE LISTBOX.
      Call Populate_lstCorrelations
      'REFRESH WINDOW.
      Call frmFoulingWaterDatabase_Refresh
    Case 3:   'SAVE. //////////////////////////////////////////////////////////////////////
      If (FORM_MODE = FORM_MODE_VIEW) Then Exit Sub
      Select Case FORM_MODE
        Case FORM_MODE_EDIT:
          Call Store_Local_Record(lstCorrelations.ListIndex + 1)
        Case FORM_MODE_ADDNEW:
          Number_Water_Correlations = Number_Water_Correlations + 1
          New_Rec_Index = Number_Water_Correlations
          Call Store_Local_Record(New_Rec_Index)
      End Select
      FORM_MODE = FORM_MODE_VIEW
      'REPOPULATE LISTBOX.
      Call Populate_lstCorrelations
      lstCorrelations.ListIndex = lstCorrelations.ListCount - 1
      'REFRESH WINDOW.
      Call frmFoulingWaterDatabase_Refresh
    Case 4:   'CANCEL EDIT. ///////////////////////////////////////////////////////////////
      If (FORM_MODE = FORM_MODE_VIEW) Then Exit Sub
      'RetVal = MsgBox("Do you want to exit without " & _
      '  "saving the edited data ?", vbQuestion + vbYesNo, _
      '  AppName_For_Display_Short)
      'If (RetVal = vbNo) Then Exit Sub
      FORM_MODE = FORM_MODE_VIEW
      'REPOPULATE LISTBOX.
      Call Populate_lstCorrelations
      'REFRESH WINDOW.
      Call frmFoulingWaterDatabase_Refresh
  End Select
End Sub


Private Sub Form_Load()
Dim i As Integer, j As Integer
  Call CenterOnForm(Me, frmFouling)
  i = False
  Call Load_Water_Correlations(i)
  If i Then Number_Water_Correlations = 0
  Call Populate_lstCorrelations
  If (Number_Water_Correlations >= 1) Then
    Call Load_Local_Record(1)
  End If
  FORM_MODE = FORM_MODE_VIEW
  'POPULATE UNIT CONTROLS.
  Call frmFoulingWaterDatabase_PopulateUnits
  'REFRESH WINDOW.
  Call frmFoulingWaterDatabase_Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub lstCorrelations_Click()
Dim ThisIndex As Integer
  If (HALT_LSTCORRELATIONS) Then Exit Sub
  ThisIndex = lstCorrelations.ListIndex + 1
  If (ThisIndex <= lstCorrelations.ListCount) Then
    Call Load_Local_Record(lstCorrelations.ListIndex + 1)
  End If
  'REFRESH WINDOW.
  Call frmFoulingWaterDatabase_Refresh
End Sub


Private Sub txtCoeff_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtCoeff(Index)
  Call unitsys_control_txtx_gotfocus(Ctl)
End Sub
Private Sub txtCoeff_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtCoeff_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtCoeff(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
  Select Case Index
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
    Case 2: Val_Low = -1E+20: Val_High = 1E+20
    Case 3: Val_Low = -1E+20: Val_High = 1E+20
    Case 4: Val_Low = -1E+20: Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 1: Local_Record.K1 = NewValue
        Case 2: Local_Record.K2 = NewValue
        Case 3: Local_Record.K3 = NewValue
        Case 4: Local_Record.K4 = NewValue
      End Select
      'RAISE DIRTY FLAG IF NECESSARY.
      If (Raise_Dirty_Flag) Then
        ''THROW DIRTY FLAG.
        'Call frmCompoProp_DirtyStatus_Throw
      End If
      'REFRESH WINDOW.
      Call frmFoulingWaterDatabase_Refresh
    End If
  End If
End Sub


Private Sub txtName_GotFocus()
Dim Ctl As Control
Set Ctl = txtName
  Call Global_GotFocus(Ctl)
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtName_LostFocus()
Dim Ctl As Control
Set Ctl = txtName
Dim OldValueStr As String
  'HANDLE STRING FIELDS.
  OldValueStr = Trim$(Local_Record.Name)
  'NOTE: ZERO-LENGTH STRINGS FOR 8 AND 11 ARE ALLOWED.
  If (Trim$(Ctl.Text) = "") Then
    Ctl.Text = OldValueStr
    'Call Show_Error("You must enter a non-blank string for the carbon name.")
    'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
    'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
  Else
    If (Trim$(OldValueStr) <> Trim$(Ctl.Text)) Then
      Local_Record.Name = Trim$(Ctl.Text)
      ''THROW DIRTY FLAG.
      'Call DirtyStatus_Throw
    End If
  End If
  Call Global_LostFocus(Ctl)
  'Call GenericStatus_Set("")
End Sub


