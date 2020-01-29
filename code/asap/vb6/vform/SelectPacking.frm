VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmSelectPacking 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Packing"
   ClientHeight    =   5250
   ClientLeft      =   1500
   ClientTop       =   2775
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Select Current Packing Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4020
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4110
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4020
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4710
      Width           =   3495
   End
   Begin Threed.SSFrame fraPackingDatabase 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6159
      _ExtentY        =   8700
      _StockProps     =   14
      Caption         =   "Original Database"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboSelectPacking 
         Appearance      =   0  'Flat
         Height          =   4275
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   450
         Width           =   3255
      End
   End
   Begin Threed.SSFrame fraPackingProperties 
      Height          =   3435
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Width           =   4125
      _Version        =   65536
      _ExtentX        =   7276
      _ExtentY        =   6059
      _StockProps     =   14
      Caption         =   "Packing Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtPackingNumericalProperties 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   0
         Left            =   2550
         TabIndex        =   1
         Top             =   990
         Width           =   1452
      End
      Begin VB.TextBox txtPackingNumericalProperties 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   2550
         TabIndex        =   2
         Top             =   1350
         Width           =   1452
      End
      Begin VB.TextBox txtPackingNumericalProperties 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   2550
         TabIndex        =   3
         Top             =   1710
         Width           =   1452
      End
      Begin VB.TextBox txtPackingNumericalProperties 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   3
         Left            =   2550
         TabIndex        =   4
         Top             =   2070
         Width           =   1452
      End
      Begin VB.TextBox txtPackingStringProperties 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   0
         Left            =   750
         TabIndex        =   0
         Top             =   390
         Width           =   3252
      End
      Begin VB.TextBox txtPackingStringProperties 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   1470
         TabIndex        =   5
         Top             =   2550
         Width           =   2532
      End
      Begin VB.TextBox txtPackingStringProperties 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   1470
         TabIndex        =   6
         Top             =   2910
         Width           =   2532
      End
      Begin VB.Label lblPackingProperties 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   150
         TabIndex        =   17
         Top             =   390
         Width           =   495
      End
      Begin VB.Label lblPackingProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Size"
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
         Left            =   150
         TabIndex        =   16
         Top             =   990
         Width           =   2175
      End
      Begin VB.Label lblPackingProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Factor"
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
         Left            =   150
         TabIndex        =   15
         Top             =   1350
         Width           =   2175
      End
      Begin VB.Label lblPackingProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sp. Surf. Area"
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
         Left            =   150
         TabIndex        =   14
         Top             =   1710
         Width           =   2175
      End
      Begin VB.Label lblPackingProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Crit. Surf. Tension"
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
         Left            =   150
         TabIndex        =   13
         Top             =   2070
         Width           =   2175
      End
      Begin VB.Label lblPackingProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
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
         Index           =   5
         Left            =   150
         TabIndex        =   12
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label lblPackingProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
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
         Index           =   6
         Left            =   150
         TabIndex        =   11
         Top             =   2910
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   4110
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Line Line2 
         X1              =   30
         X2              =   4110
         Y1              =   870
         Y2              =   870
      End
   End
   Begin VB.Menu mnuPackDatabaseMenu 
      Caption         =   "&Database"
      Begin VB.Menu mnuPackDatabase 
         Caption         =   "&Original Database"
         Index           =   0
      End
      Begin VB.Menu mnuPackDatabase 
         Caption         =   "&User-Modified Database"
         Index           =   1
      End
      Begin VB.Menu mnuPackDatabase 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPackDatabase 
         Caption         =   "O&ptions"
         Index           =   3
         Begin VB.Menu mnuPackDatabaseOptions 
            Caption         =   "&Remove Packing"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmSelectPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSelectPacking_Click()
    Dim i As Integer, j As Integer

    i = cboSelectPacking.ListIndex + 1

    If mnuPackDatabase(0).Checked = True Then

       If PackingChanged = False And PackingDatabaseSource = ORIGINALPACKINGDATABASE Then
          If Trim$(cboSelectPacking.Text) <> CurrentScreen.Packing.Name Then
             PackingChanged = True
          End If
       End If

       If PackingChanged = False And PackingDatabaseSource = USERMODIFIEDPACKINGDATABASE Then
          PackingChanged = True
       End If

       txtPackingStringProperties(0).Text = DatabasePacking(i).Name
       txtPackingNumericalProperties(0).Text = Format$(DatabasePacking(i).NominalSize, GetTheFormat(DatabasePacking(i).NominalSize))
       txtPackingNumericalProperties(1).Text = Format$(DatabasePacking(i).PackingFactor, GetTheFormat(DatabasePacking(i).PackingFactor))
       txtPackingNumericalProperties(2).Text = Format$(DatabasePacking(i).SpecificSurfaceArea, GetTheFormat(DatabasePacking(i).SpecificSurfaceArea))
       txtPackingNumericalProperties(3).Text = Format$(DatabasePacking(i).CriticalSurfaceTension, GetTheFormat(DatabasePacking(i).CriticalSurfaceTension))
       txtPackingStringProperties(1).Text = DatabasePacking(i).Material
       txtPackingStringProperties(2).Text = DatabasePacking(i).source

    Else
       PackingValuesChanged = False
       
       If PackingChanged = False And PackingDatabaseSource = USERMODIFIEDPACKINGDATABASE Then
          If Trim$(cboSelectPacking.Text) <> CurrentScreen.Packing.Name Then
             PackingChanged = True
          End If
       End If

       If PackingChanged = False And PackingDatabaseSource = ORIGINALPACKINGDATABASE Then
          PackingChanged = True
       End If
       
       txtPackingStringProperties(0).Text = UserPacking(i).Name
       txtPackingNumericalProperties(0).Text = Format$(UserPacking(i).NominalSize, GetTheFormat(UserPacking(i).NominalSize))
       txtPackingNumericalProperties(1).Text = Format$(UserPacking(i).PackingFactor, GetTheFormat(UserPacking(i).PackingFactor))
       txtPackingNumericalProperties(2).Text = Format$(UserPacking(i).SpecificSurfaceArea, GetTheFormat(UserPacking(i).SpecificSurfaceArea))
       txtPackingNumericalProperties(3).Text = Format$(UserPacking(i).CriticalSurfaceTension, GetTheFormat(UserPacking(i).CriticalSurfaceTension))
       txtPackingStringProperties(1).Text = UserPacking(i).Material
       txtPackingStringProperties(2).Text = UserPacking(i).source
    End If
End Sub

Private Sub cboSelectPacking_KeyPress(KeyAscii As Integer)
    Dim ComboText As String
    Dim i As Integer
    Dim msg As String

    If KeyAscii <> 13 Then Exit Sub   'Exit if Enter not pressed
    KeyAscii = 0

    ComboText = cboSelectPacking.Text

    'Check if Text Entered in Combo Box matches a
    'Packing type in the database and if so, select that
    'packing


    If mnuPackDatabase(0).Checked = True Then
       For i = 1 To NumPackingsInDatabase
           If Trim$(ComboText) = DatabasePacking(i).Name Then
              cboSelectPacking.ListIndex = i - 1
              cboSelectPacking_Click
              Exit Sub
           End If
       Next i
    Else
       For i = 1 To NumUserPackings
           If Trim$(ComboText) = UserPacking(i).Name Then
              cboSelectPacking.ListIndex = i - 1
              cboSelectPacking_Click
              Exit Sub
           End If
       Next i
    End If

    'If the packing type entered does not match, prepare
    'for input if we're in user-modified database

    If mnuPackDatabase(0).Checked = False Then
       txtPackingStringProperties(0).Text = ComboText
       txtPackingNumericalProperties(0).Text = Format$(DefaultPacking.NominalSize, GetTheFormat(DefaultPacking.NominalSize))
       txtPackingNumericalProperties(1).Text = Format$(DefaultPacking.PackingFactor, GetTheFormat(DefaultPacking.PackingFactor))
       txtPackingNumericalProperties(2).Text = Format$(DefaultPacking.SpecificSurfaceArea, GetTheFormat(DefaultPacking.SpecificSurfaceArea))
       txtPackingNumericalProperties(3).Text = Format$(DefaultPacking.CriticalSurfaceTension, GetTheFormat(DefaultPacking.CriticalSurfaceTension))
       txtPackingStringProperties(1).Text = DefaultPacking.Material
       txtPackingStringProperties(2).Text = "User"
      
       PackingValuesChanged = True
       PackingChanged = True
       txtPackingNumericalProperties(0).SetFocus
    Else
       For i = 1 To NumPackingsInDatabase
           If Trim$(ComboText) = Left$(DatabasePacking(i).Name, Len(ComboText)) Then
              cboSelectPacking.ListIndex = i - 1
              cboSelectPacking_Click
              Exit Sub
           End If
       Next i

       msg = "Packing with specified name not in original database.  "
       msg = msg + "Please specify another name or switch to the "
       msg = msg + "user-modified database if you would like to "
       msg = msg + "enter your own packing data."
       MsgBox msg, MB_ICONSTOP, "Packing Not Found"


       If cboSelectPacking.ListCount > 0 Then
          For i = 1 To NumPackingsInDatabase
              If UCase$(DatabasePacking(i).Name) = UCase$(scr1.Packing.Name) Then
                 cboSelectPacking.ListIndex = i - 1
                 Exit For
              Else
                 cboSelectPacking.ListIndex = 0
              End If
          Next i
          cboSelectPacking_Click
       End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim i As Integer
    
    If PackingDatabaseSource = ORIGINALPACKINGDATABASE Then
       If Not mnuPackDatabase(0).Checked Then
          mnuPackDatabase(0).Checked = True
          mnuPackDatabase(1).Checked = False
          mnuPackDatabase(3).Enabled = False
          cboSelectPacking.Clear
          For i = 1 To NumPackingsInDatabase
              cboSelectPacking.AddItem DatabasePacking(i).Name
          Next i
       End If
    Else
       If Not mnuPackDatabase(1).Checked Then
          mnuPackDatabase(1).Checked = True
          mnuPackDatabase(0).Checked = False
          mnuPackDatabase(3).Enabled = True
          cboSelectPacking.Clear
          For i = 1 To NumUserPackings
              cboSelectPacking.AddItem UserPacking(i).Name
          Next i
       End If
    End If

    mnuPackDatabase(0).Checked = True
    mnuPackDatabase(1).Checked = False
    
    PackingValuesChanged = False
    PackingChanged = False
    frmSelectPacking.Hide
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, msg As String
    Dim Response As Integer, Answer As Integer
    Dim MustOverwrite As Integer
    Dim NumToOverwrite As Integer  'Index of packing to overwrite (if necessary) in user-modified packing database
    Dim demostr$

'  DEMO STUFF ::TACK
    demostr$ = cboSelectPacking.Text
    If demomode_check_packing(demostr$) = 1 Then Exit Sub
'  END DEMO STUFF

    If Not PackingChanged Then
       frmSelectPacking.Hide
       Exit Sub
    End If

    If mnuPackDatabase(0).Checked = True Then
       PackingDatabaseSource = ORIGINALPACKINGDATABASE
    Else
       PackingDatabaseSource = USERMODIFIEDPACKINGDATABASE
    
        'Warn User if accepting this new contaminant would
        'overwrite an item already in database if any of
        'the values for packing properties have changed

        If PackingValuesChanged Then

           'Make sure user's name does not duplicate a
           'packing name in the original database and
           'if it does, exit the subroutine
           For i = 1 To NumPackingsInDatabase
               If UCase$(DatabasePacking(i).Name) = UCase$(txtPackingStringProperties(0).Text) Then
                  msg = "The name of your packing can not duplicate "
                  msg = msg + "the name of a packing in the original "
                  msg = msg + "packing database.  Please select a new "
                  msg = msg + "name for your packing."
                  MsgBox msg, MB_ICONSTOP, "Duplicate Name Error"
                  txtPackingStringProperties(0).SetFocus
                  Exit Sub
               End If
           Next i

           MustOverwrite = False
           For i = 1 To NumUserPackings
               If txtPackingStringProperties(0).Text = UserPacking(i).Name Then
                  msg = "Packing with name " & UserPacking(i).Name & Chr$(13)
                  msg = msg + "already exists in database." & Chr$(13) & Chr$(13)
                  msg = msg + "Do you wish to overwrite previous database values?"
                  Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "Warning")
                  If Response = IDNO Then
                     txtPackingStringProperties(0).SetFocus
                     Exit Sub
                  End If
                  MustOverwrite = True
                  NumToOverwrite = i
                  Exit For
               End If
           Next i
        End If
    End If

    CurrentScreen.Packing.Name = txtPackingStringProperties(0).Text
    CurrentScreen.Packing.NominalSize = CDbl(txtPackingNumericalProperties(0).Text)
    CurrentScreen.Packing.PackingFactor = CDbl(txtPackingNumericalProperties(1).Text)
    CurrentScreen.Packing.SpecificSurfaceArea = CDbl(txtPackingNumericalProperties(2).Text)
    CurrentScreen.Packing.CriticalSurfaceTension = CDbl(txtPackingNumericalProperties(3).Text)
    CurrentScreen.Packing.Material = txtPackingStringProperties(1).Text
    CurrentScreen.Packing.source = txtPackingStringProperties(2).Text
    CurrentScreen.Packing.SourceDatabase = PackingDatabaseSource
    CurrentScreen.Packing.ValChanged = True

    If Trim$(fraPackingDatabase.Caption) = "Original Database" Then
       CurrentScreen.Packing.UserInput = False
    Else
       CurrentScreen.Packing.UserInput = True
    End If

    If ScreenNumber = 1 Then
       frmPTADScreen1!lblPackingType.Caption = CurrentScreen.Packing.Name
    ElseIf ScreenNumber = 2 Then
       frmPTADScreen2!lblPackingType.Caption = CurrentScreen.Packing.Name
    End If

    'Add new packing to user-modified database if it has been modified

    If (PackingDatabaseSource = USERMODIFIEDPACKINGDATABASE) And (PackingValuesChanged = True) Then
       Select Case MustOverwrite
          Case True
             i = NumToOverwrite
          Case False
             msg = CurrentScreen.Packing.Name & " has been added to the "
             msg = msg + "user-modified packing database."
             MsgBox msg, MB_ICONINFORMATION, ""
             i = NumUserPackings + 1
             NumUserPackings = i
             frmSelectPacking.cboSelectPacking.AddItem CurrentScreen.Packing.Name
       End Select

       UserPacking(i).Name = CurrentScreen.Packing.Name
       UserPacking(i).NominalSize = CurrentScreen.Packing.NominalSize
       UserPacking(i).PackingFactor = CurrentScreen.Packing.PackingFactor
       UserPacking(i).SpecificSurfaceArea = CurrentScreen.Packing.SpecificSurfaceArea
       UserPacking(i).CriticalSurfaceTension = CurrentScreen.Packing.CriticalSurfaceTension
       UserPacking(i).Material = CurrentScreen.Packing.Material
       UserPacking(i).source = CurrentScreen.Packing.source
       UserPacking(i).SourceDatabase = PackingDatabaseSource
       UserPacking(i).ValChanged = False
       UserPacking(i).UserInput = CurrentScreen.Packing.UserInput

       'Give user the option to permanently update the user-modified database by writing the changes to disk
       WriteUserPackingDB

    End If

    mnuPackDatabase(0).Checked = True
    mnuPackDatabase(1).Checked = False
    
    PackingChanged = False
    ShownPackingProperties = False
    frmSelectPacking.Hide

End Sub

Private Sub Form_Activate()
  Call CenterThisForm(Me)
End Sub

Private Sub Form_Load()

  Call CenterThisForm(Me)

'  DEMO STUFF STOP THE UER FROM CHANGING PROPERTIES OF PACKING MATERIALS::TACK
    If DemoMode% Then
        frmSelectPacking!txtPackingStringProperties(0).Enabled = False
        frmSelectPacking!txtPackingStringProperties(1).Enabled = False
        frmSelectPacking!txtPackingStringProperties(2).Enabled = False
        frmSelectPacking!txtPackingNumericalProperties(0).Enabled = False
        frmSelectPacking!txtPackingNumericalProperties(1).Enabled = False
        frmSelectPacking!txtPackingNumericalProperties(2).Enabled = False
        frmSelectPacking!txtPackingNumericalProperties(3).Enabled = False
    End If
'  END DEMO STUFF

    frmSelectPacking.WindowState = 0
    

    Call LabelsSelectPackingPropertiesSI

    PackingChanged = False
    PackingValuesChanged = False
End Sub

Private Sub mnuPackDatabase_Click(Index As Integer)
    Dim i As Integer, CurrPackingIndex As Integer

    CurrPackingIndex = -1

    Select Case Index
       Case 0    'Original Database
            If Not mnuPackDatabase(0).Checked Then
               mnuPackDatabase(0).Checked = True
               fraPackingDatabase.Caption = "Original Database"
               mnuPackDatabase(1).Checked = False
               mnuPackDatabase(3).Enabled = False
               cboSelectPacking.Clear
               For i = 1 To NumPackingsInDatabase
                   cboSelectPacking.AddItem DatabasePacking(i).Name
                   If DatabasePacking(i).Name = CurrentScreen.Packing.Name Then CurrPackingIndex = i
               Next i
               If CurrPackingIndex = -1 Then   'Currently Selected Packing not in database
                  PackingChanged = True
                  cboSelectPacking.ListIndex = 0
                  cboSelectPacking_Click
                  
               Else
                  PackingChanged = False
                  cboSelectPacking.ListIndex = CurrPackingIndex - 1
                  cboSelectPacking_Click
                  
               End If
            End If
       Case 1    'User-Modified Database
            If Not mnuPackDatabase(1).Checked Then
               mnuPackDatabase(0).Checked = False
               mnuPackDatabase(1).Checked = True
               mnuPackDatabase(3).Enabled = True
               fraPackingDatabase.Caption = "User-Modified Database"
               cboSelectPacking.Clear
               
               For i = 1 To NumUserPackings
                   cboSelectPacking.AddItem UserPacking(i).Name
                   If UserPacking(i).Name = CurrentScreen.Packing.Name Then CurrPackingIndex = i
               Next i
            End If
               If CurrPackingIndex = -1 Then   'Currently Selected Packing not in database
                  PackingChanged = True
                  cboSelectPacking.ListIndex = 0
                  cboSelectPacking_Click
                  
               Else
                  PackingChanged = False
                  cboSelectPacking.ListIndex = CurrPackingIndex - 1
                  cboSelectPacking_Click
                  
               End If

    End Select
End Sub

Private Sub mnuPackDatabaseOptions_Click(Index As Integer)
    Dim msg As String, i As Integer
    Dim NumToRemove As Integer
    Dim Response As Integer

    Select Case Index
       Case 0    'Remove selected packing from database

          NumToRemove = -1
          For i = 1 To NumUserPackings
              If Trim$(cboSelectPacking.Text) = UserPacking(i).Name Then
                 NumToRemove = i
                 If (Trim$(cboSelectPacking.Text) = CurrentScreen.Packing.Name) And (PackingDatabaseSource = USERMODIFIEDPACKINGDATABASE) Then
                    MsgBox "You are not allowed to remove the currently selected packing type from the database.  Recommendation: First, select a different packing type, and second, remove this packing type.", MB_ICONSTOP, "Error"
                    Exit Sub
                 End If
                 Exit For
              End If
          Next i

          If NumToRemove = -1 Then
             MsgBox "The packing selected for removal (" & Trim$(cboSelectPacking.Text) & ") is not an item in the user-modified database.", MB_ICONSTOP, "Error"
             Exit Sub
          End If

          msg = "Do you really want to remove "
          msg = msg + cboSelectPacking.Text + " from the "
          msg = msg + "User-Modified Packing Database?"
          Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "")

          If Response = IDNO Then Exit Sub
          cboSelectPacking.RemoveItem (NumToRemove - 1)
          For i = NumToRemove To NumUserPackings - 1
              UserPacking(i).Name = UserPacking(i + 1).Name
              UserPacking(i).NominalSize = UserPacking(i + 1).NominalSize
              UserPacking(i).PackingFactor = UserPacking(i + 1).PackingFactor
              UserPacking(i).SpecificSurfaceArea = UserPacking(i + 1).SpecificSurfaceArea
              UserPacking(i).CriticalSurfaceTension = UserPacking(i + 1).CriticalSurfaceTension
              UserPacking(i).Material = UserPacking(i + 1).Material
              UserPacking(i).source = UserPacking(i + 1).source
              UserPacking(i).ValChanged = UserPacking(i + 1).ValChanged
              UserPacking(i).UserInput = UserPacking(i + 1).UserInput
          Next i

          NumUserPackings = NumUserPackings - 1

          'Give user the option to permanently update the user-modified database by writing the changes to disk

          'NOTE (4/9/98): These two confirmation requests are handled
          'right in the WriteUserPackingDB() routine, therefore they are commented out.  EJO.
          'msg = "Would you like to PERMANENTLY update "
          'msg = msg + "the changes in the user-modified database "
          'msg = msg + "to disk?"
          'Response = MsgBox(msg, MB_ICONQUESTION + MB_YESNO, "")
          'If Response = IDYES Then
          '   Response = MsgBox("This can not be undone.", MB_OKCANCEL + MB_ICONEXCLAMATION, "Warning")
          '   If Response = IDOK Then
                Call WriteUserPackingDB
          '   End If
          'End If

          If cboSelectPacking.ListCount > 0 Then
             cboSelectPacking.ListIndex = 0
             cboSelectPacking_Click
          End If

     End Select
          
End Sub

Private Sub txtPackingNumericalProperties_Change(Index As Integer)
    Dim i As Integer

    cmdOK.Enabled = True
    
    If txtPackingStringProperties(0).Text = "" Then cmdOK.Enabled = False

    If cmdOK.Enabled Then
       For i = 0 To 3
           If Val(txtPackingNumericalProperties(i).Text) <= 0# Then
              cmdOK.Enabled = False
           End If
       Next i
    End If

End Sub

Private Sub txtPackingNumericalProperties_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtPackingNumericalProperties(Index), Temp_Text)
End Sub

Private Sub txtPackingNumericalProperties_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim msg As String

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    If mnuPackDatabase(0).Checked = True Then
       KeyAscii = 0   'User does not have option to modify values in the original database
       msg = "Values in the original database can not be modified.  "
       msg = msg + "If you wish to enter your own values, switch to the "
       msg = msg + "User-Modified Database in the Database menu."
       MsgBox msg, 16, "Warning"
    Else
       Call NumberCheck(KeyAscii)
    End If

End Sub

Private Sub txtPackingNumericalProperties_LostFocus(Index As Integer)
    Dim ValueChanged As Integer
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtPackingNumericalProperties(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

    Call TextHandleError(IsError, txtPackingNumericalProperties(Index), Temp_Text)

    If Not IsError Then
       Call TextNumberChanged(ValueChanged, txtPackingNumericalProperties(Index), Temp_Text)

       If ValueChanged Then
          PackingChanged = True
          PackingValuesChanged = True
       End If
    End If
  Call LostFocus_Handle(Me, txtPackingNumericalProperties(Index), flag_ok)


End Sub

Private Sub txtPackingStringProperties_Change(Index As Integer)
    Dim i As Integer

Debug.Print "Change: `" & txtPackingStringProperties(Index).Text & "`"

    If Index <> 0 Then Exit Sub   'The user is not required to enter values for Material or Source to execute program
    'txtPackingStringProperties(Index).Text = Trim$(txtPackingStringProperties(Index).Text)
    cmdOK.Enabled = True
    If txtPackingStringProperties(Index).Text = "" Then cmdOK.Enabled = False
    If cmdOK.Enabled Then
       For i = 0 To 3
           If Val(txtPackingNumericalProperties(i).Text) <= 0# Then
              cmdOK.Enabled = False
           End If
       Next i
    End If

End Sub

Private Sub txtPackingStringProperties_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtPackingStringProperties(Index), Temp_Text)
End Sub

Private Sub txtPackingStringProperties_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim msg As String

Debug.Print KeyAscii

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    If mnuPackDatabase(0).Checked = True Then
       KeyAscii = 0   'User does not have option to modify values in the original database
       msg = "Values in the original database can not be modified.  "
       msg = msg + "If you wish to enter your own values, switch to the "
       msg = msg + "User-Modified Database in the Database menu."
       MsgBox msg, 16, "Warning"
    End If

End Sub

Private Sub txtPackingStringProperties_LostFocus(Index As Integer)
    Dim ValueChanged As Integer
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtPackingStringProperties(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True


    txtPackingStringProperties(Index).Text = Trim$(txtPackingStringProperties(Index).Text)
    Call TextStringChanged(ValueChanged, txtPackingStringProperties(Index), Temp_Text)
    If ValueChanged Then
       PackingValuesChanged = True
       PackingChanged = True
    End If
  Call LostFocus_Handle(Me, txtPackingStringProperties(Index), flag_ok)


End Sub


