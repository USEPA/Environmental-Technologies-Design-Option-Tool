VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFileNote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Note"
   ClientHeight    =   4680
   ClientLeft      =   2280
   ClientTop       =   4560
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7500
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   7680
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   8
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   3690
      Width           =   1455
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   90
      Locked          =   -1  'True
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FileNote.frx":0000
      Top             =   510
      Width           =   7305
   End
   Begin Threed.SSCommand cmdButton 
      Height          =   495
      Index           =   0
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4080
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Delete"
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
   Begin Threed.SSCommand cmdButton 
      Height          =   495
      Index           =   1
      Left            =   1020
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4080
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Edit"
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
   Begin Threed.SSCommand cmdButton 
      Height          =   495
      Index           =   2
      Left            =   1950
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4080
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Close"
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
   Begin Threed.SSCommand cmdButton 
      Height          =   495
      Index           =   3
      Left            =   5100
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4080
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Save"
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
   Begin Threed.SSCommand cmdButton 
      Height          =   495
      Index           =   4
      Left            =   6030
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "C&ancel Edit"
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
   Begin Threed.SSCommand cmdButton 
      Height          =   495
      Index           =   5
      Left            =   3420
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Insert Date/Time"
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
   Begin VB.Label lblInstructions 
      Caption         =   "You may enter up to 500 characters of text.  Line breaks are acceptable."
      Height          =   405
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   5955
   End
End
Attribute VB_Name = "frmFileNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NoteText As String
Dim Rollback_NoteText As String
Dim RaiseDirtyFlag As Boolean
Dim Rollback_RaiseDirtyFlag As Boolean

Dim FORM_MODE As Integer
Const FORM_MODE_VIEW = 1
Const FORM_MODE_EDIT = 2




Const frmFileNote_declarations_end = True


Sub frmFileNote_Run( _
    IO_NoteText As String, _
    O_RaiseDirtyFlag As Boolean)
  NoteText = IO_NoteText
  ''''NoteText = Parser_ReplaceStrings(NoteText, Chr$(255), Chr$(13) & Chr$(10))
  frmFileNote.Show 1
  If (RaiseDirtyFlag) Then
    IO_NoteText = NoteText
    ''''MsgBox NoteText
    ''''IO_NoteText = Parser_ReplaceStrings(IO_NoteText, Chr$(13) & Chr$(10), Chr$(255))
  End If
  O_RaiseDirtyFlag = RaiseDirtyFlag
End Sub


Sub frmFileNote_Refresh()
  Select Case FORM_MODE
    Case FORM_MODE_VIEW:
      cmdButton(0).Enabled = True
      cmdButton(1).Enabled = True
      cmdButton(2).Enabled = True
      cmdButton(3).Enabled = False
      cmdButton(4).Enabled = False
      cmdButton(5).Enabled = False
      txtData.Locked = True
      lblInstructions.Caption = ""
    Case FORM_MODE_EDIT:
      cmdButton(0).Enabled = False
      cmdButton(1).Enabled = False
      cmdButton(2).Enabled = False
      cmdButton(3).Enabled = True
      cmdButton(4).Enabled = True
      cmdButton(5).Enabled = True
      txtData.Locked = False
      lblInstructions.Caption = _
          "You may enter up to 500 characters of text.  " & _
          "Line breaks are acceptable."
  End Select
  txtData.Text = NoteText
End Sub


Private Sub cmdButton_Click(Index As Integer)
Dim RetVal As Integer
  Select Case Index
    Case 0:   'DELETE.
      RetVal = MsgBox("Are you sure you want to delete this " & _
          "file note ?", vbQuestion + vbYesNo, _
          AppName_For_Display_Short & " : Delete File Note ?")
      If (RetVal = vbNo) Then Exit Sub
      NoteText = ""
      RaiseDirtyFlag = True
      Call frmFileNote_Refresh
    Case 1:   'EDIT.
      'STORE OLD VALUE IF USER CANCELS EDIT.
      Rollback_NoteText = NoteText
      Rollback_RaiseDirtyFlag = RaiseDirtyFlag
      FORM_MODE = FORM_MODE_EDIT
      Call frmFileNote_Refresh
    Case 2:   'CLOSE.
      Unload Me
      Exit Sub
    Case 3:   'SAVE.
      FORM_MODE = FORM_MODE_VIEW
      Call frmFileNote_Refresh
''temp
'Open "c:\test.out" For Output As #1
'Dim i As Integer
'For i = 1 To Len(NoteText)
'  Print #1, Asc(Mid$(NoteText, i, 1))
'Next i
'Close #1
    Case 4:   'CANCEL EDIT.
      FORM_MODE = FORM_MODE_VIEW
      NoteText = Rollback_NoteText
      RaiseDirtyFlag = Rollback_RaiseDirtyFlag
      Call frmFileNote_Refresh
    Case 5:   'INSERT DATE/TIME.
      txtData.Text = txtData.Text & Now
      Call txtData_LostFocus
  End Select
End Sub


Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
  'MISC INITS.
  Call CenterOnForm(Me, frmMain)
  FORM_MODE = FORM_MODE_VIEW
  Call frmFileNote_Refresh
  RaiseDirtyFlag = False
End Sub


Private Sub txtData_GotFocus()
Dim Ctl As Control
Set Ctl = txtData
  Call Global_GotFocus(Ctl)
  'FORCE BACKGROUND COLOR BACK TO WHITE.
  Ctl.BackColor = RGB(255, 255, 255)
End Sub
Private Sub txtData_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_MultilineTextKeyPress(KeyAscii)
End Sub
Private Sub txtData_LostFocus()
'Dim NewValue_Okay As Integer
'Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtData
'Dim Val_Low As Double
'Dim Val_High As Double
'Dim Raise_Dirty_Flag As Boolean
'Dim Too_Small As Integer
Dim OldValueStr As String
  'HANDLE STRING FIELDS.
  OldValueStr = Trim$(NoteText)
  'NOTE: ZERO-LENGTH STRINGS ARE ALLOWED.
  If (Trim$(OldValueStr) <> Trim$(Ctl.Text)) Then
    NoteText = Trim$(Ctl.Text)
    RaiseDirtyFlag = True
  End If
  Call Global_LostFocus(Ctl)
  Call frmFileNote_Refresh
  'Call GenericStatus_Set("")
  Exit Sub
'  End If
'  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
'  Select Case index
'    Case 0: Val_Low = 1E-20: Val_High = 1E+20
'    Case 1: Val_Low = 0#: Val_High = 1E+20
'    Case 2: Val_Low = 0#: Val_High = 1E+20
'    Case 3: Val_Low = 1E-20: Val_High = 1E+20
'    Case 4: Val_Low = 1E-20: Val_High = 1E+20
'    Case 5: Val_Low = 0#: Val_High = 1E+20
'    Case 6: Val_Low = 0#: Val_High = 1E+20
'  End Select
'  NewValue_Okay = False
'  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
'    NewValue_Okay = True
'  End If
'  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
'  If (NewValue_Okay) Then
'    If (Raise_Dirty_Flag) Then
'      'STORE TO MEMORY.
'      Select Case index
'        Case 0:         'FREUNDLICH K.
'          frmEditIsothermData_Record.k = NewValue
'        Case 1:         'MINIMUM CONCENTRATION.
'          frmEditIsothermData_Record.Cmin = NewValue
'        Case 2:         'MINIMUM pH.
'          frmEditIsothermData_Record.pHmin = NewValue
'        Case 3:         'TEMPERATURE.
'          frmEditIsothermData_Record.Tmin = NewValue
'        Case 4:         'FREUNDLICH 1/n.
'          frmEditIsothermData_Record.OneOverN = NewValue
'        Case 5:         'MAXIMUM CONCENTRATION.
'          frmEditIsothermData_Record.Cmax = NewValue
'        Case 6:         'MAXIMUM pH.
'          frmEditIsothermData_Record.pHmax = NewValue
'      End Select
'      'RAISE DIRTY FLAG IF NECESSARY.
'      If (Raise_Dirty_Flag) Then
'        ''THROW DIRTY FLAG.
'        'Call frmCompoProp_DirtyStatus_Throw
'      End If
'      'REFRESH WINDOW.
'      Call frmEditIsothermData_Refresh
'    End If
'  End If
End Sub



