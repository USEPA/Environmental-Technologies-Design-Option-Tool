VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmNewName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "{use_title}"
   ClientHeight    =   1575
   ClientLeft      =   1755
   ClientTop       =   7380
   ClientWidth     =   5685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1575
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtdata 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Text            =   "txtdata"
      Top             =   750
      Width           =   5505
   End
   Begin Threed.SSCommand Button 
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   556
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
   Begin Threed.SSCommand Button 
      Height          =   315
      Index           =   1
      Left            =   4290
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   556
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
   Begin VB.Label lblInstructions 
      Caption         =   "lblInstructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   5505
   End
End
Attribute VB_Name = "frmNewName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim return_text As String
Dim USER_HIT_CANCEL As Boolean




Const frmNewName_declarations_end = True


Public Function frmNewName_GetName( _
    Use_Title As String, _
    use_label As String, _
    use_default As String, _
    is_aborted As Boolean) As String
  Load frmNewName
  Me.Caption = Use_Title
  lblInstructions.Caption = use_label
  txtData.Text = use_default
  frmNewName.Show 1
  is_aborted = USER_HIT_CANCEL
  frmNewName_GetName = return_text
End Function


Private Sub Button_Click(Index As Integer)
  Select Case Index
    Case 0:     'OK
      return_text = Trim$(txtData.Text)
      If (return_text = "") Then
        Call Show_Error("You must enter a non-blank string as a name.")
        Exit Sub
      End If
      USER_HIT_CANCEL = False
      Unload Me
    Case 1:     'Cancel
      USER_HIT_CANCEL = True
      Unload Me
  End Select
End Sub


Private Sub Form_Load()
  Me.Height = 1965
  Me.Width = 5805
  Call CenterOnForm(Me, frmMain)
End Sub


Private Sub txtData_GotFocus()
Dim txtCtl As Control
Set txtCtl = txtData
  'Call DisplayDataEntryError
  Call Global_GotFocus(txtCtl)
End Sub
Private Sub txtData_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call Button_Click(0)
  End If
  'keyascii = Global_TextKeyPress(keyascii)
'  If (KeyAscii = 13) Then SendKeys "{TAB}", True
End Sub
Private Sub txtData_LostFocus()
Dim txtCtl As Control
Set txtCtl = txtData
  Call Global_LostFocus(txtCtl)
End Sub

