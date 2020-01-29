VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEditIsothermCAS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "{me.caption}"
   ClientHeight    =   2985
   ClientLeft      =   2085
   ClientTop       =   4650
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7530
   Begin Threed.SSCheck chkData 
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   8
      Top             =   1560
      Width           =   6825
      _Version        =   65536
      _ExtentX        =   12039
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "chkData(0)"
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
   Begin VB.TextBox txtData 
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
      Height          =   315
      Index           =   3
      Left            =   2880
      TabIndex        =   6
      Text            =   "txtData(3)"
      Top             =   1170
      Width           =   4500
   End
   Begin VB.TextBox txtData 
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
      Height          =   315
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Text            =   "txtData(2)"
      Top             =   810
      Width           =   2385
   End
   Begin VB.TextBox txtData 
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
      Height          =   315
      Index           =   1
      Left            =   2880
      TabIndex        =   2
      Text            =   "txtData(1)"
      Top             =   450
      Width           =   4500
   End
   Begin VB.TextBox txtData 
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
      Height          =   315
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Text            =   "txtData(0)"
      Top             =   90
      Width           =   2385
   End
   Begin Threed.SSCheck chkData 
      Height          =   285
      Index           =   1
      Left            =   540
      TabIndex        =   9
      Top             =   1920
      Width           =   6825
      _Version        =   65536
      _ExtentX        =   12039
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "chkData(1)"
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
   Begin Threed.SSCommand cmdSaveCancel 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2370
      Width           =   2000
      _Version        =   65536
      _ExtentX        =   3528
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "cmdSaveCancel(0)"
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
   Begin Threed.SSCommand cmdSaveCancel 
      Height          =   495
      Index           =   1
      Left            =   5880
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2370
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
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
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "lblDesc(3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   2715
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "lblDesc(2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2715
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "lblDesc(1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2715
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "lblDesc(0)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2715
   End
End
Attribute VB_Name = "frmEditIsothermCAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean


Dim Use_Title As String
Dim Use_TextLabel_1a As String
Dim Use_TextLabel_1b As String
Dim Use_TextLabel_2a As String
Dim Use_TextLabel_2b As String
Dim Use_OptionLabel_1 As String
Dim Use_OptionLabel_2 As String
Dim Use_OK_Caption As String

Dim VarText_1a As String
Dim VarText_1b As String
Dim VarText_2a As String
Dim VarText_2b As String
Dim VarBool_1 As Boolean
Dim VarBool_2 As Boolean


Const frmEditIsothermCAS_declarations_end = True


Public Sub frmEditIsothermCAS_Run( _
    INPUT_Use_Title As String, _
    INPUT_Use_TextLabel_1a As String, _
    INPUT_Use_TextLabel_1b As String, _
    INPUT_Use_TextLabel_2a As String, _
    INPUT_Use_TextLabel_2b As String, _
    INPUT_Use_OptionLabel_1 As String, _
    INPUT_Use_OptionLabel_2 As String, _
    INPUT_Use_OK_Caption As String, _
    OUTPUT_USER_HIT_CANCEL As Boolean, _
    IO_VarText_1a As String, _
    IO_VarText_1b As String, _
    IO_VarText_2a As String, _
    IO_VarText_2b As String, _
    IO_VarBool_1 As Boolean, _
    IO_VarBool_2 As Boolean)
  Use_Title = INPUT_Use_Title
  Use_TextLabel_1a = INPUT_Use_TextLabel_1a
  Use_TextLabel_1b = INPUT_Use_TextLabel_1b
  Use_TextLabel_2a = INPUT_Use_TextLabel_2a
  Use_TextLabel_2b = INPUT_Use_TextLabel_2b
  Use_OptionLabel_1 = INPUT_Use_OptionLabel_1
  Use_OptionLabel_2 = INPUT_Use_OptionLabel_2
  Use_OK_Caption = INPUT_Use_OK_Caption
  VarText_1a = IO_VarText_1a
  VarText_1b = IO_VarText_1b
  VarText_2a = IO_VarText_2a
  VarText_2b = IO_VarText_2b
  VarBool_1 = IO_VarBool_1
  VarBool_2 = IO_VarBool_2
  frmEditIsothermCAS.Show 1
  OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
  If (USER_HIT_OK) Then
    IO_VarText_1a = VarText_1a
    IO_VarText_1b = VarText_1b
    IO_VarText_2a = VarText_2a
    IO_VarText_2b = VarText_2b
    IO_VarBool_1 = VarBool_1
    IO_VarBool_2 = VarBool_2
  End If
End Sub






Private Sub cmdSaveCancel_Click(Index As Integer)
  Select Case Index
    Case 0:     'OK.
      'TRANSFER DATA FROM CONTROLS TO MEMORY.
      VarText_1a = txtData(0).Text
      VarText_1b = txtData(1).Text
      VarText_2a = txtData(2).Text
      VarText_2b = txtData(3).Text
      VarBool_1 = chkData(0).Value
      VarBool_2 = chkData(1).Value
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
    Case 1:     'CANCEL.
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub Form_Load()
  'MISC INITS.
  Me.Caption = Use_Title
  cmdSaveCancel(0).Caption = Use_OK_Caption
  If (Left$(Use_TextLabel_1a, 1) = "^") Then
    Use_TextLabel_1a = Right$(Use_TextLabel_1a, Len(Use_TextLabel_1a) - 1)
    txtData(0).Locked = True
    txtData(0).BackColor = QBColor(7)
  End If
  If (Left$(Use_TextLabel_1b, 1) = "^") Then
    Use_TextLabel_1b = Right$(Use_TextLabel_1b, Len(Use_TextLabel_1b) - 1)
    txtData(1).Locked = True
    txtData(1).BackColor = QBColor(7)
  End If
  lblDesc(0).Caption = Use_TextLabel_1a
  lblDesc(1).Caption = Use_TextLabel_1b
  lblDesc(2).Caption = Use_TextLabel_2a
  lblDesc(3).Caption = Use_TextLabel_2b
  chkData(0).Caption = Use_OptionLabel_1
  chkData(1).Caption = Use_OptionLabel_2
  txtData(0).Text = VarText_1a
  txtData(1).Text = VarText_1b
  txtData(2).Text = VarText_2a
  txtData(3).Text = VarText_2b
  chkData(0).Value = VarBool_1
  chkData(1).Value = VarBool_2
  If (lblDesc(2).Caption = "*") Then
    lblDesc(2).Visible = False
    lblDesc(3).Visible = False
    txtData(2).Visible = False
    txtData(3).Visible = False
  End If
  If (chkData(0).Caption = "*") Then
    chkData(0).Visible = False
    chkData(1).Visible = False
  End If
  Me.Height = 3390
  Me.Width = 7650
  Call CenterOnForm(Me, frmEditAdsorber)
End Sub






Private Sub txtData_GotFocus(Index As Integer)
Dim txtCtl As Control
Set txtCtl = txtData(Index)
  If (txtCtl.Locked) Then Exit Sub
  Call Global_GotFocus(txtCtl)
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtData_LostFocus(Index As Integer)
Dim txtCtl As Control
Set txtCtl = txtData(Index)
  If (txtCtl.Locked) Then Exit Sub
  Call Global_LostFocus(txtCtl)
End Sub



