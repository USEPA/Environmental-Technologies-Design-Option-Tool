VERSION 5.00
Begin VB.Form frmNewName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "{use_title}"
   ClientHeight    =   1215
   ClientLeft      =   2025
   ClientTop       =   3495
   ClientWidth     =   5685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1215
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
      Top             =   390
      Width           =   5505
   End
   Begin VB.CommandButton button 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4260
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton button 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   840
      Width           =   1335
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
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   5505
   End
End
Attribute VB_Name = "frmNewName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim return_text As String
Dim did_abort As Integer

Public Function frmNewName_GetName( _
    use_title As String, _
    use_label As String, _
    use_default As String, _
    is_aborted As Integer) As String
    
  Load frmNewName
  Me.Caption = use_title
  lblInstructions.Caption = use_label
  txtdata.Text = use_default
  frmNewName.Show 1
  is_aborted = did_abort
  frmNewName_GetName = return_text
    
End Function


Private Sub Button_Click(Index As Integer)

  Select Case Index
    Case 0:     'OK
      return_text = Trim$(txtdata.Text)
      If (return_text = "") Then
        Call Show_Error("You must enter a non-blank string as a name.")
        Exit Sub
      End If
      did_abort = False
      Unload Me
    Case 1:     'Cancel
      did_abort = True
      Unload Me
  End Select
  
End Sub


Private Sub Form_Load()

  Call CenterOnForm(Me, frmMain)
  
End Sub



Private Sub txtdata_GotFocus()
Dim txtctl As Control
Set txtctl = txtdata

Call DisplayDataEntryError
Call Global_GotFocus(txtctl)

End Sub


Private Sub txtdata_KeyPress(KeyAscii As Integer)

  If (KeyAscii = 13) Then
    Call Button_Click(0)
  End If
  
  'keyascii = Global_TextKeyPress(keyascii)
  
'  If (KeyAscii = 13) Then SendKeys "{TAB}", True
  
End Sub


Private Sub txtdata_LostFocus()
Dim txtctl As Control
Set txtctl = txtdata
  
  Call Global_LostFocus(txtctl)
  
End Sub

