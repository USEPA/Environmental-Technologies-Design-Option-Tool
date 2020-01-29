VERSION 5.00
Begin VB.Form frmGoAway 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "{Caller-Defined Caption}"
   ClientHeight    =   3210
   ClientLeft      =   2115
   ClientTop       =   3945
   ClientWidth     =   7290
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3210
   ScaleWidth      =   7290
   Begin VB.CheckBox chkDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "{Caller-Defined Text}"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2700
      Width           =   5415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6960
      ScaleHeight     =   405
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   -60
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   6915
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "{Caller-Defined Text}"
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   6615
      End
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmGoAway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmGoAway_ParentForm As Form
Dim frmGoAway_Caption As String
Dim frmGoAway_Text As String
Dim frmGoAway_CheckText As String
Dim frmGoAway_CheckValue As Integer





Const frmGoAway_declarations_end = True


Public Sub frmGoAway_Run( _
    INPUT_frmGoAway_ParentForm As Form, _
    INPUT_frmGoAway_Caption As String, _
    INPUT_frmGoAway_Text As String, _
    INPUT_frmGoAway_CheckText As String, _
    INPUTOUTPUT_frmGoAway_CheckValue As Integer)
  Set frmGoAway_ParentForm = INPUT_frmGoAway_ParentForm
  frmGoAway_Caption = INPUT_frmGoAway_Caption
  frmGoAway_Text = INPUT_frmGoAway_Text
  frmGoAway_CheckText = INPUT_frmGoAway_CheckText
  frmGoAway_CheckValue = INPUTOUTPUT_frmGoAway_CheckValue
  frmGoAway.Show 1
  INPUTOUTPUT_frmGoAway_CheckValue = frmGoAway_CheckValue
End Sub


Private Sub cmdOK_Click()
  If (chkDisplay.Value = True) Then
    frmGoAway_CheckValue = 1
  Else
    frmGoAway_CheckValue = 0
  End If
  Unload Me
End Sub
Private Sub Form_Load()
Dim ht As Double
  Me.Height = 3615
  Me.Width = 7410
  Me.Caption = frmGoAway_Caption
  Label1.Caption = frmGoAway_Text
  chkDisplay.Caption = frmGoAway_CheckText
  If (frmGoAway_CheckValue = 1) Then
    chkDisplay.Value = True
  Else
    chkDisplay.Value = False
  End If
  Call CenterOnForm(Me, frmGoAway_ParentForm)
  ht = Picture1.TextHeight(frmGoAway_Text)
  Label1.Height = ht * 1.05
  Frame1.Height = ht * 1.2
End Sub
Private Sub Form_Unload(Cancel As Integer)
  'contam_prop_form.SetFocus
End Sub

