VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmStringS 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search for String..."
   ClientHeight    =   1920
   ClientLeft      =   2490
   ClientTop       =   3300
   ClientWidth     =   5040
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
   ScaleHeight     =   1920
   ScaleWidth      =   5040
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3612
   End
   Begin Threed.SSCommand cmdStart 
      Height          =   372
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Start searching..."
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
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the string to find:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3012
   End
End
Attribute VB_Name = "frmStringS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub cmdCancel_Click()
    Find_String = ""
    Unload Me
End Sub

Private Sub cmdStart_Click()
   Find_String = Trim$(txtFind)
   If Find_String = "" Then
     MsgBox "The string is an empty string", 48, "Error"
   Else
     Unload Me
   End If
End Sub

Private Sub Form_Load()

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
      Move contam_prop_form.Left + (contam_prop_form.Width / 2) - (frmStringS.Width / 2), contam_prop_form.Top + (contam_prop_form.Height / 2) - (frmStringS.Height / 2)
    
    End If

  txtFind = Find_String
  txtFind.SelLength = Len(Find_String)
  
End Sub

Private Sub txtFind_GotFocus()
  Call GotFocus_Handle(Me, txtFind, Temp_Text)
End Sub

Private Sub txtFind_KeyPress(keyascii As Integer)
    If keyascii = 13 Then '<Return> pressed
       cmdStart_Click
    End If
End Sub

Private Sub txtFind_LostFocus()
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtFind)) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtFind, flag_ok)

End Sub

