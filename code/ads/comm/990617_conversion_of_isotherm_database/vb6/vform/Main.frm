VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   3780
   ClientLeft      =   765
   ClientTop       =   1680
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   13650
   Begin VB.TextBox txtName 
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   8
      Text            =   "testing"
      Top             =   1200
      Width           =   11000
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Text            =   "testing"
      Top             =   480
      Width           =   11000
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Text            =   "testing"
      Top             =   840
      Width           =   11000
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Text            =   "testing"
      Top             =   120
      Width           =   11000
   End
   Begin VB.CommandButton cmdButtons 
      Caption         =   "Convert Isotherms Table"
      Height          =   345
      Index           =   1
      Left            =   7080
      TabIndex        =   1
      Top             =   3000
      Width           =   2565
   End
   Begin VB.CommandButton cmdButtons 
      Caption         =   "Scan Isotherms Table"
      Height          =   345
      Index           =   0
      Left            =   7080
      TabIndex        =   0
      Top             =   2520
      Width           =   2565
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      Caption         =   "Merge file:"
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   9
      Top             =   1245
      Width           =   1845
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      Caption         =   "Database #1 name:"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   525
      Width           =   1845
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      Caption         =   "Database #2 name:"
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   885
      Width           =   1845
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      Caption         =   "Log directory:"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   165
      Width           =   1845
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Const frmMain_decl_end = True


Sub Launch_Notepad_Log_TXT()
  Call Launch_Notepad(frmMain.txtName(0).Text & "\log.txt")
End Sub


Private Sub cmdButtons_Click(Index As Integer)
  Select Case Index
    Case 0:     'Scan Isotherms Table.
      Call Scan_Isotherms_Table
    Case 1:     'Convert Isotherms Table.
      Call Convert_Isotherms_Table
  End Select
  Call Launch_Notepad_Log_TXT
End Sub


Private Sub Form_Load()
Dim FDIR_ROOT As String
  '
  ' SET UP DEFAULT DIRECTORIES AND DATABASE NAMES.
  '
  FDIR_ROOT = "X:\etdot10\code\ads\comm\990617_conversion_of_isotherm_database\vb6\data"
  ''''''''txtName(0).Text = FDIR_ROOT & "\db4_working\vb6\log.txt"
  ''''txtName(0).Text = FDIR_ROOT & "\code\vb6\log.txt"
  txtName(0).Text = FDIR_ROOT
  txtName(1).Text = FDIR_ROOT & "\in\isotherm.mdb"
  txtName(2).Text = FDIR_ROOT & "\out\isotherm.mdb"
  txtName(3).Text = FDIR_ROOT & "\in\chemical_name_change_notes.txt"
  '
  ' MISC OTHER FORM INITS.
  '
''''  pbItem.Value = 0
''''  pbItem.Value = 1000
''''  pbItem.Value = 500
''''  pbItem.Value = 0
''''  txtItemName.Text = ""
''''  ''''MsgBox "`" & CAS_Convert("00001238-123-124214   0") & "`"
''''  Call Populate_cboUnifacDatabase
''''  Call Populate_cboMosdapSearches
''''  '
''''  ' LOAD HINE-MOOKERJE FACTORS.
''''  '
''''  If (False = Load_Factors_HineMookerje()) Then
''''    End
''''  End If
End Sub

