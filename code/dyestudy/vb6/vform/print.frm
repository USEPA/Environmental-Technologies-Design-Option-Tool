VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmPrint 
   Caption         =   "frmPrint"
   ClientHeight    =   6315
   ClientLeft      =   315
   ClientTop       =   2430
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6315
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   1005
      Left            =   2100
      TabIndex        =   1
      Top             =   4470
      Visible         =   0   'False
      Width           =   2805
      _Version        =   65536
      _ExtentX        =   4948
      _ExtentY        =   1773
      _StockProps     =   14
      Caption         =   "Invisible"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   330
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VCIF1Lib.F1Book f1book 
      Height          =   4275
      Left            =   60
      OleObjectBlob   =   "print.frx":0000
      TabIndex        =   0
      Top             =   90
      Width           =   6345
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As ..."
         Index           =   40
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   49
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Page Setup ..."
         Index           =   50
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Printer Setup ..."
         Index           =   55
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print (Current Sheet Only) ..."
         Index           =   60
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   198
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Close"
         Index           =   199
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Copy"
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmPrint_loading_now As Boolean






Const frmPrint_declarations_end = 0


Function f1file_saveas(fn_force As String) As Integer
Dim fn_saveas As String
  If (fn_force <> "") Then
    fn_saveas = fn_force
  Else
    'INPUT NEW FILENAME.
    On Error GoTo err_filesaveas
    frmPrint.CommonDialog1.Filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
    frmPrint.CommonDialog1.FilterIndex = 2
    frmPrint.CommonDialog1.ShowSave
    fn_saveas = Trim$(frmPrint.CommonDialog1.FileName)
    If (fn_saveas = "") Then
      GoTo exit_save_did_not_go_okay
    End If
  End If
  'SAVE THIS FILE.
  f1book.Write frmPrint.CommonDialog1.FileName, F1FileExcel5
exit_save_went_okay:
  'SAVE WENT OKAY.
  f1file_saveas = True
  Exit Function
exit_save_did_not_go_okay:
  'SAVE DID NOT GO OKAY.
  f1file_saveas = False
  Exit Function
err_filesaveas:
  f1file_saveas = False
  Resume exit_save_did_not_go_okay
End Function


Private Sub Form_Activate()
  If (frmPrint_loading_now) Then
    frmPrint_loading_now = False
    Call Form_Resize
  End If
End Sub
Private Sub Form_Load()
Dim msg As String
  msg = "Pre-Print of "
  If (Trim$(Current_Filename) = "") Then
    msg = msg & "(untitled)"
  Else
    msg = msg & Current_Filename
  End If
  Me.Caption = msg
  Me.Width = 9000
  Me.Height = 7000
  Call CenterOnForm(Me, frmMain)
  Me.WindowState = 2  'maximized
  frmPrint_loading_now = True
  Call PrintTo_f1book(frmPrint.f1book, NowProj)
End Sub
Private Sub Form_Resize()
Dim XXX As Long
  If (frmPrint_loading_now) Then Exit Sub
  If (Me.WindowState = 1) Then
    'CANNOT RESIZE WHEN MINIMIZED; EXIT OUTTA HERE.
    Exit Sub
  End If
  If (Me.Width < 7000) Then Me.Width = 7000
  If (Me.Height < 3000) Then Me.Height = 3000
  XXX = Me.Width - (Me.Width - Me.ScaleWidth) - 2 * f1book.Left
  If (XXX < 1000) Then
    XXX = 1000
  End If
  f1book.Width = XXX
  XXX = Me.Height - (Me.Height - Me.ScaleHeight) - 2 * f1book.Top
  If (XXX > 1000) Then
    f1book.Height = XXX
  End If
End Sub


Private Sub mnuEditItem_Click(Index As Integer)
  Select Case Index
    Case 10:    'copy
      f1book.EditCopy
  End Select
End Sub
Private Sub mnuFileItem_Click(Index As Integer)
On Error GoTo err_mnuFileItem_Click
  Select Case Index
    Case 40:    'save as ...
      Call f1file_saveas("")
    Case 50:    'page setup ...
      f1book.FilePageSetupDlg
    Case 55:    'print setup ...
      f1book.FilePrintSetupDlg
    Case 85:    'print ...
      f1book.FilePrint True
    Case 199:   'close
      Unload Me
  End Select
exit_err_mnuFileItem_Click:
  Exit Sub
err_mnuFileItem_Click:
  'PROBABLY A CANCEL FROM f1book.FilePrint.
  Resume exit_err_mnuFileItem_Click
End Sub


