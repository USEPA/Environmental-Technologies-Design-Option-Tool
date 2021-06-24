VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmExcelCurves 
   Caption         =   "frmExcelCurves"
   ClientHeight    =   6315
   ClientLeft      =   1740
   ClientTop       =   3285
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
      OleObjectBlob   =   "ExcelCurves.frx":0000
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
         Caption         =   "&Copy Selection to Clipboard"
         Index           =   10
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Copy &Entire Table to Clipboard"
         Index           =   20
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmExcelCurves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmExcelCurves_loading_now As Boolean






Const frmExcelCurves_declarations_end = 0


Sub Do_The_Print_To_F1Book()


Dim f1 As Control
Set f1 = Me.f1book
'Dim sserror As Integer
Dim i As Integer
Dim j As Integer
Dim last_row As Integer
Dim last_col As Integer
  If (CPHSDM_Excel) Then
    '
    '    ----    CPHSDM RESULTS.    ----
    '
    f1.MinCol = 1
    f1.MinRow = 1
    f1.EntryRC(1, 1) = "CPHSDM Results -- Filename = " & Chr$(34) & Trim$(Filename) & Chr$(34)
    f1.EntryRC(3, 1) = "Time"
    f1.EntryRC(3, 2) = "BVT"
    f1.EntryRC(3, 3) = "Usage Rate"
    f1.EntryRC(3, 4) = CPM_Results.Component.Name
    f1.EntryRC(4, 1) = "Minutes"
    f1.EntryRC(4, 2) = "-"
    f1.EntryRC(4, 3) = "m³/kg GAC"
    f1.EntryRC(4, 4) = "-"
    For i = 1 To 100
      f1.EntryRC(4 + i, 1) = Trim$(Str$(CPM_Results.T(i)))
      f1.EntryRC(4 + i, 2) = Trim$(Str$(CPM_Results.T(i) * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2#) ^ 2))
      f1.EntryRC(4 + i, 3) = Trim$(Str$(CPM_Results.T(i) * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight))
      f1.EntryRC(4 + i, 4) = Trim$(Str$(CPM_Results.C_Over_C0(i)))
      '
      ' FORMAT NUMBERS IN THIS ROW AS "GENERAL" FORMAT.
      '
      f1.SelStartRow = 4 + i
      f1.SelEndRow = 4 + i
      f1.SelStartCol = 1
      f1.SelEndCol = 4
      f1.FormatGeneral
    Next i
    last_row = 4 + 100
    last_col = 4
    f1.MaxCol = last_col
    f1.MaxRow = last_row
  Else
    '
    '    ----    PSDM RESULTS.    ----
    '
Dim dblConvertedCP As Double
Dim dbl_CPConversionFactor() As Double
Dim OUT_strYAxisTitle As String
    f1.MinCol = 1
    f1.MinRow = 1
    f1.EntryRC(1, 1) = "PSDM Results -- Filename = " & Chr$(34) & Trim$(Filename) & Chr$(34)
    '
    ' GET THE UNIT CONVERSION FACTOR FOR EACH CHEMICAL.
    '
    ReDim dbl_CPConversionFactor(1 To Results.NComponent)
    For i = 1 To Results.NComponent
      dbl_CPConversionFactor(i) = _
          CBOYAXISTYPE_GetUnitConversion( _
          CInt(frmModelPSDMResults.cboYAxisType.ItemData( _
              frmModelPSDMResults.cboYAxisType.ListIndex)), _
          Results.is_psdm_in_room_model, _
          Results.AnyCrCloseToZero, _
          i, _
          Results.Bed.Phase, _
          OUT_strYAxisTitle)
    Next i
    '
    ' SET UP THE COLUMN HEADINGS.
    '
    f1.EntryRC(3, 1) = "Time"
    f1.EntryRC(3, 2) = "BVT"
    f1.EntryRC(3, 3) = "Usage Rate"
    For i = 1 To Results.NComponent
      f1.EntryRC(3, 3 + i) = Trim$(Results.Component(i).Name)
      f1.EntryRC(4, 3 + i) = OUT_strYAxisTitle
      ''''f1.EntryRC(4, 3 + i) = "-"
    Next i
    f1.EntryRC(4, 1) = "Minutes"
    f1.EntryRC(4, 2) = "-"
    f1.EntryRC(4, 3) = "m³/kg GAC"
    '
    ' OUTPUT THE VALUES.
    '
    For i = 1 To Results.npoints
      f1.EntryRC(4 + i, 1) = Trim$(Str$(Results.T(i)))
      f1.EntryRC(4 + i, 2) = Trim$(Str$(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2))
      f1.EntryRC(4 + i, 3) = Trim$(Str$(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.Weight))
      For j = 1 To Results.NComponent
        dblConvertedCP = Results.CP(j, i) * dbl_CPConversionFactor(j)
        f1.EntryRC(4 + i, 3 + j) = Trim$(Str$(dblConvertedCP))
        ''''f1.EntryRC(4 + i, 3 + j) = Trim$(Str$(Results.CP(j, i)))
      Next j
      '
      ' FORMAT NUMBERS IN THIS ROW AS "GENERAL" FORMAT.
      '
      f1.SelStartRow = 4 + i
      f1.SelEndRow = 4 + i
      f1.SelStartCol = 1
      f1.SelEndCol = 3 + j
      f1.FormatGeneral
      ''''f1.NumberFormat = "0.0"
    Next i
    last_row = 4 + Results.npoints
    last_col = 3 + Results.NComponent
    f1.MaxCol = last_col
    f1.MaxRow = last_row
  End If
  'temp = ""
  'For i = 1 To Results.NPoints
  ' temp = Format$(Results.T(i), "0.00")
  ' temp = temp & "       " & Format$(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.Length / Pi / (Results.Bed.Diameter / 2) ^ 2, "0.00")
  ' temp = temp & "       " & Format$(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.Weight, "0.00")
  ' For j = 1 To Results.NComponent
  '   temp = temp & "          " & Format$(Results.CP(j, i), "0.000")
  ' Next j
  ' Print #f, temp
  ' temp = ""
  'Next i
  'Close f
End Sub


Function f1file_saveas(fn_force As String) As Integer
Dim fn_saveas As String
  If (fn_force <> "") Then
    fn_saveas = fn_force
  Else
    'INPUT NEW FILENAME.
    On Error GoTo err_filesaveas
    frmExcelCurves.CommonDialog1.Filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
    frmExcelCurves.CommonDialog1.FilterIndex = 2
    frmExcelCurves.CommonDialog1.CancelError = True
    frmExcelCurves.CommonDialog1.flags = _
          cdlOFNOverwritePrompt + _
          cdlOFNPathMustExist
    frmExcelCurves.CommonDialog1.ShowSave
    fn_saveas = Trim$(frmExcelCurves.CommonDialog1.Filename)
    If (fn_saveas = "") Then
      GoTo exit_save_did_not_go_okay
    End If
  End If
  '
  ' SAVE THIS FILE.
  '
  Me.MousePointer = 11
  f1book.Write _
      frmExcelCurves.CommonDialog1.Filename, _
      F1FileExcel4
  '
  ' TODO: DETERMINE WHAT VERSIONS OF EXCEL THIS IS COMPATIBLE WITH!
  '
  Me.MousePointer = 0
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
  If (frmExcelCurves_loading_now) Then
    frmExcelCurves_loading_now = False
    Call Form_Resize
  End If
End Sub
Private Sub Form_Load()
Dim msg As String
  If (CPHSDM_Excel) Then
    msg = "Pre-Print of CPHSDM Results"
  Else
    msg = "Pre-Print of PSDM Results"
  End If
  'If (Trim$(NowProj.Filename) = "") Then
  '  msg = msg & "(untitled)"
  'Else
  '  msg = msg & NowProj.Filename
  'End If
  Me.Caption = msg
  Me.Width = 9000
  Me.Height = 7000
  Call CenterOnForm(Me, frmMain)
  Me.WindowState = 2  'maximized
  frmExcelCurves_loading_now = True
  Call Do_The_Print_To_F1Book
End Sub
Private Sub Form_Resize()
Dim XXX As Long
  If (frmExcelCurves_loading_now) Then Exit Sub
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
    Case 20:    'EDIT--COPY ENTIRE TABLE.
      f1book.SelStartCol = 1
      f1book.SelEndCol = f1book.MaxCol
      f1book.SelStartRow = 1
      f1book.SelEndRow = f1book.MaxRow
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
    Case 60:    'print ...
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


