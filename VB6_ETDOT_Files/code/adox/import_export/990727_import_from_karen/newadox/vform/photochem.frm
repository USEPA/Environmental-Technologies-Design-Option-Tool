VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmPhotoChem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing Photochemical Parameters"
   ClientHeight    =   6600
   ClientLeft      =   2265
   ClientTop       =   2265
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6600
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboUnits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cboUnits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   5610
      TabIndex        =   1
      Top             =   150
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   6720
      TabIndex        =   0
      Top             =   150
      Width           =   1065
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4725
      Left            =   60
      TabIndex        =   2
      Top             =   1740
      Width           =   3915
      _Version        =   65536
      _ExtentX        =   6906
      _ExtentY        =   8334
      _StockProps     =   14
      Caption         =   "Wavelengths:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdWavelen 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   1320
         Width           =   800
      End
      Begin VB.CommandButton cmdWavelen 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   900
         TabIndex        =   5
         Top             =   1320
         Width           =   800
      End
      Begin VB.CommandButton cmdWavelen 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1710
         TabIndex        =   4
         Top             =   1320
         Width           =   800
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   795
         Left            =   90
         TabIndex        =   20
         Top             =   300
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   1402
         _StockProps     =   14
         Caption         =   "Light Specification Method:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboLightSpecMethod 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   330
            Width           =   3525
         End
      End
      Begin VCIF1Lib.F1Book f1book_wavelen 
         Height          =   2865
         Left            =   90
         OleObjectBlob   =   "photochem.frx":0000
         TabIndex        =   3
         Top             =   1710
         Width           =   3735
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   4725
      Left            =   3960
      TabIndex        =   7
      Top             =   1740
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
      _ExtentY        =   8334
      _StockProps     =   14
      Caption         =   "Values for Each Compound:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboTarget 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   330
         Width           =   3975
      End
      Begin VCIF1Lib.F1Book f1book_vals 
         Height          =   2835
         Left            =   150
         OleObjectBlob   =   "photochem.frx":0584
         TabIndex        =   10
         Top             =   1710
         Width           =   5205
      End
      Begin VB.Label lbldesc 
         Caption         =   "Compound:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   100
         Left            =   120
         TabIndex        =   9
         Top             =   390
         Width           =   1215
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1635
      Left            =   60
      TabIndex        =   11
      Top             =   30
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
      _ExtentY        =   2884
      _StockProps     =   14
      Caption         =   "Lamp Parameters"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1890
         TabIndex        =   17
         Text            =   "txtData(2)"
         Top             =   1170
         Width           =   1095
      End
      Begin VB.TextBox txtDataStr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1890
         TabIndex        =   15
         Text            =   "txtDataStr(0)"
         Top             =   750
         Width           =   3375
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1890
         TabIndex        =   12
         Text            =   "txtData(0)"
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label lbldesc 
         Caption         =   "UV Path Length"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   19
         Top             =   1215
         Width           =   1755
      End
      Begin VB.Label lbldesc2 
         Caption         =   "cm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3090
         TabIndex        =   18
         Top             =   1215
         Width           =   825
      End
      Begin VB.Label lbldesc 
         Caption         =   "Lamp Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   795
         Width           =   1755
      End
      Begin VB.Label lbldesc 
         Caption         =   "Lamp Power"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   375
         Width           =   1755
      End
      Begin VB.Label lbldesc2 
         Caption         =   "watts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3090
         TabIndex        =   13
         Top             =   375
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmPhotoChem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Integer
Dim TempProj As Project_Type





Const frmPhotoChem_declarations_end = 0

Sub Populate_frmPhotoChem_Units()
  Call unitsys_register(frmPhotoChem, lbldesc(0), txtData(0), cboUnits(0), "power", _
      "w", "w", "", "", 500#, True)
  Call unitsys_register(frmPhotoChem, lbldesc(2), txtData(2), cboUnits(2), "length", _
      "cm", "cm", "", "", 6.8, True)
End Sub
  
'RETURNS:
'  TRUE = USER HIT OK
'  FALSE = USER HIT CANCEL
Public Function frmPhotoChem_DoEdit() As Integer
Dim is_aborted As Integer
Dim name_new As String
  
  'IMPORT THIS PROJECT FROM MEMORY TO THE FORM.
  TempProj = NowProj
  
  'SHOW THE FORM.
  frmPhotoChem.Show 1
  
  'UPDATE MEMORY.
  If (Not USER_HIT_CANCEL) Then
    NowProj = TempProj
  End If
  
  'RETURN TO MAIN WINDOW.
  frmPhotoChem_DoEdit = Not USER_HIT_CANCEL

End Function


Private Sub wavelengths_sort()
Dim done As Integer
Dim i As Integer
Dim j As Integer
Dim do_swap As Integer
Dim temp As Double
Dim it_now As Integer
Dim it_max As Integer
Dim tempw As Wavelength_Type

  'PERFORM BUBBLESORT ON WAVELENGTHS.
  'TO DO SO, THE FOLLOWING MUST BE SORTED BY SWITCHING ELEMENTS "i":
  '    WAVELENGTHS(i)
  '    EXTCOEF(j,i)
  '    QUATYD(j,i)
  it_now = 0
  it_max = 1000
  Do
    i = 1
    done = False
    do_swap = False
    Do
      If (i >= TempProj.Wavelength_Count) Then
        done = True
        Exit Do
      End If
      If (TempProj.Wavelengths(i + 1).lwave < TempProj.Wavelengths(i).lwave) Then
        done = False
        do_swap = True
        Exit Do
      End If
      i = i + 1
    Loop Until (1 <> 1)
    If (done) Then Exit Do
    If (do_swap) Then
      'SWAP ELEMENT i WITH ELEMENT i+1.
      '    SWAP WAVELENGTHS(i).
      tempw = TempProj.Wavelengths(i)
      TempProj.Wavelengths(i) = TempProj.Wavelengths(i + 1)
      TempProj.Wavelengths(i + 1) = tempw

      '    SWAP EXTCOEF(j,i).
      For j = 1 To TempProj.TargetCompounds_Count
        temp = TempProj.extcoef(j, i)
        TempProj.extcoef(j, i) = TempProj.extcoef(j, i + 1)
        TempProj.extcoef(j, i + 1) = temp
      Next j
      
      '    SWAP QUATYD(j,i).
      For j = 1 To TempProj.TargetCompounds_Count
        temp = TempProj.quatyd(j, i)
        TempProj.quatyd(j, i) = TempProj.quatyd(j, i + 1)
        TempProj.quatyd(j, i + 1) = temp
      Next j
    
      '    SWAP EXTCOEF_H2O2(i).
      temp = TempProj.extcoef_h2o2(i)
      TempProj.extcoef_h2o2(i) = TempProj.extcoef_h2o2(i + 1)
      TempProj.extcoef_h2o2(i + 1) = temp
      
      '    SWAP QUATYD_H2O2(i).
      temp = TempProj.quatyd_h2o2(i)
      TempProj.quatyd_h2o2(i) = TempProj.quatyd_h2o2(i + 1)
      TempProj.quatyd_h2o2(i + 1) = temp
    End If
    
    it_now = it_now + 1
    If (it_now > it_max) Then
      'IT'S DOUBTFUL THIS MAXIMUM ITERATION CHECK IS REQUIRED,
      'BUT IT'S BETTER TO BE PARANOID THAN TO GENERATE AN
      'INFINITE LOOP.
      Exit Do
    End If
  Loop Until (1 <> 1)
  
  

End Sub


Private Sub cboLightSpecMethod_Click()
  If (cboLightSpecMethod.ListIndex <> val(cboLightSpecMethod.Tag)) Then
    Select Case cboLightSpecMethod.ItemData(cboLightSpecMethod.ListIndex)
      Case IDUVI_EINSTEINS_L_S: TempProj.iduvi = IDUVI_EINSTEINS_L_S
      Case IDUVI_WATTS: TempProj.iduvi = IDUVI_WATTS
      Case IDUVI_EFFICIENCY: TempProj.iduvi = IDUVI_EFFICIENCY
    End Select
    Call DirtyFlag_Throw(TempProj)
    Call refresh_frmPhotoChem(TempProj)
  End If
End Sub

Private Sub cboTarget_Click()
  If (cboTarget.ListIndex <> cboTarget.Tag) Then
    Call refresh_frmPhotoChem(TempProj)
  End If
End Sub


Private Sub cboUnits_Click(Index As Integer)
Dim ctl As Control
Set ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(ctl)
End Sub


Private Sub cmdCancel_Click()
  USER_HIT_CANCEL = True
  Unload Me
End Sub


Private Sub cmdOK_Click()
  If (cboTarget.Enabled = False) Then
    Call Show_Error("Before you can save the data on this form, " & _
        "you must first complete your data entry on the " & _
        "extinction coefficient / quantum yield grid.  " & _
        "To do so, click on the grid, and then press the Enter key.")
    Exit Sub
  End If
  USER_HIT_CANCEL = False
  Unload Me
End Sub


Private Sub cmdWavelen_Click(Index As Integer)
Dim name_original As String
Dim name_new As String
Dim is_aborted As Integer
Dim idx As Integer
Dim retval As Integer
Dim msg As String
Dim idx_del As Integer
Dim idx_max As Integer
Dim i As Integer
Dim j As Integer
Dim temp() As Double
Dim wavelen_new As Double

  Select Case Index
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Case 0:   'add
      Do While (1 = 1)
add_try_again:
        name_new = frmNewName.frmNewName_GetName( _
            "Enter New Wavelength", _
            "The unit for wavelength is nanometers (nm).", _
            name_new, _
            is_aborted)
        If (is_aborted) Then
          'USER HIT CANCEL.
          Exit Sub
        End If
        If (Not IsValidNumber0(name_new, vbDouble)) Then
          Call Show_Error("Invalid number.  Please re-enter or cancel.")
          GoTo add_try_again
        End If
        wavelen_new = CDbl(name_new)
        If (wavelen_new <= 0#) Then
          Call Show_Error("Invalid number.  Only positive numbers " & _
              "may be specified for this value.  Please re-enter or cancel.")
          GoTo add_try_again
        End If
        For i = 1 To TempProj.Wavelength_Count
          If (wavelen_new = TempProj.Wavelengths(i).lwave) Then
            Call Show_Error("That wavelength already exists.  Choose another wavelength or cancel.")
            GoTo add_try_again
          End If
        Next i
        Exit Do
      Loop

      'ADD NEW WAVELENGTH STRUCTURE.
      idx_del = idx
      idx_max = TempProj.Wavelength_Count
      ReDim Preserve TempProj.Wavelengths(1 To idx_max + 1)
      TempProj.Wavelengths(idx_max + 1).lwave = wavelen_new
      TempProj.Wavelengths(idx_max + 1).uvi = UVI_DEFAULT_VALUE
      
      'ADD NEW SET OF EXTINCTION COEFFICIENTS (ONE PER WAVELENGTH).
      'NOTE: THIS STUPID TEMPORARY ARRAY IS NECESSARY TO GET AROUND THE
      'VISUAL BASIC STIPULATION THAT YOU CAN ONLY "REDIM PRESERVE" AN ARRAY
      'IF YOU ARE CHANGING THE *LAST* ARRAY INDEX.
      ReDim temp(1 To TempProj.TargetCompounds_Count, 1 To idx_max + 1)
      For i = 1 To TempProj.TargetCompounds_Count
        For j = 1 To idx_max
          temp(i, j) = TempProj.extcoef(i, j)
        Next j
      Next i
      For i = 1 To TempProj.TargetCompounds_Count
        temp(i, idx_max + 1) = EXTCOEF_DEFAULT_VALUE
      Next i
      ReDim TempProj.extcoef(1 To TempProj.TargetCompounds_Count, 1 To idx_max + 1)
      For i = 1 To TempProj.TargetCompounds_Count
        For j = 1 To idx_max + 1
          TempProj.extcoef(i, j) = temp(i, j)
        Next j
      Next i
      
      'ADD NEW SET OF QUANTUM YIELDS (ONE PER WAVELENGTH).
      'NOTE: THIS STUPID TEMPORARY ARRAY IS NECESSARY TO GET AROUND THE
      'VISUAL BASIC STIPULATION THAT YOU CAN ONLY "REDIM PRESERVE" AN ARRAY
      'IF YOU ARE CHANGING THE *LAST* ARRAY INDEX.
      ReDim temp(1 To TempProj.TargetCompounds_Count, 1 To idx_max + 1)
      For i = 1 To TempProj.TargetCompounds_Count
        For j = 1 To idx_max
          temp(i, j) = TempProj.quatyd(i, j)
        Next j
      Next i
      For i = 1 To TempProj.TargetCompounds_Count
        temp(i, idx_max + 1) = QUATYD_DEFAULT_VALUE
      Next i
      ReDim TempProj.quatyd(1 To TempProj.TargetCompounds_Count, 1 To idx_max + 1)
      For i = 1 To TempProj.TargetCompounds_Count
        For j = 1 To idx_max + 1
          TempProj.quatyd(i, j) = temp(i, j)
        Next j
      Next i
      
      'ADD NEW EXTINCTION COEFFICIENT FOR H2O2 (ONE PER WAVELENGTH).
      ReDim Preserve TempProj.extcoef_h2o2(1 To idx_max + 1)
      TempProj.extcoef_h2o2(idx_max + 1) = EXTCOEF_H2O2_DEFAULT_VALUE
      
      'ADD NEW QUANTUM YIELD FOR H2O2 (ONE PER WAVELENGTH).
      ReDim Preserve TempProj.quatyd_h2o2(1 To idx_max + 1)
      TempProj.quatyd_h2o2(idx_max + 1) = QUATYD_H2O2_DEFAULT_VALUE
      
      'UPDATE THE NUMBER OF TARGET COMPOUNDS.
      TempProj.Wavelength_Count = idx_max + 1
    
      'SORT THE WHOLE MESS.
      Call wavelengths_sort
    
      'REFRESH PHOTOCHEMICAL WINDOW, ESPECIALLY THE GRIDS.
      Call refresh_frmPhotoChem(TempProj)
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Case 1:   'delete
      idx = f1book_wavelen.Row
      If (idx < 1) Then Exit Sub
      If (TempProj.Wavelength_Count = 1) Then
        Call Show_Error("You cannot delete the last wavelength; " & _
            "there must always be at least one wavelength defined.")
        Exit Sub
      End If
      
      msg = "If you delete wavelength [" & _
            Trim$(Str$(TempProj.Wavelengths(idx).lwave)) & "] you will not be able to " & _
            "undelete it.  Are you sure you want to delete it?"
      retval = MsgBox(msg, vbCritical + vbYesNo, App.title)
      If (retval = vbNo) Then Exit Sub
      
      'DELETE THIS WAVELENGTH STRUCTURE.
      idx_del = idx
      idx_max = TempProj.Wavelength_Count
      For i = idx_del To idx_max - 1
        TempProj.Wavelengths(i) = TempProj.Wavelengths(i + 1)
      Next i
      ReDim Preserve TempProj.Wavelengths(1 To idx_max - 1)
      
      'DELETE THIS SET OF EXTINCTION COEFFICIENTS (ONE PER WAVELENGTH).
      'NOTE: THIS STUPID TEMPORARY ARRAY IS NECESSARY TO GET AROUND THE
      'VISUAL BASIC STIPULATION THAT YOU CAN ONLY "REDIM PRESERVE" AN ARRAY
      'IF YOU ARE CHANGING THE *LAST* ARRAY INDEX.
      ReDim temp(1 To TempProj.TargetCompounds_Count, 1 To idx_max - 1)
      For i = 1 To TempProj.TargetCompounds_Count
        For j = 1 To idx_del - 1
          temp(i, j) = TempProj.extcoef(i, j)
        Next j
      Next i
      For i = 1 To TempProj.TargetCompounds_Count
        For j = idx_del To idx_max - 1
          temp(i, j) = TempProj.extcoef(i, j + 1)
        Next j
      Next i
      ReDim TempProj.extcoef(1 To TempProj.TargetCompounds_Count, 1 To idx_max - 1)
      For i = 1 To TempProj.TargetCompounds_Count
        For j = 1 To idx_max - 1
          TempProj.extcoef(i, j) = temp(i, j)
        Next j
      Next i
      
      'DELETE THIS SET OF QUANTUM YIELDS (ONE PER WAVELENGTH).
      'NOTE: THIS STUPID TEMPORARY ARRAY IS NECESSARY TO GET AROUND THE
      'VISUAL BASIC STIPULATION THAT YOU CAN ONLY "REDIM PRESERVE" AN ARRAY
      'IF YOU ARE CHANGING THE *LAST* ARRAY INDEX.
      ReDim temp(1 To TempProj.TargetCompounds_Count, 1 To idx_max - 1)
      For i = 1 To TempProj.TargetCompounds_Count
        For j = 1 To idx_del - 1
          temp(i, j) = TempProj.quatyd(i, j)
        Next j
      Next i
      For i = 1 To TempProj.TargetCompounds_Count
        For j = idx_del To idx_max - 1
          temp(i, j) = TempProj.quatyd(i, j + 1)
        Next j
      Next i
      ReDim TempProj.quatyd(1 To TempProj.TargetCompounds_Count, 1 To idx_max - 1)
      For i = 1 To TempProj.TargetCompounds_Count
        For j = 1 To idx_max - 1
          TempProj.quatyd(i, j) = temp(i, j)
        Next j
      Next i

      'DELETE THIS EXTINCTION COEFFICIENT FOR H2O2 (ONE PER WAVELENGTH).
      For i = idx_del To idx_max - 1
        TempProj.extcoef_h2o2(i) = TempProj.extcoef_h2o2(i + 1)
      Next i
      ReDim Preserve TempProj.extcoef_h2o2(1 To idx_max - 1)
      
      'DELETE THIS QUANTUM YIELD FOR H2O2 (ONE PER WAVELENGTH).
      For i = idx_del To idx_max - 1
        TempProj.quatyd_h2o2(i) = TempProj.quatyd_h2o2(i + 1)
      Next i
      ReDim Preserve TempProj.quatyd_h2o2(1 To idx_max - 1)
      
      'UPDATE THE NUMBER OF TARGET COMPOUNDS.
      TempProj.Wavelength_Count = idx_max - 1
    
      'REFRESH PHOTOCHEMICAL WINDOW, ESPECIALLY THE GRIDS.
      Call refresh_frmPhotoChem(TempProj)
      
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Case 2:   'edit
      idx = f1book_wavelen.Row
      If (idx < 1) Then Exit Sub
            
      name_new = Trim$(Str$(TempProj.Wavelengths(idx).lwave))
      
      Do While (1 = 1)
edit_try_again:
        name_new = frmNewName.frmNewName_GetName( _
            "Enter New Wavelength to Replace This Wavelength With", _
            "The unit for wavelength is nanometers (nm).", _
            name_new, _
            is_aborted)
        If (is_aborted) Then
          'USER HIT CANCEL.
          Exit Sub
        End If
        If (Not IsValidNumber0(name_new, vbDouble)) Then
          Call Show_Error("Invalid number.  Please re-enter or cancel.")
          GoTo edit_try_again
        End If
        wavelen_new = CDbl(name_new)
        If (wavelen_new <= 0#) Then
          Call Show_Error("Invalid number.  Only positive numbers " & _
              "may be specified for this value.  Please re-enter or cancel.")
          GoTo edit_try_again
        End If
        For i = 1 To TempProj.Wavelength_Count
          If (wavelen_new = TempProj.Wavelengths(i).lwave) Then
            Call Show_Error("That wavelength already exists.  Choose another wavelength or cancel.")
            GoTo edit_try_again
          End If
        Next i
        Exit Do
      Loop
      
      'REPLACE THIS WAVELENGTH.
      TempProj.Wavelengths(idx).lwave = wavelen_new
      
      'SORT THE WHOLE MESS.
      Call wavelengths_sort
    
      'REFRESH PHOTOCHEMICAL WINDOW, ESPECIALLY THE GRIDS.
      Call refresh_frmPhotoChem(TempProj)
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  End Select
End Sub


Private Sub f1book_vals_EndEdit(EditString As String, Cancel As Integer)
Dim idx As Integer
Dim newVal As Double
Dim idx_target As Integer

  idx = f1book_vals.Row
  If (idx < 1) Then Exit Sub
  idx_target = cboTarget.ListIndex + 1
  If (idx_target > TempProj.TargetCompounds_Count) Then
    'THIS IS THE H2O2 COMPOUND.
    idx_target = -1
  End If
  
  Select Case f1book_vals.Col
    Case 1:       'impossible; editting for this column is turned off.
    Case 2, 3:    'extinction coefficient or quantum yield
      If (Not IsValidNumber0(EditString, vbDouble)) Then
        Call Show_Error("Invalid number.  Please re-enter or cancel.")
        Cancel = True
        SendKeys "{escape}"
        Exit Sub
      End If
      newVal = CDbl(EditString)
      If (newVal <= 0#) Then
        Call Show_Error("Invalid number.  Only positive numbers " & _
            "may be specified for this value.  Please re-enter or cancel.")
        Cancel = True
        SendKeys "{escape}"
        Exit Sub
      End If
      If (idx_target = -1) Then
        'PROCESS H2O2 DATA.
        Select Case f1book_vals.Col
          Case 2:
            TempProj.extcoef_h2o2(idx) = newVal
          Case 3:
            TempProj.quatyd_h2o2(idx) = newVal
        End Select
      Else
        'PROCESS NON-H2O2 DATA.
        Select Case f1book_vals.Col
          Case 2:
            TempProj.extcoef(idx_target, idx) = newVal
          Case 3:
            TempProj.quatyd(idx_target, idx) = newVal
        End Select
      End If
      
      'REFRESH PHOTOCHEMICAL WINDOW, ESPECIALLY THE GRIDS.
      Call refresh_frmPhotoChem(TempProj)
      
      'RE-ENABLE THE TARGET SELECTIONBOX; THIS ENSURES THE USER DOES NOT LOSE
      'DATA ENTERED INTO THE GRID WITHOUT HITTING THE <Enter> KEY.
      cboTarget.Enabled = True
      
  End Select

End Sub


Private Sub f1book_vals_StartEdit(EditString As String, Cancel As Integer)
  'DISABLE THE TARGET SELECTIONBOX; THIS ENSURES THE USER DOES NOT LOSE
  'DATA ENTERED INTO THE GRID WITHOUT HITTING THE <Enter> KEY.
  cboTarget.Enabled = False
End Sub


Private Sub f1book_wavelen_EndEdit(EditString As String, Cancel As Integer)
Dim idx As Integer
Dim newVal As Double

  idx = f1book_wavelen.Row
  If (idx < 1) Then Exit Sub

  Select Case f1book_wavelen.Col
    Case 1:     'impossible; editting for this column is turned off.
    Case 2:     'uv light intensity
      If (Not IsValidNumber0(EditString, vbDouble)) Then
        Call Show_Error("Invalid number.  Please re-enter or cancel.")
        Cancel = True
        SendKeys "{escape}"
        Exit Sub
      End If
      newVal = CDbl(EditString)
      If (newVal <= 0#) Then
        Call Show_Error("Invalid number.  Only positive numbers " & _
            "may be specified for this value.  Please re-enter or cancel.")
        Cancel = True
        SendKeys "{escape}"
        Exit Sub
      End If
      TempProj.Wavelengths(idx).uvi = newVal
      
      'REFRESH PHOTOCHEMICAL WINDOW, ESPECIALLY THE GRIDS.
      Call refresh_frmPhotoChem(TempProj)
      
  End Select

End Sub


Sub populate_cboLightSpecMethod()
  cboLightSpecMethod.Clear
  cboLightSpecMethod.AddItem "Intensity, in Einsteins/L-s"
  cboLightSpecMethod.ItemData(cboLightSpecMethod.NewIndex) = IDUVI_EINSTEINS_L_S
  cboLightSpecMethod.AddItem "Intensity, in Watts"
  cboLightSpecMethod.ItemData(cboLightSpecMethod.NewIndex) = IDUVI_WATTS
  cboLightSpecMethod.AddItem "Efficiency, dim'less (range: 0-1)"
  cboLightSpecMethod.ItemData(cboLightSpecMethod.NewIndex) = IDUVI_EFFICIENCY
End Sub
Private Sub Form_Load()
  'MISC INITS.
  Call CenterOnForm(Me, frmMain)
  Call populate_cboLightSpecMethod
  Call Populate_frmPhotoChem_Units
  Call Populate_frmPhotoChem_Units
  Call refresh_frmPhotoChem(TempProj)
End Sub


Private Sub txtdata_GotFocus(Index As Integer)
Dim ctl As Control
Set ctl = txtData(Index)
  Call unitsys_control_txtx_gotfocus(ctl)
End Sub
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub

Private Sub txtdata_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim ctl As Control
Set ctl = txtData(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
Dim txtctl As Control
Set txtctl = txtData(Index)

  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  If (Index = 4) Then
    Val_Low = 1E-20 * 60#
    Val_High = 1E+20 * 60#
  Else
    Val_Low = 1E-20     '0.00000000000000000001
    Val_High = 1E+20    '100000000000000000000
  End If
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(ctl, NewValue)
 ' Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
          Select Case Index
              Case 0: TempProj.lamp_power = NewValue
              Case 2: TempProj.uvpathl = NewValue
          End Select
     
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set( _
            Project_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call refresh_frmMain
    End If
  End If
End Sub




Private Sub txtdataSTR_GotFocus(Index As Integer)
  Dim txtctl As Control
  Set txtctl = txtDataStr(Index)
  Call DisplayDataEntryError
  Call Global_GotFocus(txtctl)
End Sub
Private Sub txtdataSTR_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtdataSTR_LostFocus(Index As Integer)
  Dim txtctl As Control
  Set txtctl = txtDataStr(Index)
  Dim ok_to_save As Integer
  Dim refresh_type As Integer
  ok_to_save = False
  If (txtctl.Text <> txtctl.Tag) Then
    ok_to_save = True
  End If
  If (ok_to_save) Then
    'DATA LOOKS OKAY, LET'S GO AHEAD AND SAVE IT.
    refresh_type = 1
    Select Case Index
      Case 0: TempProj.lamp_name = Trim$(txtctl.Text)
    End Select
    
    Call AssignTextAndTag(txtctl, txtctl.Text)
    
    'THROW DIRTY FLAG, AND REFRESH EVERY WINDOWS.
    Call DirtyFlag_Throw(TempProj)
    
    Select Case refresh_type
      Case 1:   'JUST THE PHOTOCAT WINDOW.
        Call refresh_frmPhotoChem(TempProj)
    End Select
  End If
  Call Global_LostFocus(txtctl)
End Sub
Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub

Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub
