VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmexport 
   Caption         =   "Export"
   ClientHeight    =   6105
   ClientLeft      =   2010
   ClientTop       =   1185
   ClientWidth     =   6210
   ControlBox      =   0   'False
   Icon            =   "frmexpor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   6210
   Begin VB.OptionButton commaopt 
      Caption         =   "comma delimited"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   4800
      Width           =   1815
   End
   Begin VB.OptionButton spaceopt 
      Caption         =   "space delimited"
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame setframe 
      Caption         =   "User Templates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   2775
      Begin VB.ListBox designlst 
         Height          =   1230
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Chemical(s) to Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3015
      Begin VB.ListBox chemlst 
         Height          =   1230
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fieldsframe 
      Caption         =   "Select Fields For Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   5895
      Begin VB.CommandButton removecmd 
         Caption         =   "&Remove Design"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1275
      End
      Begin VB.TextBox nametbx 
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Text            =   "design1"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton savecmd 
         Caption         =   "&Save Design"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton leftcmd 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton rightcmd 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.ListBox selectlst 
         Height          =   1620
         Left            =   3240
         MultiSelect     =   1  'Simple
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.ListBox alllst 
         Height          =   1620
         Left            =   240
         MultiSelect     =   1  'Simple
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton appendcmd 
      Caption         =   "&Append to File"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton CMDCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton newcmd 
      Caption         =   "&New File"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DialogTitle     =   "Export Chemical Properties"
   End
End
Attribute VB_Name = "frmexport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub appendcmd_Click()
    Dim FNum As Integer
    Dim exportfile As String
    
    
        ' check that a chemical has been selected
    If (chemlst.SelCount = 0) Then
        MsgBox ("No chemicals selected to export")
        Exit Sub
    End If
       
    On Error GoTo cancel_error
    
    CommonDialog1.filename = "filename"
    CommonDialog1.InitDir = App.path
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNFileMustExist
    CommonDialog1.ShowSave
   
    FNum = FreeFile
        ' check that the file is good
    Open CommonDialog1.filename For Append As #FNum
    
    Close FNum
        ' the filename is good, continue with the export after closing the commondialog box
    exportfile = CommonDialog1.filename
    Call start_append_export(exportfile)
    
done_export:
    FRMExport.Hide
    FRMMain.Show
    Exit Sub
        ' this means the user hit the cancel button or there's an error of some kind
cancel_error:
    Screen.MousePointer = 1
    If Err = 32755 Then
        MsgBox ("File unchanged")
        FRMExport.Hide
        FRMMain.Refresh
        
        Exit Sub
    End If
    Screen.MousePointer = 1
    MsgBox ("Error opening " & CommonDialog1.filename)
    FRMExport.Hide
    FRMMain.Refresh
    
End Sub

Private Sub CMDCancel_Click()

    
    Unload Me
    FRMMain.Show
    
End Sub


Private Sub commaopt_Click()
    Dim answer As Integer
    'If existing = True And modified = False Then
     '   answer = MsgBox("Modify existing design?", 256 + 4 + 32)
     '   If answer = vbNo Then
      '      nametbx.Text = "unnamed"
       '     existing = False
            
       ' Else
       '     modified = True
        'End If
    'End If
    spaceopt.value = False
    commaopt.value = True
    
End Sub

Private Sub form_Unload(Cancel As Integer)
    
    Call write_def_for_export
    
End Sub

Private Sub removecmd_Click()

    Call remove_design
    
End Sub

Private Sub savecmd_Click()

    Dim newname As String
    Dim answer As Integer
    Dim i As Integer
    Dim identifier As Boolean
    Dim modified As Boolean
    newname = nametbx.Text
        ' warn the user if there's no identifying field (c#, cas# or chemical name)
    identifier = False
    For i = 0 To selectlst.ListCount - 1
        If selectlst.List(i) = "C#" Then
            identifier = True
            Exit For
        ElseIf selectlst.List(i) = "Cas #" Then
            identifier = True
            Exit For
        ElseIf selectlst.List(i) = "Chemical Name" Then
            identifier = True
            Exit For
        End If
    Next i
    If identifier = False Then
        answer = MsgBox("Your design has no unique identifying field, Continue?", 1 + 32 + 256)
        If answer = vbCancel Then
            Exit Sub
        End If
    End If
        ' now check if it's a modification of an existing design and if so, warn the user
    For i = 0 To designlst.ListCount - 1
        If newname = designlst.List(i) Then
            answer = MsgBox("export design " & Chr(34) & newname & Chr(34) & " already exists, replace?", vbYesNo, "Pearls")
            If answer = False Then
                nametbx.Text = "unnamed"
                Exit Sub
            Else
                
                modified = True
                
                Exit For
            End If
        End If
    Next i
    
    Call fill_design_array
    ex_design_existing = True
    ex_design_modified = False
End Sub

Private Sub designlst_DblClick()
    Call update_design_description
    
End Sub


Private Sub leftcmd_Click()

    Dim answer As Integer
    Dim i As Integer
    Dim J As Integer
    Dim num_to_remove As Integer
    Dim marked(15) As String
    Dim success As Boolean
    success = False
    If ex_design_existing = True And ex_design_modified = False Then
       answer = MsgBox("Modify existing design?", 256 + 4 + 32)
        If answer = vbNo Then
            nametbx.Text = "unnamed"
            ex_design_existing = False
            
        Else
            ex_design_modified = True
        End If
    End If
    num_to_remove = 0
    For i = 0 To selectlst.ListCount - 1
        If selectlst.Selected(i) = True Then
            alllst.AddItem selectlst.List(i)
            marked(num_to_remove) = selectlst.List(i)
            num_to_remove = num_to_remove + 1
            success = True
        
        End If
     Next i
     For i = 0 To num_to_remove - 1
        For J = 0 To selectlst.ListCount - 1
            If J = selectlst.ListCount Then
                Exit For
            End If
            If selectlst.List(J) = marked(i) Then
                selectlst.RemoveItem J
            End If
        Next J
    Next i
    If success = False Then
        MsgBox ("Select a field to remove")
    End If
    
    
End Sub

Private Sub nametbx_Change()

    Dim i As Integer
    For i = 0 To designlst.ListCount - 1
        If nametbx.Text = designlst.List(i) Then
            ex_design_existing = True
            Exit Sub
        End If
    Next i
    ex_design_existing = False
End Sub

Private Sub newcmd_Click()

    Dim exportfile As String
    Dim FNum As Integer
    If (FRMExport!chemlst.SelCount = 0) Then
        MsgBox ("No chemicals selected to export")
        Exit Sub
    End If
    On Error GoTo cancel_error
   
    CommonDialog1.filename = "filename"
    CommonDialog1.InitDir = App.path
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    CommonDialog1.ShowSave
   
    FNum = FreeFile
    
    Open CommonDialog1.filename For Output As #FNum
    Close FNum
        ' the filename is good, continue with the export after closing the commondialog box
    exportfile = CommonDialog1.filename
    Call start_new_export(exportfile)
done_export:
    FRMExport.Hide
    FRMMain.Show
    Exit Sub
cancel_error:
    Screen.MousePointer = 1
    If Err = 32755 Then
        MsgBox ("File unchanged")
        FRMExport.Hide
        FRMMain.Refresh
        Exit Sub
    End If
    Screen.MousePointer = 1
    MsgBox ("Error opening " & CommonDialog1.filename)
    FRMExport.Hide
    FRMMain.Refresh
End Sub






Private Sub rightcmd_Click()

   Dim answer As Integer
   Dim i As Integer
   Dim J As Integer
   Dim success As Boolean
   Dim num_to_remove As Integer
   Dim marked(15) As String
   success = False
    If ex_design_existing = True And ex_design_modified = False Then
        answer = MsgBox("Modify existing design?", 256 + 4 + 32)
        If answer = vbNo Then
            nametbx.Text = "unnamed"
            ex_design_existing = False
        Else
            ex_design_modified = True
        End If
        
    End If
    num_to_remove = 0
    For i = 0 To alllst.ListCount - 1
        If alllst.Selected(i) = True Then
            selectlst.AddItem alllst.List(i)
            marked(num_to_remove) = alllst.List(i)
            num_to_remove = num_to_remove + 1
            success = True
        End If
    Next i
    For i = 0 To num_to_remove - 1
        For J = 0 To alllst.ListCount - 1
            If J = alllst.ListCount Then
                Exit For
            End If
            If alllst.List(J) = marked(i) Then
                alllst.RemoveItem J
            End If
        Next J
    Next i
    If success = False Then
        MsgBox ("Select field(s) to add")
    End If
    
End Sub




Private Sub spaceopt_Click()
    Dim answer As Integer
    If ex_design_existing = True And ex_design_modified = False And ex_from_design = False Then
        answer = MsgBox("Modify existing design?", 256 + 4 + 32)
        If answer = vbNo Then
            nametbx.Text = "unnamed"
            ex_design_existing = False
            
        Else
            ex_design_modified = True
        End If
    End If
    If ex_from_design = True Then
        ex_from_design = False
    End If
    commaopt.value = False
    spaceopt.value = True
End Sub




