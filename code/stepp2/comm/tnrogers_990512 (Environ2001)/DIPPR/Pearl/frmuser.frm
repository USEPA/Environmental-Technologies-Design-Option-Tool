VERSION 5.00
Begin VB.Form frmuser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Login"
   ClientHeight    =   1920
   ClientLeft      =   1590
   ClientTop       =   3015
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1920
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton okcmd 
      Caption         =   "continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Top             =   1440
      Width           =   1635
   End
   Begin VB.TextBox usertbx 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label helplbl 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5715
   End
   Begin VB.Label userlbl 
      Caption         =   "user name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Form_Load()

    helplbl.caption = "Type a user name to identify your session (optional)"
    usertbx.Text = ""
    CenterForm Me

End Sub


Private Sub okcmd_Click()

    ' This subroutine handles checking for an existing .def file and,
    ' if not found, sets the deffile variable to the appropriate name for
    ' use by the path finding and export routines
    ' Should be as case insensitive as possible
    
    Dim username As String
    Dim tempfile As String
    Dim i As Integer
    Dim FNum As Integer
    Dim match As Boolean
    Dim tempname As String
    
    tempfile = "temp.def"
    ' if the user has typed something, check if the file already exists
    If Len(usertbx.Text) > 1 Then
        ' make it all lowercase for easier comparison
        ' this will cover case sensitivity in the input
        username = LCase(Trim(usertbx.Text))
        deffile = Dir("*.def")
        If deffile = Trim(usertbx.Text) & ".def" Or deffile = UCase(usertbx.Text) & ".def" Then
            match = True
            GoTo after_for
        End If
        While deffile <> "" And deffile <> " "
            deffile = Dir
            If deffile = Trim(usertbx.Text) & ".def" Or deffile = UCase(usertbx.Text) & ".def" Then
                match = True
                GoTo after_for
            End If
        Wend
        
    Else
    ' if the user didn't enter anything, use the default .def file
use_default:
        deffile = "pearls.def"
    End If
    
    ' now double check it's a valid file.  Check that the file
    ' is there and if not create it
    
after_for:
    match = deffile Like "*.def" Or LCase(deffile) Like "*.def"
    If match = True Then
        On Error GoTo else_default
        FNum = FreeFile
        Open App.path & "\" & deffile For Append As #FNum
        Close #FNum
        
        ' must be okay, continue on
        GoTo prompt_db
    ElseIf Len(username) > 4 & Len(username) < 15 Then
        ' create the file if it wasn't there using lowercase
        deffile = LCase(username) & ".def"
        ' check it's a legal filename
        On Error GoTo else_default
        FNum = FreeFile
        Open App.path & "\" & deffile For Append As #FNum
        Close #FNum
        ' must be okay, continue on
    Else
else_default:
        MsgBox ("Invalid username " & Chr(40) & "must be between 5 and 14 characters" & Chr(41))
        Exit Sub
    End If
prompt_db:
    
    On Error GoTo cancel_error
    ' get the info out of the def file
    Call parse_def_file
    Call load_frm_master_info
    Call CenterForm(frmmaster)
    frmmaster.Show 1
   
   ' of there's no problem so far, just return
    FRMUser.Hide
    DoEvents    'make sure it hides
    
    Exit Sub
cancel_error:
    MasterDBName = "master.mdb"
    FRMUser.Hide
End Sub

