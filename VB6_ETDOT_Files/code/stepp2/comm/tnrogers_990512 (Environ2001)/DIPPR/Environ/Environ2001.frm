VERSION 5.00
Begin VB.Form frmEnviron2001 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Environ 2001"
   ClientHeight    =   4425
   ClientLeft      =   255
   ClientTop       =   1980
   ClientWidth     =   7200
   Icon            =   "Environ2001.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Environ2001.frx":0442
   ScaleHeight     =   4425
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPearl 
      Caption         =   "PDE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   1875
   End
   Begin VB.CommandButton cmdDBMan 
      Caption         =   "DBManager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2550
      TabIndex        =   1
      Top             =   3480
      Width           =   1875
   End
   Begin VB.CommandButton cmdDCUT 
      Caption         =   "DCUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4920
      TabIndex        =   0
      Top             =   3480
      Width           =   1875
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Program"
         Begin VB.Menu mnuPearls 
            Caption         =   "&Pearls"
         End
         Begin VB.Menu mnuDBMan 
            Caption         =   "DB&Manager"
         End
         Begin VB.Menu mnuDcut 
            Caption         =   "D&Cut"
         End
      End
      Begin VB.Menu mnuspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmEnviron2001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdDBMan_Click()
Dim rc
rc = run_app("DBManager.exe")
End Sub

Private Sub cmdDCUT_Click()
Dim rc
rc = run_app("DCUT.exe")
End Sub

Private Sub cmdPearl_Click()
Dim rc
rc = run_app("Pearls.exe")
End Sub

Function run_app(AppName As String) As Boolean
Dim rc As Double
Dim FileName As String

FileName = Set_FileName(App.path, AppName)
If FileExists(FileName) Then
    rc = Shell(FileName, 1)
    'AppActivate rc
Else
    If MsgBox("This program file doesn't exist. Would you like me to locate it for you.", vbYesNo, "Error") = vbYes Then
        FileName = find_file(AppName)
        If FileName = "" Then
            MsgBox "The file '" & AppName & "' could not be found."
        Else
            rc = Shell(FileName & AppName, 1)
            'AppActivate rc
        End If
    End If
End If
End Function

'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function


'-----------------------------------------------------------
' Function: puts path and filename together and returns
'           the resulting string
'-----------------------------------------------------------
'
Function Set_FileName(ByVal strPath As String, ByVal strName As String) As String
Dim i As Integer
Dim mypos As String
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    Set_FileName = strPath & strName
End Function

Public Function find_file(name As String) As String

    ' this function locates a file and returns the path
    Dim path As String
    ' an array holding the paths during the search
    Dim paths(1000) As String
    Dim temppath As String
    Dim localpath As String
    Dim test_path As String
    Dim position As Integer
    Dim iteration_position As Integer
    Dim MAX_POSITION As Integer
    Dim drive_var As String
    On Error Resume Next
    
    path = ""
    ' first check if we already have the path
        
    MAX_POSITION = 999
    ' find the drive to start on based on the app.path
    drive_var = Left(App.path, 1) & ":"
    paths(0) = drive_var
    position = 1
    iteration_position = 0
    
    Screen.MousePointer = 11
    
    While position < 1000 And iteration_position < 1000
        ' search all files in this directory
        localpath = paths(iteration_position)
        temppath = Dir(paths(iteration_position) & "\" & name)
        While Trim(temppath) <> ""
            If UCase(Right(Trim(temppath), Len(name))) = UCase(Trim(name)) Then
                path = localpath & Left(temppath, Len(temppath) - Len(name))
                find_file = path & "\"
                Screen.MousePointer = 1
                Exit Function
            End If
            temppath = Dir
        Wend
        ' now get the directories
        temppath = Dir(paths(iteration_position) & "\", vbDirectory)
        
        While Trim(temppath) <> "" And position < MAX_POSITION
            If Trim(temppath) = ".." Or Trim(temppath) = "." Then
                GoTo next_directory_iteration
            End If
            If (GetAttr(localpath & "\" & temppath) And vbDirectory) = vbDirectory Then
                paths(position) = localpath & "\" & temppath
                position = position + 1
            End If
next_directory_iteration:
            temppath = Dir
        Wend
        iteration_position = iteration_position + 1
    Wend
    
    Screen.MousePointer = 1
    path = ""
    find_file = path
        
End Function

Private Sub Form_Load()
    frmReleaseNotes.Show 1
End Sub

Private Sub mnuabout_Click()
    frmSplash.Show
End Sub

Private Sub mnuDBMan_Click()
    Call cmdDBMan_Click
End Sub

Private Sub mnuDcut_Click()
    Call cmdDCUT_Click
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub

Private Sub mnuPearls_Click()
    Call cmdPearl_Click
End Sub
