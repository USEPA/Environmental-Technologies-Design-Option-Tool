VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmmaster 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change/ConfirmDatabase Settings"
   ClientHeight    =   2445
   ClientLeft      =   1035
   ClientTop       =   2430
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2445
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdexit 
      Caption         =   "&exit"
      Height          =   405
      Left            =   5700
      TabIndex        =   9
      Top             =   1830
      Width           =   1875
   End
   Begin VB.Frame frprompt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Database Sources"
      Height          =   1455
      Left            =   180
      TabIndex        =   2
      Top             =   210
      Width           =   7605
      Begin VB.Label txtname 
         Alignment       =   1  'Right Justify
         Caption         =   "Master List Source"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   2445
      End
      Begin VB.Label txtpath 
         Caption         =   "Label1"
         Height          =   285
         Index           =   0
         Left            =   2820
         TabIndex        =   7
         Top             =   330
         Width           =   4605
      End
      Begin VB.Label txtname 
         Alignment       =   1  'Right Justify
         Caption         =   "User List Source"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   630
         Width           =   2445
      End
      Begin VB.Label txtname 
         Alignment       =   1  'Right Justify
         Caption         =   "Block 5 Source"
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   960
         Width           =   2445
      End
      Begin VB.Label txtpath 
         Caption         =   "Label1"
         Height          =   315
         Index           =   1
         Left            =   2820
         TabIndex        =   4
         Top             =   630
         Width           =   4725
      End
      Begin VB.Label txtpath 
         Caption         =   "Label1"
         Height          =   315
         Index           =   2
         Left            =   2820
         TabIndex        =   3
         Top             =   960
         Width           =   4875
      End
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "&browse"
      Height          =   405
      Left            =   3030
      TabIndex        =   1
      Top             =   1860
      Width           =   1995
   End
   Begin VB.CommandButton cmdaccept 
      Caption         =   "&accept"
      Default         =   -1  'True
      Height          =   405
      Left            =   420
      TabIndex        =   0
      Top             =   1860
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdbrowse 
      Left            =   2280
      Top             =   1830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
End
Attribute VB_Name = "frmmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDAccept_Click()

frmmaster.Hide
DoEvents

End Sub

Private Sub cmdbrowse_Click()

    ' This needs to get the entire path with file name into the Path variable,
    ' and just the file name into the filename variable
    Dim i As Integer
    Dim browsefor As Integer
    Dim initial_directory As String
    browsefor = 0

    For i = 0 To 3
        If txtname(i).Font.Bold = True Then
            browsefor = i
            Exit For
        End If
    Next i
    On Error GoTo cancel_error
    Select Case browsefor
        Case 0
            ' the master
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Master Database Selection"
            frmmaster!cdbrowse.filename = "master"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.mdb,*.dbm)|*.mdb;*.dbm"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the master
            ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathMaster = Trim(tempname)
            Path911 = PathMaster
            Path801 = PathMaster
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    MasterDBName = Right(tempname, i - 1)
                    txtpath(0).caption = PathMaster
                    Exit For
                End If
            Next i
        Case 1
            ' user database
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Saved Database Selection"
            frmmaster!cdbrowse.filename = "dbsave"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.mdb;*.prl)|*.mdb;*.prl"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the user
            ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathSave = Trim(tempname)
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    savefile(0) = Right(tempname, i - 1)
                    txtpath(1).caption = PathSave
                    Exit For
                End If
            Next i
       
        Case 2
            ' block 5 database
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Block 5 Database Selection"
            frmmaster!cdbrowse.filename = "block5"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.mdb)|*.mdb"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the block 5
         ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathBlock5 = Trim(tempname)
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    dbblock5file(0) = Right(tempname, i - 1)
                    txtpath(2).caption = PathBlock5
                    Exit For
                End If
            Next i
    End Select
    Exit Sub
cancel_error:
    
End Sub

Private Sub cmdexit_Click()

    Dim answer As Integer
    
    answer = MsgBox("Exit Pearls?", vbYesNo)
    If answer = vbYes Then
        End
    End If
End Sub

Private Sub txtname_Click(Index As Integer)

Dim i As Integer
For i = 0 To 4
    If txtname(i).Font.Bold = True Then
        txtname(i).Font.Bold = False
        txtpath(i).Font.Bold = False
        Exit For
    End If
Next i
If txtname(Index).Visible = True Then
    txtname(Index).Font.Bold = True
    txtpath(Index).Font.Bold = True
End If
End Sub


Private Sub txtname_DblClick(Index As Integer)
Dim i As Integer
For i = 0 To 4
    If txtname(i).Font.Bold = True Then
        txtname(i).Font.Bold = False
        txtpath(i).Font.Bold = False
        Exit For
    End If
Next i
If txtname(Index).Visible = True Then
    txtname(Index).Font.Bold = True
    txtpath(Index).Font.Bold = True
End If

' now show the browser
    On Error GoTo cancel_error
    Select Case Index
        Case 0
            ' the master
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Master Database Selection"
            frmmaster!cdbrowse.filename = "master"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.mdb,*.dbm)|*.mdb;*.dbm"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the master
            ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathMaster = tempname
            Path911 = PathMaster
            Path801 = PathMaster
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    MasterDBName = Right(tempname, i - 1)
                    txtpath(0).caption = PathMaster
                    Exit For
                End If
            Next i
        Case 1
            ' user database
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Saved Database Selection"
            frmmaster!cdbrowse.filename = "dbsave"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.mdb;*.prl)|*.mdb;*.prl"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the user
            ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathSave = Trim(tempname)
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    savefile(0) = Right(tempname, i - 1)
                    txtpath(1).caption = PathSave
                    Exit For
                End If
            Next i
        
        Case 2
            ' block 5 database
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Block 5 Database Selection"
            frmmaster!cdbrowse.filename = "block5"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.mdb)|*.mdb"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the block 5
         ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathBlock5 = Trim(tempname)
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    dbblock5file(0) = Right(tempname, i - 1)
                    txtpath(2).caption = PathBlock5
                    Exit For
                End If
            Next i
    End Select
    Exit Sub
cancel_error:
End Sub


Private Sub txtpath_Click(Index As Integer)
Dim i As Integer
For i = 0 To 4
    If txtname(i).Font.Bold = True Then
        txtname(i).Font.Bold = False
        txtpath(i).Font.Bold = False
        Exit For
    End If
Next i
If txtname(Index).Visible = True Then
    txtname(Index).Font.Bold = True
    txtpath(Index).Font.Bold = True
End If
End Sub


Private Sub txtpath_DblClick(Index As Integer)
Dim i As Integer
For i = 0 To 4
    If txtname(i).Font.Bold = True Then
        txtname(i).Font.Bold = False
        txtpath(i).Font.Bold = False
        Exit For
    End If
Next i
If txtname(Index).Visible = True Then
    txtname(Index).Font.Bold = True
    txtpath(Index).Font.Bold = True
End If

' now show the browser
On Error GoTo cancel_error
    Select Case Index
        Case 0
            ' the master
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Master Database Selection"
            frmmaster!cdbrowse.filename = "master"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.dbm,*.mdb)|*.dbm;*.mdb"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the master
            ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathMaster = Trim(tempname)
            Path911 = PathMaster
            Path801 = PathMaster
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    MasterDBName = Right(tempname, i - 1)
                    txtpath(0).caption = PathMaster
                    Exit For
                End If
            Next i
        Case 1
            ' user database
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Saved Database Selection"
            frmmaster!cdbrowse.filename = "dbsave"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.mdb,*.prl)|*.mdb;*.prl"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the user
            ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathSave = Trim(tempname)
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    savefile(0) = Right(tempname, i - 1)
                    txtpath(1).caption = PathSave
                    Exit For
                End If
            Next i
        
        Case 2
            ' block 5 database
             'Show user available files
            frmmaster!cdbrowse.DialogTitle = "Block 5 Database Selection"
            frmmaster!cdbrowse.filename = "block5"
            frmmaster!cdbrowse.CancelError = True
            frmmaster!cdbrowse.Filter = "(*.mdb)|*.mdb"
            frmmaster!cdbrowse.FilterIndex = 1
            frmmaster!cdbrowse.InitDir = App.path & "\" & Dir(App.path & "\database", vbDirectory)
            frmmaster!cdbrowse.DefaultExt = "mdb"
            frmmaster!cdbrowse.Action = 1
            
            ' if they chose something, use it for the block 5
         ' separate the name from the path
            tempname = Trim(frmmaster!cdbrowse.filename)
            PathBlock5 = Trim(tempname)
            For i = 0 To Len(tempname) - 1
                If Left(Right(tempname, i), 1) = "\" Then
                    dbblock5file(0) = Right(tempname, i - 1)
                    txtpath(2).caption = PathBlock5
                    Exit For
                End If
            Next i
    End Select
    Exit Sub
cancel_error:
End Sub


