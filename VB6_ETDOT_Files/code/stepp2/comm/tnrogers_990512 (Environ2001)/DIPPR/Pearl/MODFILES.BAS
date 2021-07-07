Attribute VB_Name = "modfiles"
Option Explicit
Global Const NULLPATH = "e"
Global Const MAXEXPORTLINES = 20
Global Const MAXCOMMENTLINES = 20

Public Function AddBackSlash(path As String) As String
    If Right(path, 1) = "\" Then
        AddBackSlash = path
    Else
        AddBackSlash = path + "\"
    End If
    
End Function
Public Sub test_paths()
    Dim errorflag As Integer
    Dim whichfile As Integer
    Dim filefound As String
    Dim temptabledef As TableDef
    Dim TestRecords As Recordset
    Dim i As Integer
    Dim numrecords As Long
    Dim tablefound As Boolean
    Dim whicherror
    
    ' this is a generic file and database path checker
    ' to be called before any files or databases are opened
    ' to make sure the existing paths are valid
    ' if not, a directory browser is opened to allow
    ' the user to set the required path(s)
    ' BUGS:  checks that the tables and files are there
    '       but not that the fields are valid, an invalid
    '       field error must be handled by the procedure
    '       opening the file/table
    ' there's a little bit of redundancy here - we set the
    ' path to NULLPATH if there's an error on the Dir command OR
    ' if the Dir command doesn't return a filename.
    ' check for required files
    ' 'whichfile' just keeps track of where we are when we
    '       get an error
    
        ' reports
        On Error GoTo get_which_file
       
    
    whichfile = 4
    If PathMaster = NULLPATH Then GoTo get_which_file
        
    tablefound = False
    Set DBJetMaster = OpenDatabase(PathMaster, False, True)
    
    For i = 0 To DBJetMaster.TableDefs.count - 1
            ' arbitrarily chose this table to check for
        If DBJetMaster.TableDefs(i).name = "UNIFAC" Then
            ' if there's something in the table, we're good
            Set TestRecords = DBJetMaster.OpenRecordset("UNIFAC")
            If TestRecords.EOF = False Then
                tablefound = True
                Exit For
            Else
                PathMaster = NULLPATH
                Path911 = NULLPATH
                Path801 = NULLPATH
                Exit For
            End If
        End If
    Next i
    DBJetMaster.Close
        
        
testdemo:
            ' this test needs to be more robust
        whichfile = 7
        If PathDemo = NULLPATH Then GoTo get_which_file
        filefound = Dir(PathDemo)
        If filefound = "" Then
            PathDemo = NULLPATH
        End If
        
testuser:
            ' this test needs to be more robust
        whichfile = 8
        If PathUser = NULLPATH Then GoTo get_which_file
        filefound = Dir(PathUser)
        If filefound = "" Then
            PathUser = NULLPATH
        End If
        'Exit Sub
testsave:
            ' this test needs to be more robust
        whichfile = 9
        If PathSave = NULLPATH Then GoTo get_which_file
        filefound = Dir(PathSave)
        If filefound = "" Then
            PathSave = NULLPATH
        End If
        'Exit Sub
testblock5:
            ' this test needs to be more robust
        whichfile = 10
        If PathBlock5 = NULLPATH Then GoTo get_which_file
        filefound = Dir(PathBlock5)
        If filefound = "" Then
            PathBlock5 = NULLPATH
        End If
        Exit Sub
        ' this simply sets the appropriate path to NULLPATH and sends control back to the
        ' next checkpoint in the event of an error
get_which_file:
    Select Case whichfile
       
        Case 4:
            PathMaster = NULLPATH
            GoTo testdemo
        Case 7:
            PathDemo = NULLPATH
            GoTo testuser
        Case 8:
            PathUser = NULLPATH
            GoTo testsave
        Case 9:
            PathSave = NULLPATH
            GoTo testblock5
        Case 10:
            PathBlock5 = NULLPATH
        Case Else
            ' what to do??
        End Select
End Sub

Public Sub parse_def_file()

    ' NOTE: this assumes that the following order in the
    '   def file:
    '       1.  comments
    '       2.  paths
    '       3.  export settings
    
    ' a function to parse the pearls.def file
    ' if it's there, if not, the user needs to
    ' set the paths for required files
    ' the GLOBALS used are: PathReport, Path801, Path911, PathMaster
    '           PathUser and PathDemo and PathBlock5
    ' REVISIONS:    4/18/97    DMW  - added Block5 stuff
    '               5/31/97    DMW  - made def file name a variable
    '               3/13/98    DMW  - took out reference to dbuser, demo, and report
    
    Dim FNum As Integer
    Dim fname As String
    Dim setting As String
    Dim answer As Integer
    Dim exitflag As Boolean
    Dim temp As String
    On Error GoTo error_reading_file_closed
    fname = App.path & "\" & deffile
    FNum = FreeFile
    Open fname For Input As #FNum
    On Error GoTo error_reading_file_open
        
        ' some simple checks for problems in the file
        ' it has to be a comment (indicated by ') or start
        ' with "Pearls"
    While Not EOF(FNum)
    
        Input #FNum, setting
             ' if it's a comment, go on to the next line
        If (Left(setting, 1)) = "'" Then
            GoTo continue_loop
        End If
            ' if it's the cue line, go on to next line
        If Trim(setting) = "Pearls.Cue" Then
            GoTo continue_loop
        End If
        If (Left(setting, 7)) <> "Pearls." Then
            answer = MsgBox("Your default file has an error, rebuild file?", vbYesNo)
            If answer = vbYes Then
                Close #FNum
                If UCase(MasterDBName) = UCase(default_master_name) Then
                    PathMaster = NULLPATH
                    Path911 = NULLPATH
                    Path801 = NULLPATH
                End If
                PathSave = NULLPATH
                PathBlock5 = NULLPATH
                Exit Sub
            ElseIf answer = vbNo Then
                answer = MsgBox("Quit Pearls?", vbYesNo)
                If answer = vbNo Then
                    Close #FNum
                    If UCase(MasterDBName) = UCase(default_master_name) Then
                        PathMaster = NULLPATH
                        Path911 = NULLPATH
                        Path801 = NULLPATH
                    End If
                    PathSave = NULLPATH
                    PathBlock5 = NULLPATH
                    Exit Sub
                Else
                    Close #FNum
                    exitflag = True
                End If
            End If
        End If
        
            ' get the paths out of the file and set the
            ' globals if they're there
            ' first we're stripping off the Pearls part
            ' NOTE:  we're assuming each path/line is ended with a '.'
        setting = Right(setting, Len(setting) - 7)
        If Left(setting, 9) = "Database." Then
            setting = Right(setting, Len(setting) - 9)
            If UCase(MasterDBName) = UCase(default_master_name) And Left(setting, 7) = "Master." Then
                PathMaster = Right(setting, Len(setting) - 7)
                If Right(PathMaster, 1) = "." Then
                    PathMaster = Left(PathMaster, Len(PathMaster) - 1)
                    Path911 = PathMaster
                    Path801 = PathMaster
                End If
           
            ElseIf Left(setting, 5) = "Save." Then
                PathSave = Right(setting, Len(setting) - 5)
                If Right(PathSave, 1) = "." Then
                    PathSave = Left(PathSave, Len(PathSave) - 1)
                End If
                If PathSave = "None" Then
                    PathSave = NULLPATH
                End If
                ' now get the actual file name
                
            ElseIf Left(setting, 7) = "Block5." Then
                PathBlock5 = Right(setting, Len(setting) - 7)
                If Right(PathBlock5, 1) = "." Then
                    PathBlock5 = Left(PathBlock5, Len(PathBlock5) - 1)
                End If
                If PathBlock5 = "None" Then
                    PathBlock5 = NULLPATH
                
                End If
            End If
        End If
continue_loop:
    Wend
    Close #FNum
    Exit Sub
error_reading_file_open:
    
        PathMaster = NULLPATH
        Path911 = NULLPATH
        Path801 = NULLPATH
        PathSave = NULLPATH
        PathBlock5 = NULLPATH
        Close #FNum
        Exit Sub
error_reading_file_closed:
    
        PathMaster = NULLPATH
        Path911 = NULLPATH
        Path801 = NULLPATH
        PathSave = NULLPATH
        PathBlock5 = NULLPATH
End Sub

Public Sub write_def_file()
    ' write_def_file:  writes the paths for various files to
    '       the defs file.  Note that the defs file has to be
    '       in the application directory.
    '       we're expecting all the file path stuff to be in the beginning of
    '       the defs file (in no particular order) with the export stuff after
    '       This function needs to preserve the export stuff in
    '       the defs file as well as any comments
    Dim fname As String
    Dim fileline As String
    Dim exportline(MAXEXPORTLINES) As String
    Dim commentline(MAXCOMMENTLINES) As String
    Dim cueline As String
    Dim numexportlines As Integer
    Dim numcommentlines As Integer
    Dim FNum As Integer
    Dim J As Integer
    Dim K As Integer
    
    fname = App.path & "\" & deffile
    FNum = FreeFile
    On Error GoTo error_reading_file
        
    
        ' first get the comments and export lines so we don't lose those
    J = 0
    numcommentlines = 0
    numexportlines = 0
    cueline = ""
    Open fname For Input As #FNum
        While Not EOF(FNum)
            Input #FNum, fileline
            If Left(fileline, 14) = "Pearls.Export." Then
                exportline(J) = fileline
                J = J + 1
            ElseIf Left(fileline, 1) = "'" Then
                commentline(K) = fileline
                K = K + 1
            ElseIf Left(fileline, 10) = "Pearls.Cue" Then
                cueline = fileline
            End If
        Wend
        numexportlines = J
        numcommentlines = K
    Close #FNum
    
        ' now write everything into the file, the export
        ' lines aren't checked, just written back into the file
        ' as is
        ' NOTE: this assumes that the following order in the
        '   def file:
        '       1.  comments
        '       2.  paths (for master, block5, save, in no part. order
        '       3.  export settings
        
writefile:
    On Error GoTo error_writing_file
    Open fname For Output As #FNum
        ' first write the comments
    For J = 0 To numcommentlines - 1
        Write #FNum, commentline(J)
    Next J
        ' now the path stuff
    
    If PathMaster <> NULLPATH Then
        Write #FNum, "Pearls.Database.Master." & PathMaster & "."
    Else
        Write #FNum, "Pearls.Database.Master.None" & "."
    End If
   
    If PathSave <> NULLPATH Then
        Write #FNum, "Pearls.Database.Save." & PathSave & "."
    Else
        Write #FNum, "Pearls.Database.Save.None" & "."
    End If
    If PathBlock5 <> NULLPATH Then
        Write #FNum, "Pearls.Database.Block5." & PathBlock5 & "."
    Else
        Write #FNum, "Pearls.Database.Block5.None" & "."
    End If
        ' now add the export stuff
    For J = 0 To numexportlines - 1
        Write #FNum, exportline(J)
    Next J
        ' now add the cue line if it was there
    If cueline <> "" Then
        Write #FNum, cueline
    End If
    Close #FNum
    Exit Sub
    
error_reading_file:
    GoTo writefile
error_writing_file:
    MsgBox ("Unable to write to " & deffile & " file, continuing with PEARLS execution")
    Exit Sub
End Sub

Public Function simple_check_files() As Boolean

    Dim errorstring As String
    Dim success As Boolean
    Dim answer As String
    errorstring = "Files Not Found:" & Chr(13)
    success = True
    answer = ""
    If PathMaster = NULLPATH Or Not Right(PathMaster, 4) Like ".*" Then
        answer = Dir(App.path & "\master.mdb")
        If Trim(answer) = "" Then
            success = False
        Else
            PathMaster = App.path & "\master.mdb"
        End If
        'errorstring = errorstring & Chr(9) & "master database" & Chr(13)
        
    End If
    If PathSave = NULLPATH Or Not Right(PathSave, 4) Like ".*" Then
        answer = Dir(App.path & "\dbsave.mdb")
        If Trim(answer) = "" Then
            success = False
        Else
            PathSave = App.path & "\dbsave.mdb"
        End If
        'errorstring = errorstring & Chr(9) & "saved database" & Chr(13)
        
    End If
    If PathBlock5 = NULLPATH Or Not Right(PathBlock5, 4) Like ".*" Then
        answer = Dir(App.path & "\block5.mdb")
        If Trim(answer) = "" Then
            success = False
        Else
            PathBlock5 = App.path & "\block5.mdb"
        End If
        'errorstring = errorstring & Chr(9) & "report files" & Chr(13)
        
    End If
    simple_check_files = success
    
End Function
