Attribute VB_Name = "moddbman"
Option Explicit

Public Sub main()
' *paul
        
    'Initialize flood gauge and show title screen
    Screen.MousePointer = 11
    frmSplash.Show
    frmSplash.Refresh
    
    dbstatus = STATUS_CLOSED
    Call initialize_globals
    Call init_wizard_globals
    Call read_dbman_file
    
    frmSplash.Hide
    frmdbman.Show
    Unload frmSplash
        
    Screen.MousePointer = 1
End Sub

Public Sub initialize_globals()
' *paul
Dim i As Integer
Dim j As Integer
Dim K As Integer
def_modified = False
dbman_apps = 0

For i = 0 To MAX_DB - 1
    dbman_(i, 0) = ""
    dbman_(i, 1) = ""
Next i

dbpath = ""
curname = ""
curpath = ""
selected_name = ""
selected_cas = 0
selected_structure = ""
selected_smiles = ""

' settings initialized
BIPCode(0) = 1
BIPCode(1) = 2
BIPCode(2) = 3
BIPCode(3) = 4
End Sub

Public Sub read_dbman_file()
' *paul
    ' this function assumes that the only application
    ' that uses lines containing more than one word separated
    ' by spaces is dbman
    ' for now no comment lines
    Dim filename As String
    Dim FNum As Integer
    Dim id_string As String
    Dim localdb As String
    Dim localpath As String
    Dim Save As Boolean
    Dim not_loaded As String
    On Error GoTo RecieveError
    FNum = FreeFile
    filename = "dbman.def"
    'filepath = find_file(filename)
    Open AppPath & filename For Input As FNum
    
    dbman_apps = 0
    not_loaded = ""
    Save = False
    Input #FNum, id_string
    While Not EOF(FNum) And id_string <> "end"
        If Trim(id_string) = "dbman" Then
            Input #FNum, localdb
            Input #FNum, localpath

            If Not FileExists(localpath & localdb) Then
                If LCase(localdb) = "master.mdb" Then
                    localpath = find_file(localdb)
                Else
                    localpath = ""
                End If
                If localpath = "" Then
                    If not_loaded <> "" Then
                        not_loaded = not_loaded & Chr(13) & "      " & localdb
                    Else
                        not_loaded = "      " & localdb
                    End If
                    Save = True
                    GoTo next_file
                End If
            End If
            dbman_(dbman_apps, 0) = localdb
            dbman_(dbman_apps, 1) = localpath
            dbman_apps = dbman_apps + 1
        End If
next_file:
        Input #FNum, id_string
    Wend
end_def_file:
    Close #FNum
    If Save Then
        Call write_dbman_file
    End If
    If not_loaded <> "" Then
        MsgBox "File(s):" & Chr(13) & not_loaded & Chr(13) & "  could not be found."
    End If
    
Exit Sub

RecieveError:
    MsgBox "In ModDBman : read_dbman_file" & Chr(13) & "Error '" & Err & "' " & Error
End Sub

Public Sub write_dbman_file()
' *paul
    Dim filename As String
    Dim textline As String
    Dim fileline(30) As String
    Dim FNum As Integer
    Dim i As Integer
    Dim num_other_lines As Integer
    
    FNum = FreeFile
    'filepath = find_file("ppms.def")
    filename = "dbman.def"
    Open AppPath & filename For Input As #FNum

    ' first read in all the stuff from the .def file and save
    ' stuff not belonging to dbman
    i = 0
    Do While Not EOF(FNum)
        Line Input #1, textline
        If Left(textline, 6) <> Chr(34) & "dbman" Then
            fileline(i) = textline
            i = i + 1
        End If
    Loop
    Close #FNum
    num_other_lines = i

    ' write stuff back to file starting with dbman stuff

    Open AppPath & filename For Output As FNum
    For i = 0 To dbman_apps - 1
                Write #FNum, "dbman", dbman_(i, 0), dbman_(i, 1)
    Next i

    ' write the other lines back
    For i = 0 To num_other_lines - 1
        Print #FNum, fileline(i)
    Next i

    Close #FNum
End Sub


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
    drive_var = Left(AppPath, 1) & ":"
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

Public Sub update_globals(Action As Integer)
' *paul
    ' adds or removes a database
    ' action = 1 means add it
    ' action <> 1 means remove it (not delete it)
Dim i As Integer
Dim found_name As Integer
    ' now set the global arrays
    If Action = 1 Then
            ' add the database
            dbman_(dbman_apps, 0) = curname
            dbman_(dbman_apps, 1) = curpath
            dbman_apps = dbman_apps + 1
    Else
        For i = 0 To dbman_apps - 1
            If UCase(Trim(dbman_(i, 0))) = UCase(Trim(curname)) Then
                found_name = i
                Exit For
            End If
        Next i
        ' remove the database
        If found_name <> -1 Then
            For i = found_name + 1 To MAX_DB
                dbman_(i - 1, 0) = dbman_(i, 0)
                dbman_(i - 1, 1) = dbman_(i, 1)
            Next i
            dbman_apps = dbman_apps - 1
        End If
    End If
End Sub
