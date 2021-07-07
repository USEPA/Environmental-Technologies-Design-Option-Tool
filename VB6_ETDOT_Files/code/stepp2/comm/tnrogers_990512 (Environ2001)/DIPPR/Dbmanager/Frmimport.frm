VERSION 5.00
Begin VB.Form frmimport 
   Caption         =   " Import Data for Existing Chemical"
   ClientHeight    =   5745
   ClientLeft      =   1215
   ClientTop       =   1695
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5745
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Frame frsource 
      Caption         =   "Source"
      Height          =   1215
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   6255
      Begin VB.TextBox tbxsrctbl 
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox tbxsrcdb 
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblbrsource 
         Caption         =   "browse..."
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
         Left            =   5280
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblsrctbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Table"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblsrcdb 
         Alignment       =   1  'Right Justify
         Caption         =   "Database"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frdest 
      Caption         =   "Destination"
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   6255
      Begin VB.TextBox tbxtabledest 
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox tbxdbdest 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblbrdest 
         Caption         =   "browse..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbldesttable 
         Alignment       =   1  'Right Justify
         Caption         =   "Table"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbldestdb 
         Alignment       =   1  'Right Justify
         Caption         =   "Database"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.ListBox lstfields 
      Height          =   1230
      Left            =   3480
      MultiSelect     =   1  'Simple
      TabIndex        =   4
      Top             =   3720
      Width           =   2895
   End
   Begin VB.ListBox lsttable 
      Height          =   1230
      Left            =   240
      MultiSelect     =   1  'Simple
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmdimport 
      Caption         =   "&Import"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label lblchemname 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblcas 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblfields 
      Caption         =   "Available Fields"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lbltables 
      Caption         =   "Available Tables"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmimport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()

    frmimport.Hide
    Unload Me
    frmedit.Show
End Sub

Private Sub cmdimport_Click()
    Call merge_database
End Sub

Private Sub lsttable_Click()

    lstfields.Clear
End Sub


Public Function merge_database() As Boolean

' NOTE: proceed with extreme caution, this is a take-off of the
' merge utility, meant to be edited before use.
' REMEMBER that the chembrowsedb is the current database and is
' already open (and needs to stay open)
    'Dim dbone As Database
    Dim dbtwo As Database
    Dim tableone As Recordset
    Dim tabletwo As Recordset
    Dim tabledest As Recordset
    'Dim dbonename As String
    Dim dbtwoname As String
    Dim table1name As String
    Dim table2name As String
    Dim merge_field As String
    Dim i As Integer
    Dim answer As Integer
    Dim recordcount As Integer
    Dim FNum As Integer
    Dim errorcount As Integer
    Dim curcas As Long
    Dim tabledestname As String
    Dim dbdestname As String
    Dim errorfilename As String
    Dim errormessage As String
    Dim merge_prompt As String
    
    ' the databases, modify here ??????
    dbtwoname = Left(AppPath, 2) & "\master.mdb"
    ' dbtwoname = ""
    table1name = "new_ref_chem"
    table2name = "SMILES Indices"
    'tabledestname = "reference chemicals"
    merge_field = "SMILES"
    merge_prompt = "Confirm merge action: " & Chr(13) & " merging databases: " & Chr(13) & dbtwoname & " and " & dbdestname & Chr(13) & "tables: " & table1name & " and " & table2name & Chr(13) & "fields: " & merge_field
    answer = MsgBox(merge_prompt, vbYesNo)
    If answer = vbNo Then
        GoTo after_merge
    End If
    ' opening the various dbs and tables, modify here
    Set dbtwo = OpenDatabase(dbtwoname, False, False)
    'Set dbtwo = OpenDatabase(dbtwoname, False, False)
    Set tableone = chembrowsedb.OpenRecordset(table1name, dbOpenTable)
    Set tabletwo = dbtwo.OpenRecordset(table2name, dbOpenTable)
    ' Set dbthree = OpenDatabase(dbdestname, False, False)
    'Set tabledest = chembrowsedb.OpenRecordset(tabledestname, dbOpenTable)
    
    Screen.MousePointer = 11
    
    recordcount = 0
    FNum = FreeFile
    errorfilename = "dbmergeerror"
    Open errorfilename For Output As FNum
    
    tableone.MoveFirst
    While Not tableone.EOF
        ' get a CAS # from table one
        On Error Resume Next
        curcas = tableone("CAS")
        
        ' look up the CAS # in table two
        tabletwo.Index = "CASindex"
        tabletwo.Seek "=", curcas
    
        If tabletwo.NoMatch Then
            Write #FNum, curcas
            errorcount = errorcount + 1
            'MsgBox (curcas & " not found")
            GoTo next_iteration
        End If
   
        ' if found put all info in destination table
        ' modify here to addnew or edit
        tableone.Edit
        
        tableone(merge_field) = tabletwo(merge_field)
        
        tableone.Update
next_iteration:
        tableone.MoveNext
        recordcount = recordcount + 1
    Wend
    tableone.Close
    tabletwo.Close
    dbtwo.Close
    'tabledest.Close
    Close FNum
    Screen.MousePointer = 1
    If errorcount > 0 Then
        errormessage = errorcount & " records not matched, see " & errorfilename & " for records not merged" & Chr(13)
    Else
        errormessage = ""
    End If
    MsgBox (errormessage & recordcount & " records successfully merged")
    Exit Function
after_merge:
    
End Function

