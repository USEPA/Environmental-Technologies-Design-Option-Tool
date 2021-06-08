VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   960
   ClientTop       =   1020
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6450
   ScaleWidth      =   8220
   Begin VB.Frame Frame2 
      Caption         =   "Search Option"
      Height          =   1575
      Left            =   3720
      TabIndex        =   17
      Top             =   2760
      Width           =   2535
      Begin VB.OptionButton CT 
         Caption         =   "Combinatorial, Truncating"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton SNT 
         Caption         =   "Sequential, Non-Truncating"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton ST 
         Caption         =   "Sequential, Truncating"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Query Type"
      Height          =   1095
      Left            =   3720
      TabIndex        =   14
      Top             =   1560
      Width           =   1815
      Begin VB.OptionButton OFile 
         Caption         =   "File"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton OString 
         Caption         =   "String"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.FileListBox filList 
      Height          =   1260
      Left            =   3720
      TabIndex        =   13
      Top             =   120
      Width           =   1815
   End
   Begin VB.DirListBox dirList 
      Height          =   1665
      Left            =   5640
      TabIndex        =   12
      Top             =   480
      Width           =   1575
   End
   Begin VB.DriveListBox drvList 
      Height          =   315
      Left            =   5640
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Outfile 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Subfile 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Detections 
      Height          =   3075
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox SResult 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox SMILES 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "c1ccccc1"
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   555
      Left            =   4560
      TabIndex        =   0
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Output File Name"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Substructure File"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Subfragment Detections"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Search Result"
      Height          =   195
      Left            =   4920
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "SMILES string"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim lfcr
Dim i As Integer
Dim mySMILES As String
Dim QueryType As Byte
Dim SearchType As Byte
Dim SearchResult As Byte
Dim intSF_ID() As Long, intSF_Quant() As Long
Dim intMF_ID() As Long, intMF_Quant() As Long

lfcr = Chr(13) + Chr(10)
SearchResult = 0

ReDim intSF_ID(0 To 99) As Long, intSF_Quant(0 To 99) As Long
ReDim intMF_ID(0 To 20) As Long, intMF_Quant(0 To 20) As Long

If (OString.Value) Then
    QueryType = 0
ElseIf (OFile.Value) Then
    QueryType = 1
End If

If (SNT.Value) Then
    SearchType = 0
ElseIf (ST.Value) Then
    SearchType = 1
ElseIf (CT.Value) Then
    SearchType = 2
End If

Call MOSDAP(SMILES.Text, QueryType, Subfile.Text, Outfile.Text, SearchType, SearchResult, intSF_ID(0), intSF_Quant(0), intMF_ID(0), intMF_Quant(0))

SResult.Text = SearchResult
Detections.Text = Detections.Text + SMILES.Text + lfcr

Detections.Text = Detections.Text + "--- SF ---" + lfcr
For i = 0 To 20 Step 1
   If (intSF_ID(i) > 0) Then
        Detections.Text = Detections.Text + " " + Str(intSF_ID(i)) + " " + Str(intSF_Quant(i)) + " "
    ElseIf (intSF_ID(i) = -1) Then
        Detections.Text = Detections.Text + lfcr
        i = i + 1
    End If
Next i
Detections.Text = Detections.Text + lfcr
Detections.Text = Detections.Text + "--- MF ---" + lfcr
For i = 0 To 20 Step 1
    If (intMF_ID(i) > 0) Then
        Detections.Text = Detections.Text + " " + Str(intMF_ID(i)) + " " + Str(intMF_Quant(i)) + " "
    ElseIf (intMF_ID(i) = -1) Then
        Detections.Text = Detections.Text + lfcr
        i = i + 1
    End If
Next i

Detections.Text = Detections.Text + lfcr + lfcr

End Sub


Private Sub dirList_Change()
     ' Update the file list box to synchronize with the directory list box.
    filList.Path = dirList.Path
    Subfile.Text = filList.filename
End Sub

Private Sub drvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub filList_Click()
    Subfile.Text = filList.Path + "\" + filList.filename

End Sub

