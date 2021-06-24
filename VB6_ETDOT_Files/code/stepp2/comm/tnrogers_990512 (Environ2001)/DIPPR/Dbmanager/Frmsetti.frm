VERSION 5.00
Begin VB.Form frmsettings 
   Caption         =   "DBM Settings"
   ClientHeight    =   4350
   ClientLeft      =   1260
   ClientTop       =   1935
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4350
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frtemp 
      Caption         =   "Temperature Units"
      Height          =   1575
      Left            =   5400
      TabIndex        =   11
      Top             =   120
      Width           =   1575
      Begin VB.ComboBox cbotempunits 
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frcalcdep 
      Caption         =   "Calculation Dependency Preferences"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   6855
      Begin VB.ComboBox cbomethods 
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   600
         Width           =   3015
      End
      Begin VB.ListBox lstproperties 
         Height          =   1035
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "select a preferred method:"
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame frbip 
      Caption         =   "Binary Interaction Parameters"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ListBox lstbipsel 
         Height          =   1035
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdleft 
         Caption         =   "<<"
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
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdright 
         Caption         =   ">"
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
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.ListBox lstbipall 
         Height          =   1035
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()

    frmsettings.Hide
    Unload Me
End Sub

Private Sub cmdleft_Click()

    lstbipsel.Clear
End Sub

Private Sub cmdone_Click()

    Dim I As Integer
    
    For I = 0 To lstbipsel.ListCount - 1
        Select Case Trim(lstbipsel.List(I))
            Case "AGLB"
                BIPCode(I) = 1
            Case "AVLE"
                BIPCode(I) = 2
            Case "AENV"
                BIPCode(I) = 3
            Case "ALLE"
                BIPCode(I) = 4
        End Select
        
    Next I
    frmsettings.Hide
    Unload Me
End Sub

Private Sub cmdright_Click()

    Dim choice As String
    Dim I As Integer
    Dim there As Boolean
    
    For I = 0 To 3
        If lstbipall.Selected(I) = True Then
            choice = Trim(lstbipall.List(I))
            Exit For
        End If
    Next I
    
    there = False
    For I = 0 To lstbipsel.ListCount - 1
        If Trim(lstbipsel.List(I)) = choice Then
            there = True
            Exit For
        End If
    Next I
    
    If there = True Then
        MsgBox ("item already selected, use arrow commands to clear selected list")
    Else
        lstbipsel.AddItem choice
    End If
End Sub


Private Sub Form_Load()

    Dim I As Integer
    Dim J As Integer
    
    lstbipall.Clear
    lstbipall.AddItem "AGLB", 0
    lstbipall.AddItem "AVLE", 1
    lstbipall.AddItem "AENV", 2
    lstbipall.AddItem "ALLE", 3
    
    lstbipsel.Clear
    lstbipsel.AddItem "AGLB"
    lstbipsel.AddItem "AVLE"
    lstbipsel.AddItem "AENV"
    lstbipsel.AddItem "ALLE"
    
    frmsettings!lstproperties.Clear
    frmsettings!cbomethods.Clear
    frmsettings!cbotempunits.Clear
    
    For I = 0 To MAX_DISPLAY_PROPERTIES
        frmsettings!lstproperties.AddItem input_name(I), I
    Next I
    
    frmsettings!cbotempunits.AddItem "C"
    frmsettings!cbotempunits.AddItem "K"
    cbotempunits.ListIndex = 0
End Sub


Private Sub lstproperties_Click()

    Dim I As Integer
    Dim J As Integer
    
    cbomethods.Clear
    I = lstproperties.ListIndex
    For J = 0 To MAX_INPUTS_EACH - 1
        If Trim(wiz_methods(I, J)) <> "" Then
            cbomethods.AddItem wiz_methods(I, J)
        Else
            Exit For
        End If
    Next J
    cbomethods.ListIndex = 0
End Sub


