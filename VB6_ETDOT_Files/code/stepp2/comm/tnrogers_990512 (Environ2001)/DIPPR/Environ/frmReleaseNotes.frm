VERSION 5.00
Begin VB.Form frmReleaseNotes 
   Caption         =   "Release Notes"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   Icon            =   "frmReleaseNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Software testing and correction of identified bugs."
      Height          =   195
      Left            =   960
      TabIndex        =   24
      Top             =   4260
      Width           =   4335
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   780
      TabIndex        =   25
      Top             =   4260
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   $"frmReleaseNotes.frx":0442
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   4620
      Width           =   5535
   End
   Begin VB.Label Label12 
      Caption         =   "User login, functionality and stability currently being modified."
      Height          =   195
      Left            =   960
      TabIndex        =   13
      Top             =   4020
      Width           =   4335
   End
   Begin VB.Label Label11 
      Caption         =   "Continuing work:"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Continuing work to add more group contribution methods to DBMAN."
      Height          =   435
      Left            =   960
      TabIndex        =   11
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label9 
      Caption         =   "Continued testing of unit conversions, especially for Henry's Constants and solubility."
      Height          =   435
      Left            =   960
      TabIndex        =   10
      Top             =   3180
      Width           =   4335
   End
   Begin VB.Label Label8 
      Caption         =   "Updated splash screens using the ENVIRON 2001 name"
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   2580
      Width           =   4335
   End
   Begin VB.Label Label7 
      Caption         =   "Additional testing and review of unit conversions"
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   2340
      Width           =   4335
   End
   Begin VB.Label Label6 
      Caption         =   "Project 911 values already within the master database  -  no need to run DCUT to link ENVIRON 2001 to the Project 911 data"
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "Ability to import the Project 801 Access format file as well as the flat file using DCUT"
      Height          =   435
      Left            =   960
      TabIndex        =   6
      Top             =   1260
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "Reformatting and added functionality within DBMAN"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   1020
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Ability to export or print Antoine Coefficient information"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   780
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Temperature dependent correlations calculations corrected"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   540
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Improvements implemented:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   780
      TabIndex        =   23
      Top             =   4020
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   780
      TabIndex        =   22
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   780
      TabIndex        =   21
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   780
      TabIndex        =   20
      Top             =   2580
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   780
      TabIndex        =   19
      Top             =   2340
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   780
      TabIndex        =   18
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   780
      TabIndex        =   17
      Top             =   1260
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   780
      TabIndex        =   16
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   780
      TabIndex        =   15
      Top             =   780
      Width           =   255
   End
   Begin VB.Label lbldot 
      Caption         =   "dot"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   780
      TabIndex        =   1
      Top             =   540
      Width           =   255
   End
End
Attribute VB_Name = "frmReleaseNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
    
    For i = 0 To lbldot.Count - 1
        lbldot.Item(i) = Chr(183)
    Next
End Sub

