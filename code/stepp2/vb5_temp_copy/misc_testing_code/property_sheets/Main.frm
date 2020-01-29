VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   2580
   ClientTop       =   2535
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   8235
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   690
      TabIndex        =   0
      Top             =   960
      Width           =   2205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  List1.Clear
  List1.AddItem "Chemical Information"
  List1.AddItem "General 1"
  List1.AddItem "General 2"
  List1.AddItem "Transport"
  List1.AddItem "Partitioning/Equilibrium"
  List1.AddItem "Fire and Explosion"
  List1.AddItem "Oxygen Demand"
  List1.AddItem "Aquatic Toxicity 1"
  List1.AddItem "Aquatic Toxicity 2"
  List1.ListIndex = 0
End Sub
