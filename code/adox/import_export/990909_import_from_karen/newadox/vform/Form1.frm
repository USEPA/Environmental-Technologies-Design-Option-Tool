VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "The Form"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      DataField       =   "RecordID"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      DataField       =   "FieldName"
      DataSource      =   "Data1"
      Height          =   525
      Left            =   1740
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\newadox\examples\example1.gap"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Main"
      Top             =   2730
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
