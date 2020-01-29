VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form frmgraph 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PEARLS Graph"
   ClientHeight    =   7485
   ClientLeft      =   1725
   ClientTop       =   1605
   ClientWidth     =   9270
   ControlBox      =   0   'False
   Icon            =   "frmgrph.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7485
   ScaleWidth      =   9270
   Begin MSChartLib.MSChart GRHChem 
      Height          =   6495
      Left            =   360
      OleObjectBlob   =   "frmgrph.frx":030A
      TabIndex        =   2
      Top             =   240
      Width           =   8565
   End
   Begin VB.CommandButton CMDClose 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   6960
      Width           =   1095
   End
End
Attribute VB_Name = "frmgraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDClose_Click()

    Unload Me
    
End Sub


Private Sub CMDPrint_Click()

On Error GoTo errorprint
    Me.PrintForm

'    Screen.MousePointer = 11
'    FRMGraph!GRHChem.DrawMode = 5
'    Screen.MousePointer = 1

errorprint:

    MsgBox Error$
    GRHChem.Legend.Location.LocationType = VtChLocationTypeRight

End Sub

Private Sub Form_Load()
    
    CenterForm Me
    
End Sub











