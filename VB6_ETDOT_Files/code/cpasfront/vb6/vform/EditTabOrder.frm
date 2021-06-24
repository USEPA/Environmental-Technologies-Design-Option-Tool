VERSION 5.00
Begin VB.Form frmEditTabOrder 
   Caption         =   "Tab Order"
   ClientHeight    =   3900
   ClientLeft      =   1470
   ClientTop       =   3645
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6060
   Begin VB.CommandButton cmdMoveUpDown 
      Caption         =   "Move &Down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4260
      TabIndex        =   4
      Top             =   570
      Width           =   1575
   End
   Begin VB.CommandButton cmdMoveUpDown 
      Caption         =   "Move &Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4260
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox lstTabOrder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3825
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1530
      TabIndex        =   1
      Top             =   3420
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   3420
      Width           =   1245
   End
End
Attribute VB_Name = "frmEditTabOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempProj As ProjectType

Dim USER_HIT_CANCEL As Boolean




Const frmEditTabOrder_declarations_end = 0


Sub frmEditTabOrder_DoEdit( _
    out_USER_HIT_CANCEL As Boolean)
  TempProj = NowProj
  frmEditTabOrder.Show 1
  out_USER_HIT_CANCEL = USER_HIT_CANCEL
  If (Not USER_HIT_CANCEL) Then
    NowProj = TempProj
  End If
End Sub


Sub populate_lstTabOrder()
Dim i As Integer
  lstTabOrder.Clear
  For i = 1 To TempProj.Tabs_Count
    lstTabOrder.AddItem TempProj.Tabs(i).Name
  Next i
  lstTabOrder.ListIndex = 0
End Sub


Sub Refresh_frmEditTabOrder()



End Sub


Private Sub cmdExit_Click(Index As Integer)
  Select Case Index
    Case 0:     'OK.
      USER_HIT_CANCEL = False
      Unload Me
      Exit Sub
    Case 1:     'CANCEL.
      USER_HIT_CANCEL = True
      Unload Me
      Exit Sub
  End Select
End Sub


Sub SwapTabs(idx1 As Integer, idx2 As Integer)
Dim temp As String
Dim temp_tab As TabType
  'SWAP LIST ENTRIES ON LISTBOX CONTROL.
  temp = lstTabOrder.List(idx1)
  lstTabOrder.List(idx1) = lstTabOrder.List(idx2)
  lstTabOrder.List(idx2) = temp
  'SWAP RECORDS IN MEMORY.
  temp_tab = TempProj.Tabs(idx1 + 1)
  TempProj.Tabs(idx1 + 1) = TempProj.Tabs(idx2 + 1)
  TempProj.Tabs(idx2 + 1) = temp_tab
End Sub


Private Sub cmdMoveUpDown_Click(Index As Integer)
Dim idx_now As Integer
  idx_now = lstTabOrder.ListIndex
  Select Case Index
    Case 0:   'MOVE UP.
      If (idx_now <= 0) Then
        Beep
        Exit Sub
      End If
      Call SwapTabs(idx_now - 1, idx_now)
      lstTabOrder.ListIndex = idx_now - 1
    Case 1:   'MOVE DOWN.
      If (idx_now >= lstTabOrder.ListCount - 1) Then
        Beep
        Exit Sub
      End If
      Call SwapTabs(idx_now, idx_now + 1)
      lstTabOrder.ListIndex = idx_now + 1
  End Select
End Sub


Private Sub Form_Load()
  Call CenterOnForm(Me, frmMain)
  Call populate_lstTabOrder
  Call Refresh_frmEditTabOrder
End Sub

