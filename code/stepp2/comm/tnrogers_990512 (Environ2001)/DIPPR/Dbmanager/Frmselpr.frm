VERSION 5.00
Begin VB.Form frmselprop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Property Selection"
   ClientHeight    =   1725
   ClientLeft      =   1920
   ClientTop       =   1845
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1725
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmddone 
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   765
      TabIndex        =   2
      Top             =   1215
      Width           =   2535
   End
   Begin VB.Frame frselprop 
      Caption         =   "select from..."
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboprop 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmselprop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmddone_Click()

    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim I As Integer
    Dim new_grouptype As String
    Dim label As String
    Screen.MousePointer = 11
    For I = 0 To MAX_PROPERTIES - 1
        If Trim(cboprop.List(cboprop.ListIndex)) = input_name(I) Then
            global_cur_property = I
            Exit For
        End If
    Next I
    Call update_edwiz_method_info
    Call update_edwiz_input_info(0)
    ' if the group type has changed, fix that to current input groups needed
   For I = 0 To MAX_INPUTS_EACH - 1
        If UCase(Right(Trim(frmeditwizard!lblinputprop(I)), 6)) = "GROUPS" Then
            label = Trim(frmeditwizard!lblinputprop(I))
            new_grouptype = Trim(Left(label, Len(label) - 6))
            frmeditwizard!frgroups.Caption = Trim(new_grouptype) & " Groups"
            Exit For
        End If
    Next I
    ' don't need to do this here, gets done by update_edwiz_input_info
    'Call reset_groups(new_grouptype)
   
    frmeditwizard.Caption = selected_name & ":  " & input_name(global_cur_property)
    Screen.MousePointer = 1
End Sub
