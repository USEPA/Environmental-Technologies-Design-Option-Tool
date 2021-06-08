VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmPopupWindow 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmPopupWindow"
   ClientHeight    =   2100
   ClientLeft      =   1005
   ClientTop       =   7020
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   765
      Left            =   1110
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   1349
      _StockProps     =   14
      Caption         =   "Invisible"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picTest 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         ScaleHeight     =   285
         ScaleWidth      =   1185
         TabIndex        =   2
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Label lblText 
      BackColor       =   &H80000018&
      Caption         =   "lblText"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   1290
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3075
   End
End
Attribute VB_Name = "frmPopupWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MARGIN_lblText_HORIZ = 100
Const MARGIN_lblText_VERT = 50
Const WIDTH_WINDOW = 5000

Const MARGIN_PLACEMENT_HORIZ = 800
Const MARGIN_PLACEMENT_VERT = 800


Const frmPopupWindow_declarations_end = 0


Function Remove_Consecutive_Spaces( _
    msg_in As String) _
    As String
Dim msg As String
Dim i As Integer
Dim found As Boolean
  'REMOVE INITIAL AND TERMINAL SPACES.
  msg = Trim$(msg_in)
  'REMOVE EXTRA CONSECUTIVE SPACES.
  Do While (1 = 1)
    found = False
    For i = 1 To Len(msg) - 1
      If (Mid$(msg, i, 1) = " ") Then
        If (Mid$(msg, i + 1, 1) = " ") Then
          found = True
          msg = Left$(msg, i) & Right$(msg, Len(msg) - i - 1)
          Exit For
        End If
      End If
    Next i
    If (Not found) Then Exit Do
  Loop
  Remove_Consecutive_Spaces = msg
End Function


'Function Get_Correct_Height( _
'    msg_in As String, _
'    MaxWidth As Double, _
'    picX As Control) As Double
'Dim ThisWord As String
'Dim ThisC As String
'Dim ThisLine As String
'Dim TestLine As String
'Dim ThisLineNumber As Integer
'Dim i As Integer
'Dim msg As String
'Dim EachLine_Height As Double
'Dim TotalLines_Height As Double
'  'PARSE OUT EACH WORD.
'  ThisWord = ""
'  ThisLine = ""
'  ThisLineNumber = 1
'  msg = msg_in
'  For i = 1 To Len(msg) + 1
'    If (i <= Len(msg)) Then
'      ThisC = Mid$(msg, i, 1)
'    Else
'      ThisC = ""
'    End If
'    If (ThisC = " ") Or (i > Len(msg)) Then
'      TestLine = ThisLine
'      If (ThisLine <> "") Then
'        TestLine = TestLine & " "
'      End If
'      TestLine = TestLine & ThisWord
'      If (picX.TextWidth(TestLine) > MaxWidth) Then
'Debug.Print ThisLine
'        ThisLine = ""
'        ThisLineNumber = ThisLineNumber + 1
'        ThisLine = ThisWord
'      Else
'        ThisLine = TestLine
'      End If
'      ThisWord = ""
'    Else
'      ThisWord = ThisWord & ThisC
'    End If
'  Next i
'Debug.Print ThisLine
'  ThisLineNumber = ThisLineNumber + 1
'  EachLine_Height = CDbl(picX.TextHeight("WWW"))
'  TotalLines_Height = CDbl(ThisLineNumber) * EachLine_Height
'  Get_Correct_Height = CDbl(TotalLines_Height)
'frmMain.Caption = Trim$(Str$(ThisLineNumber))
'End Function


Sub Get_Correct_Height( _
    msg_in As String, _
    lblX As Control, _
    picX As Control)
Dim ThisWord As String
Dim thisc As String
Dim ThisLine As String
Dim TestLine As String
Dim ThisLineNumber As Integer
Dim i As Integer
Dim msg As String
Dim EachLine_Height As Double
Dim TotalLines_Height As Double
Dim EachLineStr() As String
Dim MaxWidth As Long
Dim AllOutputText As String
  MaxWidth = lblX.Width
  'PARSE OUT EACH WORD.
  ThisWord = ""
  ThisLine = ""
  ThisLineNumber = 1
  msg = msg_in
  For i = 1 To Len(msg) + 1
    If (i <= Len(msg)) Then
      thisc = Mid$(msg, i, 1)
    Else
      thisc = ""
    End If
    If (thisc = " ") Or (i > Len(msg)) Then
      TestLine = ThisLine
      If (ThisLine <> "") Then
        TestLine = TestLine & " "
      End If
      TestLine = TestLine & ThisWord
      If (picX.TextWidth(TestLine) > MaxWidth) Then
'Debug.Print ThisLine
        ReDim Preserve EachLineStr(1 To ThisLineNumber)
        EachLineStr(ThisLineNumber) = ThisLine
        ThisLine = ""
        ThisLineNumber = ThisLineNumber + 1
        ThisLine = ThisWord
      Else
        ThisLine = TestLine
      End If
      ThisWord = ""
    Else
      ThisWord = ThisWord & thisc
    End If
  Next i
'Debug.Print ThisLine
  ReDim Preserve EachLineStr(1 To ThisLineNumber)
  EachLineStr(ThisLineNumber) = ThisLine
  'SET UP THE WINDOW.
  'ThisLineNumber = ThisLineNumber + 1
  EachLine_Height = CDbl(picX.TextHeight("WWW"))
  TotalLines_Height = CDbl(ThisLineNumber) * EachLine_Height
  lblX.Height = CDbl(TotalLines_Height)
  AllOutputText = ""
  For i = 1 To ThisLineNumber
    If (Len(AllOutputText) > 0) Then
      AllOutputText = AllOutputText & vbCrLf
    End If
    AllOutputText = AllOutputText & EachLineStr(i)
  Next i
  lblX.Caption = AllOutputText
'frmMain.Caption = Trim$(Str$(ThisLineNumber))
End Sub


Sub frmPopupWindow_Show( _
    msg_in As String, _
    x_in As Double, _
    y_in As Double)
Dim msg As String
Dim ThisText_Height As Double
  'DO _NOT_ DISPLAY AN EMPTY STRING!
  If (Trim$(msg_in) = "") Then Exit Sub
  'SET UP LABEL CONTROL.
  msg = Remove_Consecutive_Spaces(msg_in)
  'lblText.Caption = msg
  Me.Width = WIDTH_WINDOW
  lblText.Left = MARGIN_lblText_HORIZ
  lblText.Top = MARGIN_lblText_VERT
  lblText.Width = Me.ScaleWidth - 2 * MARGIN_lblText_HORIZ
  'ThisText_Height = Get_Correct_Height( _
      msg, _
      lblText.Width, _
      picTest)
  Call Get_Correct_Height( _
      msg, _
      lblText, _
      picTest)
  'lblText.Height = ThisText_Height
  Me.Height = lblText.Height + 2 * MARGIN_lblText_VERT + _
      (Me.Height - Me.ScaleHeight)
  'REMOVE CAPTION FROM WINDOW; THIS COMPLETELY
  'REMOVES THE TITLE BAR.
  Me.Caption = ""
  'DETERMINE WHERE TO PLACE THE WINDOW.
Dim Try_Left As Double
Dim Try_Top As Double
Dim Use_Left As Double
Dim Use_Top As Double
  '---- DETERMINE HORIZONTAL PLACEMENT.
  '-------- TRY TO THE RIGHT.
  Try_Left = x_in + MARGIN_PLACEMENT_HORIZ
  If (Try_Left + Me.Width) > Screen.Width Then
    '-------- USE TO THE LEFT.
    Try_Left = x_in - MARGIN_PLACEMENT_HORIZ - Me.Width
  End If
  '---- DETERMINE HORIZONTAL PLACEMENT.
  '-------- TRY BELOW.
  Try_Top = y_in + 0
  If (Try_Top + Me.Height) > Screen.Height Then
    '-------- USE ABOVE.
    Try_Top = y_in - Me.Height
  End If
  'MOVE TO THE NEW POSITION.
  Use_Left = Try_Left
  Use_Top = Try_Top
  Me.Move Use_Left, Use_Top
  'SHOW THE FINISHED PRODUCT.
  frmPopupWindow.Show 0
End Sub


Sub frmPopupWindow_Hide()
  frmPopupWindow.Hide
  'frmPopupWindow.Visible = False
End Sub


Private Sub Image1_Click()

End Sub

