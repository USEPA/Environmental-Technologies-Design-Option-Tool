VERSION 2.00
Begin Form PathDlg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PathDlg"
   ClientHeight    =   2235
   ClientLeft      =   1320
   ClientTop       =   2265
   ClientWidth     =   6825
   ControlBox      =   0   'False
   Height          =   2640
   Left            =   1260
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6825
   Top             =   1920
   Width           =   6945
   Begin PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   360
      Picture         =   PATH.FRX:0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin CommandButton Command1 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin CommandButton Command2 
      Caption         =   "&Exit Setup"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin Label Label1 
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Install to:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   780
      Width           =   1575
   End
   Begin Label inDrive 
      Caption         =   "inDrive"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin Label outButton 
      Caption         =   "outButton"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin Label Label3 
      Caption         =   "To quit Setup, choose the Exit button."
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   3615
   End
   Begin Label outPath 
      Caption         =   "outPath"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
End

Dim DestPath$

Sub Command1_Click ()
    DestPath$ = text1.Text
    inDrive.Tag = Left$(DestPath$, 2)

    '-------------------------------------------
    ' The IsValidPath function not only returns
    ' True/False as to whether or not it is a valid
    ' path, but also reformats the path variable
    ' into the format, "X:\dir\dir\dir\"
    '-------------------------------------------
    If IsValidPath(DestPath$, (inDrive.Tag)) Then
        OutPath.Tag = DestPath$
        OutButton.Tag = "continue"
        PathDlg.Hide
    Else
        MsgBox "Not a valid path.", 48, PathDlg.Caption
        text1.SetFocus
        text1.SelStart = 0
        text1.SelLength = Len(text1.Text)
    End If
End Sub

Sub Command2_Click ()
    OutButton.Tag = "exit"
    PathDlg.Hide
End Sub

