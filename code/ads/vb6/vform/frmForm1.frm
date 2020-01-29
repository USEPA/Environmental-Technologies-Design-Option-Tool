VERSION 5.00
Begin VB.Form frmForm1 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   2655
   ClientTop       =   2040
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   6585
   Begin VB.CommandButton Command6 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   3240
      ScaleHeight     =   2955
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frmForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

      '--------------------------------------------------------------------
      ' Capture the entire screen.

      Private Sub Command1_Click()
         Set Picture1.Picture = CaptureScreen()
      End Sub

      ' Capture the entire form including title and border.
      Private Sub Command2_Click()
         Set Picture1.Picture = CaptureForm(Me)
      End Sub

      ' Capture the client area of the form.
      Private Sub Command3_Click()
         Set Picture1.Picture = CaptureClient(Me)
      End Sub

      ' Capture the active window after two seconds.
      Private Sub Command4_Click()
         MsgBox "Two seconds after you close this dialog " & _
            "the active window will be captured."

         ' Wait for two seconds.
         Dim EndTime As Date
         EndTime = DateAdd("s", 2, Now)
         Do Until Now > EndTime
            DoEvents
         Loop

         Set Picture1.Picture = CaptureActiveWindow()

         ' Set focus back to form.
         Me.SetFocus
      End Sub

      ' Print the current contents of the picture box.
      Private Sub Command5_Click()
         PrintPictureToFitPage Printer, Picture1.Picture
         Printer.EndDoc
      End Sub

      ' Clear out the picture box.
      Private Sub Command6_Click()
         Set Picture1.Picture = Nothing
      End Sub

      ' Initialize the form and controls.
      Private Sub Form_Load()
         Me.Caption = "Capture and Print Example"
         Command1.Caption = "&Screen"
         Command2.Caption = "&Form"
         Command3.Caption = "&Client"
         Command4.Caption = "&Active"
         Command5.Caption = "&Print"
         Command6.Caption = "C&lear"
         Picture1.AutoSize = True
      End Sub
      '--------------------------------------------------------------------


