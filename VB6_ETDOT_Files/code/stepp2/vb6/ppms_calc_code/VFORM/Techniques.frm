VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTechniques 
   Caption         =   "Techniques for {Property Name}"
   ClientHeight    =   6075
   ClientLeft      =   2790
   ClientTop       =   1590
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11985
   Begin Threed.SSPanel sspTech 
      Height          =   5805
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   8355
      _Version        =   65536
      _ExtentX        =   14737
      _ExtentY        =   10239
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel sspTechMisc 
         Height          =   1425
         Left            =   2970
         TabIndex        =   3
         Top             =   990
         Width           =   5205
         _Version        =   65536
         _ExtentX        =   9181
         _ExtentY        =   2514
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame ssfTechMisc 
            Height          =   1275
            Left            =   60
            TabIndex        =   25
            Top             =   60
            Width           =   5055
            _Version        =   65536
            _ExtentX        =   8916
            _ExtentY        =   2249
            _StockProps     =   14
            Caption         =   "Miscellaneous:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSPanel ssfTechGrid 
         Height          =   705
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   1244
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComctlLib.ListView lvMain 
            Height          =   585
            Left            =   60
            TabIndex        =   1
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1032
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin Threed.SSPanel sspSep 
         Height          =   165
         Left            =   60
         TabIndex        =   5
         Top             =   810
         Width           =   5595
         _Version        =   65536
         _ExtentX        =   9869
         _ExtentY        =   291
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel sspTechTemp 
         Height          =   4485
         Left            =   60
         TabIndex        =   6
         Top             =   990
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   7911
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame ssfTechTemp 
            Height          =   4275
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   7541
            _StockProps     =   14
            Caption         =   "Temperature Dependency, f(T):"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox cboEqForm 
               Height          =   315
               Left            =   1380
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   270
               Width           =   1275
            End
            Begin VB.ComboBox cboUnits_f 
               Height          =   315
               Left            =   1380
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   1005
               Width           =   1275
            End
            Begin VB.ComboBox cboUnits_T 
               Height          =   315
               Left            =   1380
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   645
               Width           =   1275
            End
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   7
               Left            =   1380
               TabIndex        =   38
               Text            =   "txtData()"
               Top             =   3540
               Width           =   1275
            End
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   6
               Left            =   1380
               TabIndex        =   36
               Text            =   "txtData()"
               Top             =   3180
               Width           =   1275
            End
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   5
               Left            =   1380
               TabIndex        =   34
               Text            =   "txtData()"
               Top             =   2820
               Width           =   1275
            End
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   4
               Left            =   1380
               TabIndex        =   32
               Text            =   "txtData()"
               Top             =   2460
               Width           =   1275
            End
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   3
               Left            =   1380
               TabIndex        =   30
               Text            =   "txtData()"
               Top             =   2100
               Width           =   1275
            End
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   2
               Left            =   1380
               TabIndex        =   28
               Text            =   "txtData()"
               Top             =   1740
               Width           =   1275
            End
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   1
               Left            =   1380
               TabIndex        =   26
               Text            =   "txtData()"
               Top             =   1380
               Width           =   1275
            End
            Begin VB.Label lbl_cboEqForm 
               Alignment       =   1  'Right Justify
               Caption         =   "Equation form:"
               Height          =   225
               Left            =   90
               TabIndex        =   45
               Top             =   330
               Width           =   1215
            End
            Begin VB.Label lbl_cboUnits_f 
               Alignment       =   1  'Right Justify
               Caption         =   "Units of f:"
               Height          =   225
               Left            =   90
               TabIndex        =   42
               Top             =   1065
               Width           =   1215
            End
            Begin VB.Label lbl_cboUnits_T 
               Alignment       =   1  'Right Justify
               Caption         =   "Units of T:"
               Height          =   225
               Left            =   90
               TabIndex        =   40
               Top             =   705
               Width           =   1215
            End
            Begin VB.Label lblData 
               Alignment       =   1  'Right Justify
               Caption         =   "Coefficient ""E"":"
               Height          =   225
               Index           =   7
               Left            =   90
               TabIndex        =   39
               Top             =   3585
               Width           =   1215
            End
            Begin VB.Label lblData 
               Alignment       =   1  'Right Justify
               Caption         =   "Coefficient ""D"":"
               Height          =   225
               Index           =   6
               Left            =   90
               TabIndex        =   37
               Top             =   3225
               Width           =   1215
            End
            Begin VB.Label lblData 
               Alignment       =   1  'Right Justify
               Caption         =   "Coefficient ""C"":"
               Height          =   225
               Index           =   5
               Left            =   90
               TabIndex        =   35
               Top             =   2865
               Width           =   1215
            End
            Begin VB.Label lblData 
               Alignment       =   1  'Right Justify
               Caption         =   "Coefficient ""B"":"
               Height          =   225
               Index           =   4
               Left            =   90
               TabIndex        =   33
               Top             =   2505
               Width           =   1215
            End
            Begin VB.Label lblData 
               Alignment       =   1  'Right Justify
               Caption         =   "Coefficient ""A"":"
               Height          =   225
               Index           =   3
               Left            =   90
               TabIndex        =   31
               Top             =   2145
               Width           =   1215
            End
            Begin VB.Label lblData 
               Alignment       =   1  'Right Justify
               Caption         =   "Maximum T:"
               Height          =   225
               Index           =   2
               Left            =   90
               TabIndex        =   29
               Top             =   1785
               Width           =   1215
            End
            Begin VB.Label lblData 
               Alignment       =   1  'Right Justify
               Caption         =   "Minimum T:"
               Height          =   225
               Index           =   1
               Left            =   90
               TabIndex        =   27
               Top             =   1425
               Width           =   1215
            End
         End
      End
      Begin Threed.SSPanel sspTechRef 
         Height          =   2085
         Left            =   2970
         TabIndex        =   8
         Top             =   2430
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   3678
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame ssfTechRef 
            Height          =   1935
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   3413
            _StockProps     =   14
            Caption         =   "Error (if any) and reference:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox txtError 
               BackColor       =   &H8000000F&
               Height          =   735
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               TabStop         =   0   'False
               Text            =   "Techniques.frx":0000
               Top             =   300
               Width           =   1755
            End
            Begin VB.TextBox txtReference 
               BackColor       =   &H8000000F&
               Height          =   735
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               TabStop         =   0   'False
               Text            =   "Techniques.frx":000B
               Top             =   1080
               Width           =   1755
            End
         End
      End
   End
   Begin Threed.SSPanel sspPropertyNote 
      Height          =   2595
      Left            =   5940
      TabIndex        =   17
      Top             =   840
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   4577
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtPropertyNote 
         Height          =   555
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "Techniques.frx":001A
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
   End
   Begin Threed.SSPanel sspDipprData 
      Height          =   5325
      Left            =   2250
      TabIndex        =   16
      Top             =   390
      Width           =   8565
      _Version        =   65536
      _ExtentX        =   15108
      _ExtentY        =   9393
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSFrame ssfDipprData 
         Height          =   5175
         Left            =   60
         TabIndex        =   48
         Top             =   60
         Width           =   7065
         _Version        =   65536
         _ExtentX        =   12462
         _ExtentY        =   9128
         _StockProps     =   14
         Caption         =   "{DIPPR801/DIPPR911 Data:}"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame ssfCitation 
            Height          =   1245
            Left            =   90
            TabIndex        =   63
            Top             =   3750
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   2196
            _StockProps     =   14
            Caption         =   "Reference:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox txtCitation 
               BackColor       =   &H8000000F&
               Height          =   855
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   64
               Text            =   "Techniques.frx":002C
               Top             =   270
               Width           =   3165
            End
         End
         Begin Threed.SSFrame ssfComment 
            Height          =   1245
            Left            =   90
            TabIndex        =   61
            Top             =   2520
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   2196
            _StockProps     =   14
            Caption         =   "Comment:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox txtComment 
               BackColor       =   &H8000000F&
               Height          =   855
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   62
               Text            =   "Techniques.frx":003A
               Top             =   270
               Width           =   3165
            End
         End
         Begin VB.TextBox txtDataStr 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   4
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "txtDataStr()"
            Top             =   2070
            Width           =   2000
         End
         Begin VB.TextBox txtDataStr 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   3
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "txtDataStr()"
            Top             =   1710
            Width           =   2000
         End
         Begin VB.TextBox txtDataStr 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   2
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "txtDataStr()"
            Top             =   1350
            Width           =   2000
         End
         Begin VB.TextBox txtDataStr 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   1
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "txtDataStr()"
            Top             =   990
            Width           =   2000
         End
         Begin VB.TextBox txtDataStr 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   0
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "txtDataStr()"
            Top             =   630
            Width           =   2000
         End
         Begin VB.Label lblNotAvailable 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "{ The {DIPPR801/911} data is not available. Do not move this textbox! }"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   90
            TabIndex        =   60
            Top             =   270
            Width           =   5535
         End
         Begin VB.Label lblToDoItem 
            Alignment       =   2  'Center
            Caption         =   "Kline/Rogers need to elaborate on the window as to the meaning of the rating codes and reliability codes"
            ForeColor       =   &H000000FF&
            Height          =   915
            Left            =   4260
            TabIndex        =   59
            Top             =   810
            Width           =   2325
         End
         Begin VB.Label lblDataStr 
            Alignment       =   1  'Right Justify
            Caption         =   "Pressure:"
            Height          =   225
            Index           =   4
            Left            =   60
            TabIndex        =   58
            Top             =   2115
            Width           =   1845
         End
         Begin VB.Label lblDataStr 
            Alignment       =   1  'Right Justify
            Caption         =   "Method Description:"
            Height          =   225
            Index           =   3
            Left            =   60
            TabIndex        =   56
            Top             =   1755
            Width           =   1845
         End
         Begin VB.Label lblDataStr 
            Alignment       =   1  'Right Justify
            Caption         =   "Reliability Code:"
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   54
            Top             =   1395
            Width           =   1845
         End
         Begin VB.Label lblDataStr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rating Code:"
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   52
            Top             =   1035
            Width           =   1845
         End
         Begin VB.Label lblDataStr 
            Alignment       =   1  'Right Justify
            Caption         =   "CAS:"
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   50
            Top             =   675
            Width           =   1845
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   6075
      Left            =   11010
      ScaleHeight     =   6015
      ScaleWidth      =   915
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      Begin VB.PictureBox picLeaf 
         Height          =   345
         Left            =   1560
         Picture         =   "Techniques.frx":0047
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   24
         Top             =   6210
         Width           =   345
      End
      Begin VB.PictureBox picOpen 
         Height          =   345
         Left            =   1170
         Picture         =   "Techniques.frx":0131
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   23
         Top             =   6210
         Width           =   345
      End
      Begin VB.PictureBox picClosed 
         Height          =   345
         Left            =   780
         Picture         =   "Techniques.frx":021B
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   22
         Top             =   6210
         Width           =   345
      End
      Begin VB.PictureBox picValidated 
         Height          =   345
         Left            =   60
         Picture         =   "Techniques.frx":0305
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   21
         Top             =   750
         Width           =   345
      End
      Begin VB.PictureBox picUnvalidated 
         Height          =   345
         Left            =   450
         Picture         =   "Techniques.frx":03EF
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   20
         Top             =   750
         Width           =   345
      End
      Begin MSComctlLib.ImageList ilist_Valid 
         Left            =   30
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ilist_Tree 
         Left            =   150
         Top             =   6210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin Threed.SSPanel sspDataSheets 
      Height          =   2295
      Left            =   30
      TabIndex        =   9
      Top             =   3390
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   4048
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstPropSheets 
         BackColor       =   &H8000000F&
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1485
      End
   End
   Begin Threed.SSPanel sspButtons 
      Height          =   2745
      Left            =   30
      TabIndex        =   10
      Top             =   120
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   4842
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdButton 
         Caption         =   "Re&move override"
         Height          =   345
         Index           =   3
         Left            =   120
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "Remove technique hierarchy override for this property"
         Top             =   2250
         Width           =   1485
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Specify &override"
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "Specify technique hierarchy override for this property"
         Top             =   1890
         Width           =   1485
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Remove user input"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Remove user input value for this property"
         Top             =   1410
         Width           =   1485
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Specify user input"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Specify user input value for this property"
         Top             =   1050
         Width           =   1485
      End
      Begin VB.CommandButton cmdCancelOK 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Cancel changes (if any) and return to main window"
         Top             =   150
         Width           =   1485
      End
      Begin VB.CommandButton cmdCancelOK 
         Caption         =   "&Accept"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Accept changes (if any) and return to main window"
         Top             =   510
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmTechniques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NowProj_UponEntry As Project_Type
Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim Activated_Yet As Boolean
Dim inout_First_Display As Boolean

Dim sspSep_Setting As Double    'RANGE: 0 to 1.

Public HALT_Controls As Boolean
Public Window_Property_Code As Long
Public Window_idx_Chemical As Integer
Public Window_idx_PropertyData As Integer
Public Window_idx_PropertySheetOrder_FIRST As Integer
Public Window_idx_PropertyOrder_FIRST As Integer




Const frmTechniques_decl_end = True


Function frmTechniques_Go( _
    in_idx_Chem As Integer, _
    in_Property_Code As Long, _
    out_HitCancel As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
  Window_Property_Code = in_Property_Code
  Window_idx_Chemical = in_idx_Chem
  '
  ' LOOK UP INDEX OF THIS PROPERTY FOR THIS CHEMICAL.
  '
  Window_idx_PropertyData = PropertyData_GetIndex( _
      in_idx_Chem, _
      in_Property_Code)
  If (Window_idx_PropertyData = -1) Then
    Call Show_Error("Unable to find property code of " & _
        Trim$(Str$(in_Property_Code)) & " for this chemical; " & _
        "cancelling technique view.")
    GoTo exit_err_ThisFunc
  End If
  '
  ' LOOK UP INDEXES OF THIS PROPERTY IN HIERARCHY.
  '
Dim out_idx_PropertySheetOrder() As Integer
Dim out_idx_PropertyOrder() As Integer
Dim out_Size As Integer
Dim DoCancel As Boolean
  DoCancel = False
  If (False = PropertyOrder_Property_Code_GetIndexes( _
      in_Property_Code, _
      out_idx_PropertySheetOrder(), _
      out_idx_PropertyOrder(), _
      out_Size)) Then
    DoCancel = True
  Else
    If (out_Size < 1) Then DoCancel = True
  End If
  If (DoCancel = True) Then
    Call Show_Error("Unable to find property code of " & _
        Trim$(Str$(in_Property_Code)) & " in the " & _
        "hierarchy; cancelling technique view.")
    GoTo exit_err_ThisFunc
  End If
  Window_idx_PropertySheetOrder_FIRST = _
      out_idx_PropertySheetOrder(1)
  Window_idx_PropertyOrder_FIRST = _
      out_idx_PropertyOrder(1)
Call debug_output("frmTechniques_Go: " & _
    "Window_idx_PropertySheetOrder_FIRST = " & Trim$(Str$(Window_idx_PropertySheetOrder_FIRST)) & ", " & _
    "Window_idx_PropertyOrder_FIRST = " & Trim$(Str$(Window_idx_PropertyOrder_FIRST)) & ".")
  '
  ' STORE ALL SETTINGS UPON ENTRY.
  '
  NowProj_UponEntry = NowProj
  '
  ' DISPLAY THE WINDOW.
  '
Call debug_output("frmTechniques_Go - pre-`Show`-command")
  frmTechniques.Show 1
Call debug_output("frmTechniques_Go - post-`Show`-command")
  If (USER_HIT_CANCEL = True) Then
    ' USER CANCELLED; REVERT TO ORIGINAL SETTINGS.
    NowProj = NowProj_UponEntry
    out_HitCancel = True
  Else
    out_HitCancel = False
  End If
Call debug_output("frmTechniques_Go - post-post-`Show`-command")
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
  frmTechniques_Go = True
  Exit Function
exit_err_ThisFunc:
  frmTechniques_Go = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_Go")
  GoTo exit_err_ThisFunc
End Function


Sub Populate_frmTechniques_Units()
Dim Frm As Form
Set Frm = frmTechniques
  Call unitsys_register(Frm, lblData(1), txtData(1), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(Frm, lblData(2), txtData(2), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(Frm, lblData(3), txtData(3), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(Frm, lblData(4), txtData(4), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(Frm, lblData(5), txtData(5), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(Frm, lblData(6), txtData(6), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(Frm, lblData(7), txtData(7), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Function frmTechniques_PopulateFirstTime_SeveralControls() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmTechniques
Dim Ctl_IL As Control
Dim Ctl_LV As Control
Dim ImgI As ListImage
Dim ItmX As ListItem
  'Frm.lstUser.Clear
  'Frm.lstPropSheets.Clear
  '
  ' SET UP Frm.ilist_Valid.
  '
  Set Ctl_IL = Frm.ilist_Valid
  Ctl_IL.ListImages.Clear
  Set ImgI = Ctl_IL.ListImages.Add _
      (, "validated", Frm.picValidated.Picture)
  Set ImgI = Ctl_IL.ListImages.Add _
      (, "unvalidated", Frm.picUnvalidated.Picture)
  '
  ' SET UP Frm.lvMain.
  '
  Set Ctl_LV = Frm.lvMain
  Ctl_LV.View = lvwReport
  Ctl_LV.Icons = Ctl_IL
  Ctl_LV.SmallIcons = Ctl_IL
  Ctl_LV.ColumnHeaders.Clear
  Ctl_LV.ColumnHeaders.Add , , "Valid", 600, lvwColumnLeft
  Ctl_LV.ColumnHeaders.Add , , "Used", 600, lvwColumnLeft
  Ctl_LV.ColumnHeaders.Add , , "Type", 600, lvwColumnLeft
  Ctl_LV.ColumnHeaders.Add , , "Technique", 2500, lvwColumnRight
  Ctl_LV.ColumnHeaders.Add , , "Value", 1200, lvwColumnRight
  Ctl_LV.ColumnHeaders.Add , , "Units", 1000, lvwColumnLeft
  '
  ' ------------ TEMPORARY TEST DATA FOLLOWS: ------------
  '
  Set ItmX = Ctl_LV.ListItems.Add(, "x1", " ")
  ItmX.SubItems(1) = ""
  ItmX.SubItems(2) = "User"
  ItmX.SubItems(3) = "User Data"
  ItmX.SubItems(4) = ""
  ItmX.SubItems(5) = "g/gmol"
  ItmX.Icon = 2: ItmX.SmallIcon = 2
  Set ItmX = Ctl_LV.ListItems.Add(, "x2", " ")
  ItmX.SubItems(1) = ""
  ItmX.SubItems(2) = "Data"
  ItmX.SubItems(3) = "DIPPR801"
  ItmX.SubItems(4) = "60.024"
  ItmX.SubItems(5) = "g/gmol"
  ItmX.Icon = 1: ItmX.SmallIcon = 1
  Set ItmX = Ctl_LV.ListItems.Add(, "x3", " ")
  ItmX.SubItems(1) = ""
  ItmX.SubItems(2) = "Est"
  ItmX.SubItems(3) = "UNIFAC"
  ItmX.SubItems(4) = "60.025"
  ItmX.SubItems(5) = "g/gmol"
  ItmX.Icon = 1: ItmX.SmallIcon = 1
  '
  ' ------------ TEMPORARY TEST DATA ENDS. ------------
  '
  
  
'Sub populate_lvThis()
'Dim ImgI As ListImage
'Dim ItmX As ListItem
'Dim Ctl_LV As Control
'Dim Ctl_IL As Control
'  'SET UP THE lvThis CONTROL (IMAGES, STYLES, ETC).
'  Set Ctl_IL = frmMain.ilist_Valid
'  Ctl_IL.ListImages.Clear
'  Set ImgI = Ctl_IL.ListImages.Add _
'      (, "validated", frmMain.picValidated.Picture)
'  Set ImgI = Ctl_IL.ListImages.Add _
'      (, "unvalidated", frmMain.picUnvalidated.Picture)
'  Set Ctl_LV = lvThis
'  Ctl_LV.View = lvwReport
'  Ctl_LV.Icons = Ctl_IL
'  Ctl_LV.SmallIcons = Ctl_IL
'  'Ctl_LV.ColumnHeaders.Add , , "x2", 1000, lvwColumnLeft
'  'Ctl_LV.ColumnHeaders.Add , , "x3", 1000, lvwColumnLeft
'''''  '
'''''  ' PROTOTYPE DATA.
'''''  '
'''''  Set ItmX = Ctl_LV.ListItems.Add(, "x1", "Condenser")
'''''  ItmX.SubItems(1) = "($62,853)"
'''''  ItmX.SubItems(2) = "Heat Transfer Operations : Shell and Tube Exchanger : Fixed Tubesheet : Carbon Steel, 150 psig"
'''''  ItmX.SubItems(3) = "Validated(LE)"
'''''  ItmX.Icon = 1: ItmX.SmallIcon = 1
'''''  Set ItmX = Ctl_LV.ListItems.Add(, "x2", "DC-100")
'''''  ItmX.SubItems(1) = "($517,501)"
'''''  ItmX.SubItems(2) = "Vessels : Column : Tray : Carbon Steel, 150 psig"
'''''  ItmX.SubItems(3) = "Validated(LE)"
'''''  ItmX.Icon = 1: ItmX.SmallIcon = 1
'''''  Set ItmX = Ctl_LV.ListItems.Add(, "x3", "Preheater")
'''''  ItmX.SubItems(1) = "($242,201)"
'''''  ItmX.SubItems(2) = "Heat Transfer Operations : Shell and Tube Exchanger : Fixed Tubesheet : Carbon Steel, 150 psig"
'''''  ItmX.SubItems(3) = "Unvalidated(LE)"
'''''  ItmX.Icon = 2: ItmX.SmallIcon = 2
'''''  Set ItmX = Ctl_LV.ListItems.Add(, "x4", "Reboiler")
'''''  ItmX.SubItems(1) = "($361,671)"
'''''  ItmX.SubItems(2) = "Heat Transfer Operations : Shell and Tube Exchanger : Kettle Reboiler : Carbon Steel, 150 psig"
'''''  ItmX.SubItems(3) = "Unvalidated(LE)"
'''''  ItmX.Icon = 2: ItmX.SmallIcon = 2
'End Sub
  
  
  
  
  



exit_normally_ThisFunc:
  frmTechniques_PopulateFirstTime_SeveralControls = True
  Exit Function
exit_err_ThisFunc:
  frmTechniques_PopulateFirstTime_SeveralControls = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_PopulateFirstTime_SeveralControls")
  Resume exit_err_ThisFunc
End Function


Function frmTechniques_Resize( _
    ) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Width_of_This As Double
Dim Height_of_This As Double
Dim XX As Double
Const USE_MARGIN_sspSep = 50
Const USE_MARGIN = 50
  '
  '////////// START OF MAIN RESIZING CODE. ///////////////////////////////////////
  '
  '
  ' RESIZE sspButtons AND ALL CONTAINED CONTROLS.
  '
  sspButtons.Move _
      60, _
      60, _
      sspButtons.Width, _
      sspButtons.Height
  '
  ' RESIZE sspDataSheets AND ALL CONTAINED CONTROLS.
  '
  sspDataSheets.Move _
      60, _
      sspButtons.Top + sspButtons.Height + 100, _
      sspDataSheets.Width, _
      sspDataSheets.Height
  '
  ' RESIZE sspTech AND ALL CONTAINED CONTROLS.
  '
  XX = Me.ScaleWidth - 60 - (sspButtons.Left + sspButtons.Width)
  Width_of_This = IIf(XX > 10, XX, 10)
  XX = Me.ScaleHeight - 60 - (sspButtons.Top)
  Height_of_This = IIf(XX > 10, XX, 10)
  sspTech.Move _
      sspButtons.Left + sspButtons.Width, _
      sspButtons.Top, _
      Width_of_This, _
      Height_of_This
    '
    ' RESIZE sspSep AND ALL CONTAINED CONTROLS.
    '
    sspSep.Move _
        -1000, _
        CDbl(sspTech.Height) * sspSep_Setting, _
        sspTech.Width + 1000 + 1000, _
        sspSep.Height
    '
    ' RESIZE ssfTechGrid AND ALL CONTAINED CONTROLS.
    '
    XX = sspTech.Width - 2 * USE_MARGIN
    Width_of_This = IIf(XX > 10, XX, 10)
    XX = sspSep.Top - USE_MARGIN_sspSep - USE_MARGIN
    Height_of_This = IIf(XX > 10, XX, 10)
    ssfTechGrid.Move _
        USE_MARGIN, _
        USE_MARGIN, _
        Width_of_This, _
        Height_of_This
      '
      ' RESIZE lvMain AND ALL CONTAINED CONTROLS.
      '
      XX = ssfTechGrid.Width - 60 - 60
      Width_of_This = IIf(XX > 100#, XX, 100#)
      XX = ssfTechGrid.Height - 60 - 60
      Height_of_This = IIf(XX > 100#, XX, 100#)
      lvMain.Move _
          60, _
          60, _
          Width_of_This, _
          Height_of_This
    '
    ' RESIZE sspTechTemp AND ALL CONTAINED CONTROLS.
    '
    XX = sspTech.Height - USE_MARGIN - _
        (sspSep.Top + sspSep.Height + USE_MARGIN_sspSep)
    Height_of_This = _
        IIf(XX > ssfTechTemp.Height, XX, ssfTechTemp.Height)
    sspTechTemp.Move _
        USE_MARGIN, _
        sspSep.Top + sspSep.Height + USE_MARGIN_sspSep, _
        ssfTechTemp.Width, _
        Height_of_This
      '
      ' RESIZE ssfTechTemp AND ALL CONTAINED CONTROLS.
      '
      ssfTechTemp.Move _
          0, _
          0, _
          ssfTechTemp.Width, _
          ssfTechTemp.Height
    '
    ' RESIZE sspTechMisc AND ALL CONTAINED CONTROLS.
    '
    XX = sspTech.Width - USE_MARGIN - _
        (sspTechTemp.Left + sspTechTemp.Width)
    Width_of_This = IIf(XX > 100#, XX, 100#)
    sspTechMisc.Move _
        sspTechTemp.Left + sspTechTemp.Width, _
        sspSep.Top + sspSep.Height + USE_MARGIN_sspSep, _
        Width_of_This, _
        sspTechMisc.Height
      '
      ' RESIZE ssfTechMisc AND ALL CONTAINED CONTROLS.
      '
      ssfTechMisc.Move _
          0, _
          0, _
          sspTechMisc.Width, _
          sspTechMisc.Height
    '
    ' RESIZE sspTechRef AND ALL CONTAINED CONTROLS.
    '
    Width_of_This = sspTechMisc.Width
    XX = sspTech.Height - USE_MARGIN - _
        (sspTechMisc.Top + sspTechMisc.Height)
    Height_of_This = IIf(XX > 100#, XX, 100#)
    sspTechRef.Move _
        sspTechTemp.Left + sspTechTemp.Width, _
        sspTechMisc.Top + sspTechMisc.Height, _
        Width_of_This, _
        Height_of_This
      '
      ' RESIZE ssfTechRef AND ALL CONTAINED CONTROLS.
      '
      ssfTechRef.Move _
          0, _
          0, _
          sspTechRef.Width, _
          sspTechRef.Height
        '
        ' RESIZE txtError AND ALL CONTAINED CONTROLS.
        '
        XX = ssfTechRef.Width - 120 - 120
        Width_of_This = IIf(XX > 10, XX, 10)
        Height_of_This = txtError.Height
        txtError.Move _
            120, _
            300, _
            Width_of_This, _
            Height_of_This
        '
        ' RESIZE txtReference AND ALL CONTAINED CONTROLS.
        '
        Width_of_This = txtError.Width
        XX = ssfTechRef.Height - 60 - _
            (txtError.Top + txtError.Height + 120)
        Height_of_This = IIf(XX > 10, XX, 10)
        txtReference.Move _
            120, _
            txtError.Top + txtError.Height + 120, _
            Width_of_This, _
            Height_of_This
  '
  ' RESIZE sspDipprData AND ALL CONTAINED CONTROLS.
  '
  sspDipprData.Move _
      sspTech.Left, _
      sspTech.Top, _
      sspTech.Width, _
      sspTech.Height
    '
    ' RESIZE ssfDipprData AND ALL CONTAINED CONTROLS.
    '
    XX = sspDipprData.Width - 2 * USE_MARGIN
    Width_of_This = IIf(XX > 10, XX, 10)
    XX = sspDipprData.Height - 2 * USE_MARGIN
    Height_of_This = IIf(XX > 10, XX, 10)
    ssfDipprData.Move _
        USE_MARGIN, _
        USE_MARGIN, _
        Width_of_This, _
        Height_of_This
      '
      ' RESIZE lblNotAvailable AND ALL CONTAINED CONTROLS.
      '
      XX = ssfDipprData.Width - 2 * USE_MARGIN
      Width_of_This = IIf(XX > 10, XX, 10)
      lblNotAvailable.Move _
          USE_MARGIN, _
          lblNotAvailable.Top, _
          Width_of_This, _
          lblNotAvailable.Height
      '
      ' RESIZE ssfComment AND ALL CONTAINED CONTROLS.
      '
      XX = ssfDipprData.Width - 2 * USE_MARGIN
      Width_of_This = IIf(XX > 10, XX, 10)
      ssfComment.Move _
          USE_MARGIN, _
          ssfComment.Top, _
          Width_of_This, _
          ssfComment.Height
        '
        ' RESIZE txtComment AND ALL CONTAINED CONTROLS.
        '
        XX = ssfComment.Width - 2 * 120
        Width_of_This = IIf(XX > 10, XX, 10)
        XX = ssfComment.Height - 270 - 2 * USE_MARGIN
        Height_of_This = IIf(XX > 10, XX, 10)
        txtComment.Move _
            120, _
            270, _
            Width_of_This, _
            Height_of_This
      '
      ' RESIZE ssfCitation AND ALL CONTAINED CONTROLS.
      '
      XX = ssfDipprData.Width - 2 * USE_MARGIN
      Width_of_This = IIf(XX > 10, XX, 10)
      XX = ssfDipprData.Height - USE_MARGIN - _
          (ssfComment.Top + ssfComment.Height + USE_MARGIN)
      Height_of_This = IIf(XX > 10, XX, 10)
      ssfCitation.Move _
          USE_MARGIN, _
          ssfComment.Top + ssfComment.Height + USE_MARGIN, _
          Width_of_This, _
          Height_of_This
        '
        ' RESIZE txtCitation AND ALL CONTAINED CONTROLS.
        '
        XX = ssfCitation.Width - 2 * 120
        Width_of_This = IIf(XX > 10, XX, 10)
        XX = ssfCitation.Height - 270 - 2 * USE_MARGIN
        Height_of_This = IIf(XX > 10, XX, 10)
        txtCitation.Move _
            120, _
            270, _
            Width_of_This, _
            Height_of_This
      
      
      
      
      
      
      
      
  '
  ' RESIZE sspPropertyNote AND ALL CONTAINED CONTROLS.
  '
  sspPropertyNote.Move _
      sspTech.Left, _
      sspTech.Top, _
      sspTech.Width, _
      sspTech.Height









  '
  '////////// END OF MAIN RESIZING CODE. ///////////////////////////////////////
  '
exit_normally_ThisFunc:
  frmTechniques_Resize = True
  Exit Function
exit_err_ThisFunc:
  frmTechniques_Resize = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_Resize")
  GoTo exit_err_ThisFunc
End Function


Private Sub cmdButton_Click(Index As Integer)
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmTechniques
Dim in_UnitType As String
Dim in_UnitBase As String
Dim inout_UnitDisplayed As String
Dim inout_ValueInBaseUnits As Double
Dim out_idx_PropertyData As Integer
Dim out_idx_TechniqueData As Integer
Dim out_HitCancel As Boolean
Dim Old_Key As String
Dim out_idx_Technique_Code As Integer
Dim This_Technique_Code As Long
  Select Case Index
    '
    '//////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////
    '
    ' SPECIFY USER INPUT.
    Case 0:
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(Window_idx_PropertyData)
        in_UnitType = .UnitType
        in_UnitBase = .UnitBase
        inout_UnitDisplayed = .UnitDisplayed
      End With
      If (False = TechniqueData_GetIndex( _
          Window_idx_Chemical, _
          Window_Property_Code, _
          TECHCODE_ANY_000u_USER_INPUT, _
          out_idx_PropertyData, _
          out_idx_TechniqueData)) Then
        GoTo exit_err_ThisFunc
      End If
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(out_idx_PropertyData). _
          TechniqueData(out_idx_TechniqueData)
        inout_ValueInBaseUnits = .value
      End With
      If (False = frmUnitsAndOrValue.frmUnitsAndOrValue_GoUnitsAndValue( _
          in_UnitType, _
          in_UnitBase, _
          inout_UnitDisplayed, _
          inout_ValueInBaseUnits, _
          out_HitCancel)) Then
        GoTo exit_err_ThisFunc
      End If
      If (out_HitCancel = True) Then GoTo exit_normally_ThisFunc
      '
      ' UPDATE VALUE AND/OR UNITS.
      '
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(out_idx_PropertyData). _
          TechniqueData(out_idx_TechniqueData)
        .value = inout_ValueInBaseUnits
        .IsAvail = True
        .Error_Code = ""
      End With
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(Window_idx_PropertyData)
        .UnitDisplayed = inout_UnitDisplayed
      End With
      If (False = Recalculate_OneProperty( _
          Window_idx_Chemical, _
          out_idx_PropertyData)) Then
        GoTo exit_err_ThisFunc
      End If
      '
      ' REFRESH DISPLAY.
      '
      Call frmTechniques_Populate_CurrentDataTab(False)
    '
    '//////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////
    '
    ' REMOVE USER INPUT.
    Case 1:
      '
      ' IS USER INPUT SPECIFIED CURRENTLY?
      '
      If (False = TechniqueData_GetIndex( _
          Window_idx_Chemical, _
          Window_Property_Code, _
          TECHCODE_ANY_000u_USER_INPUT, _
          out_idx_PropertyData, _
          out_idx_TechniqueData)) Then
        GoTo exit_err_ThisFunc
      End If
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(out_idx_PropertyData). _
          TechniqueData(out_idx_TechniqueData)
        If (.IsAvail = False) Then
          Call Show_Error("There is currently no user input specified.")
          GoTo exit_err_ThisFunc
        End If
      End With
      If (vbNo = MsgBox("Are you sure you want to remove the currently " & _
          "specified user input?  This step cannot be undone.", _
          vbQuestion + vbYesNo, _
          "Remove user input: Are you sure?")) Then
        GoTo exit_normally_ThisFunc
      End If
      '
      ' UPDATE VALUE AND/OR UNITS.
      '
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(out_idx_PropertyData). _
          TechniqueData(out_idx_TechniqueData)
        .value = 0#
        .IsAvail = False
        .Error_Code = TECH_ERRORCODE_NEVER_INITED
      End With
      If (False = Recalculate_OneProperty( _
          Window_idx_Chemical, _
          out_idx_PropertyData)) Then
        GoTo exit_err_ThisFunc
      End If
      '
      ' REFRESH DISPLAY.
      '
      Call frmTechniques_Populate_CurrentDataTab(False)
    '
    '//////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////
    '
    ' SPECIFY OVERRIDE.
    Case 2:
      Old_Key = "(n/a)"
      On Error Resume Next
      Old_Key = Frm.lvMain.SelectedItem.Key
      On Error GoTo err_ThisFunc
      'OnError GoTo 0
      If (Old_Key = "(n/a)") Then
        GoTo exit_err_ThisFunc
      End If
      Call frmTechniques_lvMain_Extract_Key_Info( _
          Old_Key, _
          out_idx_Technique_Code, _
          out_idx_TechniqueData)
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(Window_idx_PropertyData). _
          TechniqueData(out_idx_TechniqueData)
        This_Technique_Code = .Technique_Code
      End With
      '
      ' UPDATE TECHNIQUE USAGE.
      '
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(Window_idx_PropertyData)
        .Override_Technique_Code = This_Technique_Code
      End With
      If (False = Recalculate_OneProperty( _
          Window_idx_Chemical, _
          Window_idx_PropertyData)) Then
        GoTo exit_err_ThisFunc
      End If
      '
      ' REFRESH DISPLAY.
      '
      Call frmTechniques_Populate_CurrentDataTab(False)
    '
    '//////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////
    '
    ' REMOVE OVERRIDE.
    Case 3:
      '
      ' IS OVERRIDE SPECIFIED CURRENTLY?
      '
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(Window_idx_PropertyData)
        If (.Override_Technique_Code = -1) Then
          Call Show_Error("There is currently no technique " & _
              "hierarchy override specified.")
          GoTo exit_err_ThisFunc
        End If
      End With
      '
      ' UPDATE TECHNIQUE USAGE.
      '
      With NowProj.UserChemicals(Window_idx_Chemical). _
          PropertyData(Window_idx_PropertyData)
        .Override_Technique_Code = -1
      End With
      If (False = Recalculate_OneProperty( _
          Window_idx_Chemical, _
          Window_idx_PropertyData)) Then
        GoTo exit_err_ThisFunc
      End If
      '
      ' REFRESH DISPLAY.
      '
      Call frmTechniques_Populate_CurrentDataTab(False)
  End Select
exit_normally_ThisFunc:
  'xxxxx = True
  Exit Sub
exit_err_ThisFunc:
  'xxxxx = False
  Exit Sub
err_ThisFunc:
  Call Show_Trapped_Error("cmdButton_Click")
  Resume exit_err_ThisFunc
End Sub

Private Sub cmdCancelOK_Click(Index As Integer)
  Select Case Index
    Case 0:       'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:       'OK.
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub Form_Activate()
  If (Activated_Yet = False) Then
    Activated_Yet = True
    '
    ' FIRST REFRESHES.
    '
    Call frmTechniques_PopulateFirstTime_SeveralControls
    Call frmTechniques_Populate_lstPropSheets
    inout_First_Display = True
    Call frmTechniques_Populate_CurrentDataTab(inout_First_Display)
    Call frmTechniques_Populate_DataDetails
    Call frmTechniques_Refresh
  End If
End Sub
Private Sub Form_Load()
Dim out_Name As String
''''Call debug_output("frmTechniques - Form_Load")
  '
  ' MISC INITS.
  '
  Activated_Yet = False
  Me.Width = 9600
  Me.Height = 7200
  USER_HIT_CANCEL = False
  USER_HIT_OK = False
  Call CenterOnForm(Me, frmMain)
  sspSep_Setting = 0.35
  HALT_Controls = False
  Call Given_PropCode_Get_Name( _
      Window_Property_Code, _
      out_Name)
  Me.Caption = "Techniques for `" & out_Name & _
      "`, for chemical `" & _
      NowProj.UserChemicals(Window_idx_Chemical).Name & "`"
  '
  ' SET UP UNIT CONTROLS.
  '
  Call Populate_frmTechniques_Units
  '
  ' MISC CONTROL RESETTINGS.
  '
  '''''''sspTech.BevelWidth = 0
  ssfTechGrid.BevelWidth = 0
  sspTechMisc.BevelWidth = 0
  sspTechTemp.BevelWidth = 0
  sspTechRef.BevelWidth = 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (USER_HIT_CANCEL = False) And (USER_HIT_OK = False) Then
    Call Show_Error("To exit this window, please press either the " & _
        "`Cancel` button or the `Accept` button.")
    Cancel = True
    Exit Sub
  End If
End Sub
Private Sub Form_Resize()
  If (Me.WindowState <> vbMinimized) Then
    '
    ' WARNING: RESIZING WHILE MINIMIZED CAN CAUSE ERRORS!
    '
    Call frmTechniques_Resize
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
''''Call debug_output("frmTechniques - Form_Unload")
End Sub


Private Sub lstPropSheets_Click()
  Call frmTechniques_Populate_CurrentDataTab(inout_First_Display)
  On Error Resume Next
  lstPropSheets.SetFocus
End Sub



Private Sub ssfPropertyNote_Click()

End Sub


Private Sub lvMain_Click()
  If (HALT_Controls = True) Then Exit Sub
  'MsgBox Me.lvMain.SelectedItem.Key & " (lvMain_Click)"
  Call frmTechniques_Populate_DataDetails
End Sub
'Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
'  MsgBox Me.lvMain.SelectedItem.Key & " (lvMain_ItemClick)"
'End Sub
'Private Sub lvMain_KeyPress(KeyAscii As Integer)
'  MsgBox Me.lvMain.SelectedItem.Key & " (lvMain_KeyPress)"
'End Sub
Private Sub lvMain_KeyUp(KeyCode As Integer, Shift As Integer)
  If (HALT_Controls = True) Then Exit Sub
  'MsgBox Me.lvMain.SelectedItem.Key & " (lvMain_KeyUp)"
  Call frmTechniques_Populate_DataDetails
End Sub




