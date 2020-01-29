VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmeditgr 
   Caption         =   "Edit Groups"
   ClientHeight    =   5520
   ClientLeft      =   495
   ClientTop       =   1995
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdleft 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdright 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdaccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Frame frusergr 
      Caption         =   "Selected Groups"
      Height          =   4335
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   204
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   203
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   202
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   201
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   200
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   199
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   198
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   197
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   196
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   195
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   194
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   193
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   192
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   191
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   190
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblselindex 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   189
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   53
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   52
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   15
         Left            =   720
         TabIndex        =   51
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   14
         Left            =   720
         TabIndex        =   50
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   49
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   48
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   47
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   46
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   45
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   44
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   43
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   42
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   41
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   40
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   39
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   38
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblselno 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   36
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   13
         Left            =   720
         TabIndex        =   35
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   12
         Left            =   720
         TabIndex        =   34
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   11
         Left            =   720
         TabIndex        =   33
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   10
         Left            =   720
         TabIndex        =   32
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   31
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   23
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblsel 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   22
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frselgr 
      Caption         =   "Select Groups From"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin TabDlg.SSTab SSTab1 
         Height          =   4695
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8281
         _Version        =   393216
         Tabs            =   10
         TabsPerRow      =   5
         TabHeight       =   476
         TabCaption(0)   =   "1"
         TabPicture(0)   =   "Frmeditg.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1(5)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label1(6)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label1(7)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label1(8)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label1(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label1(14)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label1(13)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label1(12)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label1(11)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label1(10)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "2"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1(20)"
         Tab(1).Control(1)=   "Label1(21)"
         Tab(1).Control(2)=   "Label1(22)"
         Tab(1).Control(3)=   "Label1(23)"
         Tab(1).Control(4)=   "Label1(24)"
         Tab(1).Control(5)=   "Label1(25)"
         Tab(1).Control(6)=   "Label1(26)"
         Tab(1).Control(7)=   "Label1(27)"
         Tab(1).Control(8)=   "Label1(28)"
         Tab(1).Control(9)=   "Label1(29)"
         Tab(1).Control(10)=   "Label1(19)"
         Tab(1).Control(11)=   "Label1(18)"
         Tab(1).Control(12)=   "Label1(17)"
         Tab(1).Control(13)=   "Label1(16)"
         Tab(1).Control(14)=   "Label1(15)"
         Tab(1).ControlCount=   15
         TabCaption(2)   =   "3"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label1(40)"
         Tab(2).Control(1)=   "Label1(41)"
         Tab(2).Control(2)=   "Label1(42)"
         Tab(2).Control(3)=   "Label1(43)"
         Tab(2).Control(4)=   "Label1(44)"
         Tab(2).Control(5)=   "Label1(30)"
         Tab(2).Control(6)=   "Label1(31)"
         Tab(2).Control(7)=   "Label1(32)"
         Tab(2).Control(8)=   "Label1(33)"
         Tab(2).Control(9)=   "Label1(34)"
         Tab(2).Control(10)=   "Label1(35)"
         Tab(2).Control(11)=   "Label1(36)"
         Tab(2).Control(12)=   "Label1(37)"
         Tab(2).Control(13)=   "Label1(38)"
         Tab(2).Control(14)=   "Label1(39)"
         Tab(2).ControlCount=   15
         TabCaption(3)   =   "4"
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label1(50)"
         Tab(3).Control(1)=   "Label1(51)"
         Tab(3).Control(2)=   "Label1(52)"
         Tab(3).Control(3)=   "Label1(53)"
         Tab(3).Control(4)=   "Label1(54)"
         Tab(3).Control(5)=   "Label1(55)"
         Tab(3).Control(6)=   "Label1(56)"
         Tab(3).Control(7)=   "Label1(57)"
         Tab(3).Control(8)=   "Label1(58)"
         Tab(3).Control(9)=   "Label1(59)"
         Tab(3).Control(10)=   "Label1(45)"
         Tab(3).Control(11)=   "Label1(46)"
         Tab(3).Control(12)=   "Label1(47)"
         Tab(3).Control(13)=   "Label1(48)"
         Tab(3).Control(14)=   "Label1(49)"
         Tab(3).ControlCount=   15
         TabCaption(4)   =   "5"
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label1(70)"
         Tab(4).Control(1)=   "Label1(71)"
         Tab(4).Control(2)=   "Label1(72)"
         Tab(4).Control(3)=   "Label1(73)"
         Tab(4).Control(4)=   "Label1(74)"
         Tab(4).Control(5)=   "Label1(60)"
         Tab(4).Control(6)=   "Label1(61)"
         Tab(4).Control(7)=   "Label1(62)"
         Tab(4).Control(8)=   "Label1(63)"
         Tab(4).Control(9)=   "Label1(64)"
         Tab(4).Control(10)=   "Label1(65)"
         Tab(4).Control(11)=   "Label1(66)"
         Tab(4).Control(12)=   "Label1(67)"
         Tab(4).Control(13)=   "Label1(68)"
         Tab(4).Control(14)=   "Label1(69)"
         Tab(4).ControlCount=   15
         TabCaption(5)   =   "6"
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label1(80)"
         Tab(5).Control(1)=   "Label1(81)"
         Tab(5).Control(2)=   "Label1(82)"
         Tab(5).Control(3)=   "Label1(83)"
         Tab(5).Control(4)=   "Label1(84)"
         Tab(5).Control(5)=   "Label1(85)"
         Tab(5).Control(6)=   "Label1(86)"
         Tab(5).Control(7)=   "Label1(87)"
         Tab(5).Control(8)=   "Label1(88)"
         Tab(5).Control(9)=   "Label1(89)"
         Tab(5).Control(10)=   "Label1(75)"
         Tab(5).Control(11)=   "Label1(76)"
         Tab(5).Control(12)=   "Label1(77)"
         Tab(5).Control(13)=   "Label1(78)"
         Tab(5).Control(14)=   "Label1(79)"
         Tab(5).ControlCount=   15
         TabCaption(6)   =   "7"
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Label1(100)"
         Tab(6).Control(1)=   "Label1(101)"
         Tab(6).Control(2)=   "Label1(102)"
         Tab(6).Control(3)=   "Label1(103)"
         Tab(6).Control(4)=   "Label1(104)"
         Tab(6).Control(5)=   "Label1(90)"
         Tab(6).Control(6)=   "Label1(91)"
         Tab(6).Control(7)=   "Label1(92)"
         Tab(6).Control(8)=   "Label1(93)"
         Tab(6).Control(9)=   "Label1(94)"
         Tab(6).Control(10)=   "Label1(95)"
         Tab(6).Control(11)=   "Label1(96)"
         Tab(6).Control(12)=   "Label1(97)"
         Tab(6).Control(13)=   "Label1(98)"
         Tab(6).Control(14)=   "Label1(99)"
         Tab(6).ControlCount=   15
         TabCaption(7)   =   "8"
         TabPicture(7)   =   "Frmeditg.frx":001C
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Label1(109)"
         Tab(7).Control(1)=   "Label1(108)"
         Tab(7).Control(2)=   "Label1(107)"
         Tab(7).Control(3)=   "Label1(106)"
         Tab(7).Control(4)=   "Label1(105)"
         Tab(7).Control(5)=   "Label1(119)"
         Tab(7).Control(6)=   "Label1(118)"
         Tab(7).Control(7)=   "Label1(117)"
         Tab(7).Control(8)=   "Label1(116)"
         Tab(7).Control(9)=   "Label1(115)"
         Tab(7).Control(10)=   "Label1(114)"
         Tab(7).Control(11)=   "Label1(113)"
         Tab(7).Control(12)=   "Label1(112)"
         Tab(7).Control(13)=   "Label1(111)"
         Tab(7).Control(14)=   "Label1(110)"
         Tab(7).ControlCount=   15
         TabCaption(8)   =   "9"
         TabPicture(8)   =   "Frmeditg.frx":0038
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Label1(120)"
         Tab(8).Control(1)=   "Label1(121)"
         Tab(8).Control(2)=   "Label1(122)"
         Tab(8).Control(3)=   "Label1(123)"
         Tab(8).Control(4)=   "Label1(124)"
         Tab(8).Control(5)=   "Label1(125)"
         Tab(8).Control(6)=   "Label1(126)"
         Tab(8).Control(7)=   "Label1(127)"
         Tab(8).Control(8)=   "Label1(128)"
         Tab(8).Control(9)=   "Label1(129)"
         Tab(8).Control(10)=   "Label1(130)"
         Tab(8).Control(11)=   "Label1(131)"
         Tab(8).Control(12)=   "Label1(132)"
         Tab(8).Control(13)=   "Label1(133)"
         Tab(8).Control(14)=   "Label1(134)"
         Tab(8).ControlCount=   15
         TabCaption(9)   =   "10"
         TabPicture(9)   =   "Frmeditg.frx":0054
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Label1(149)"
         Tab(9).Control(1)=   "Label1(148)"
         Tab(9).Control(2)=   "Label1(147)"
         Tab(9).Control(3)=   "Label1(146)"
         Tab(9).Control(4)=   "Label1(145)"
         Tab(9).Control(5)=   "Label1(144)"
         Tab(9).Control(6)=   "Label1(143)"
         Tab(9).Control(7)=   "Label1(142)"
         Tab(9).Control(8)=   "Label1(141)"
         Tab(9).Control(9)=   "Label1(140)"
         Tab(9).Control(10)=   "Label1(139)"
         Tab(9).Control(11)=   "Label1(138)"
         Tab(9).Control(12)=   "Label1(137)"
         Tab(9).Control(13)=   "Label1(136)"
         Tab(9).Control(14)=   "Label1(135)"
         Tab(9).ControlCount=   15
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   149
            Left            =   -74760
            TabIndex        =   188
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   148
            Left            =   -74760
            TabIndex        =   187
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   147
            Left            =   -74760
            TabIndex        =   186
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   146
            Left            =   -74760
            TabIndex        =   185
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   145
            Left            =   -74760
            TabIndex        =   184
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   144
            Left            =   -74760
            TabIndex        =   183
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   143
            Left            =   -74760
            TabIndex        =   182
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   142
            Left            =   -74760
            TabIndex        =   181
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   141
            Left            =   -74760
            TabIndex        =   180
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   140
            Left            =   -74760
            TabIndex        =   179
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   139
            Left            =   -74760
            TabIndex        =   178
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   138
            Left            =   -74760
            TabIndex        =   177
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   137
            Left            =   -74760
            TabIndex        =   176
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   136
            Left            =   -74760
            TabIndex        =   175
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   135
            Left            =   -74760
            TabIndex        =   174
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   134
            Left            =   -74760
            TabIndex        =   173
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   133
            Left            =   -74760
            TabIndex        =   172
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   132
            Left            =   -74760
            TabIndex        =   171
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   131
            Left            =   -74760
            TabIndex        =   170
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   130
            Left            =   -74760
            TabIndex        =   169
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   129
            Left            =   -74760
            TabIndex        =   168
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   128
            Left            =   -74760
            TabIndex        =   167
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   127
            Left            =   -74760
            TabIndex        =   166
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   126
            Left            =   -74760
            TabIndex        =   165
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   125
            Left            =   -74760
            TabIndex        =   164
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   124
            Left            =   -74760
            TabIndex        =   163
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   123
            Left            =   -74760
            TabIndex        =   162
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   122
            Left            =   -74760
            TabIndex        =   161
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   121
            Left            =   -74760
            TabIndex        =   160
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   120
            Left            =   -74760
            TabIndex        =   159
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   110
            Left            =   -74760
            TabIndex        =   158
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   111
            Left            =   -74760
            TabIndex        =   157
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   112
            Left            =   -74760
            TabIndex        =   156
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   113
            Left            =   -74760
            TabIndex        =   155
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   114
            Left            =   -74760
            TabIndex        =   154
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   115
            Left            =   -74760
            TabIndex        =   153
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   116
            Left            =   -74760
            TabIndex        =   152
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   117
            Left            =   -74760
            TabIndex        =   151
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   118
            Left            =   -74760
            TabIndex        =   150
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   119
            Left            =   -74760
            TabIndex        =   149
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   105
            Left            =   -74760
            TabIndex        =   148
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   106
            Left            =   -74760
            TabIndex        =   147
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   107
            Left            =   -74760
            TabIndex        =   146
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   108
            Left            =   -74760
            TabIndex        =   145
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   109
            Left            =   -74760
            TabIndex        =   144
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   100
            Left            =   -74760
            TabIndex        =   143
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   101
            Left            =   -74760
            TabIndex        =   142
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   102
            Left            =   -74760
            TabIndex        =   141
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   103
            Left            =   -74760
            TabIndex        =   140
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   104
            Left            =   -74760
            TabIndex        =   139
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   90
            Left            =   -74760
            TabIndex        =   138
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   91
            Left            =   -74760
            TabIndex        =   137
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   92
            Left            =   -74760
            TabIndex        =   136
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   93
            Left            =   -74760
            TabIndex        =   135
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   94
            Left            =   -74760
            TabIndex        =   134
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   95
            Left            =   -74760
            TabIndex        =   133
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   96
            Left            =   -74760
            TabIndex        =   132
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   97
            Left            =   -74760
            TabIndex        =   131
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   98
            Left            =   -74760
            TabIndex        =   130
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   99
            Left            =   -74760
            TabIndex        =   129
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   80
            Left            =   -74760
            TabIndex        =   128
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   81
            Left            =   -74760
            TabIndex        =   127
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   82
            Left            =   -74760
            TabIndex        =   126
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   83
            Left            =   -74760
            TabIndex        =   125
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   84
            Left            =   -74760
            TabIndex        =   124
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   85
            Left            =   -74760
            TabIndex        =   123
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   86
            Left            =   -74760
            TabIndex        =   122
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   87
            Left            =   -74760
            TabIndex        =   121
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   88
            Left            =   -74760
            TabIndex        =   120
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   89
            Left            =   -74760
            TabIndex        =   119
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   75
            Left            =   -74760
            TabIndex        =   118
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   76
            Left            =   -74760
            TabIndex        =   117
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   77
            Left            =   -74760
            TabIndex        =   116
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   78
            Left            =   -74760
            TabIndex        =   115
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   79
            Left            =   -74760
            TabIndex        =   114
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   70
            Left            =   -74760
            TabIndex        =   113
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   71
            Left            =   -74760
            TabIndex        =   112
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   72
            Left            =   -74760
            TabIndex        =   111
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   73
            Left            =   -74760
            TabIndex        =   110
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   74
            Left            =   -74760
            TabIndex        =   109
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   60
            Left            =   -74760
            TabIndex        =   108
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   61
            Left            =   -74760
            TabIndex        =   107
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   62
            Left            =   -74760
            TabIndex        =   106
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   63
            Left            =   -74760
            TabIndex        =   105
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   64
            Left            =   -74760
            TabIndex        =   104
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   65
            Left            =   -74760
            TabIndex        =   103
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   66
            Left            =   -74760
            TabIndex        =   102
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   67
            Left            =   -74760
            TabIndex        =   101
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   68
            Left            =   -74760
            TabIndex        =   100
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   69
            Left            =   -74760
            TabIndex        =   99
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   50
            Left            =   -74760
            TabIndex        =   98
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   51
            Left            =   -74760
            TabIndex        =   97
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   52
            Left            =   -74760
            TabIndex        =   96
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   53
            Left            =   -74760
            TabIndex        =   95
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   54
            Left            =   -74760
            TabIndex        =   94
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   55
            Left            =   -74760
            TabIndex        =   93
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   56
            Left            =   -74760
            TabIndex        =   92
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   57
            Left            =   -74760
            TabIndex        =   91
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   58
            Left            =   -74760
            TabIndex        =   90
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   59
            Left            =   -74760
            TabIndex        =   89
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   45
            Left            =   -74760
            TabIndex        =   88
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   46
            Left            =   -74760
            TabIndex        =   87
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   47
            Left            =   -74760
            TabIndex        =   86
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   48
            Left            =   -74760
            TabIndex        =   85
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   49
            Left            =   -74760
            TabIndex        =   84
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   40
            Left            =   -74760
            TabIndex        =   83
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   41
            Left            =   -74760
            TabIndex        =   82
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   42
            Left            =   -74760
            TabIndex        =   81
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   43
            Left            =   -74760
            TabIndex        =   80
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   44
            Left            =   -74760
            TabIndex        =   79
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   30
            Left            =   -74760
            TabIndex        =   78
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   31
            Left            =   -74760
            TabIndex        =   77
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   32
            Left            =   -74760
            TabIndex        =   76
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   33
            Left            =   -74760
            TabIndex        =   75
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   34
            Left            =   -74760
            TabIndex        =   74
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   35
            Left            =   -74760
            TabIndex        =   73
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   36
            Left            =   -74760
            TabIndex        =   72
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   37
            Left            =   -74760
            TabIndex        =   71
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   38
            Left            =   -74760
            TabIndex        =   70
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   39
            Left            =   -74760
            TabIndex        =   69
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   20
            Left            =   -74760
            TabIndex        =   68
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   -74760
            TabIndex        =   67
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   -74760
            TabIndex        =   66
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   23
            Left            =   -74760
            TabIndex        =   65
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   24
            Left            =   -74760
            TabIndex        =   64
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   25
            Left            =   -74760
            TabIndex        =   63
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   26
            Left            =   -74760
            TabIndex        =   62
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   27
            Left            =   -74760
            TabIndex        =   61
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   28
            Left            =   -74760
            TabIndex        =   60
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   29
            Left            =   -74760
            TabIndex        =   59
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   58
            Top             =   3120
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   57
            Top             =   3360
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   56
            Top             =   3600
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   55
            Top             =   3840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   54
            Top             =   4080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   19
            Left            =   -74760
            TabIndex        =   21
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   18
            Left            =   -74760
            TabIndex        =   20
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   19
            Top             =   2880
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   17
            Left            =   -74760
            TabIndex        =   18
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   16
            Left            =   -74760
            TabIndex        =   17
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   15
            Left            =   -74760
            TabIndex        =   16
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   15
            Top             =   2640
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   14
            Top             =   2400
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   13
            Top             =   2160
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   12
            Top             =   1920
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   11
            Top             =   1680
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   10
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   9
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   2775
         End
      End
   End
End
Attribute VB_Name = "frmeditgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cur_selected_group As Integer
Dim cur_sel_selected_group As Integer
Private Sub cmdcalc_Click()

End Sub

Private Sub cmdAccept_Click()

    Dim i As Integer
    Dim count As Integer
    For i = 0 To MAX_GROUPS_PER_CHEM - 1
        If Trim(frmeditgr!lblselindex(i).Caption) <> "" Then
            cur_chem_groups(count) = CInt(frmeditgr!lblselindex(i).Caption)
            num_cur_chem_groups(count) = CInt(frmeditgr!lblselno(i).Caption)
            count = count + 1
        Else
            Exit For
        End If
    Next i
    For i = count To MAX_GROUPS_PER_CHEM - 1
        cur_chem_groups(i) = -1
        num_cur_chem_groups(i) = 0
    Next i
    Call update_groups
    ' reset the smiles and chemical name
    'frmstruct.tbxsmiles.Text = ""
    'frmstruct.lblchemname = "Use the 'rename' command to set the new chemical name or 'match' to find an existing chemical"
    'frmstruct.lblstatus = ""
    Unload Me
    
End Sub

Private Sub cmdcancel_Click()

    frmeditgr.Hide
    Unload Me
    
End Sub


Private Sub Command1_Click()

End Sub


Private Sub cmdleft_Click()

    Dim i As Integer
    Dim num_groups As Integer
    If cur_sel_selected_group > -1 And cur_sel_selected_group < MAX_GROUPS_PER_CHEM Then
        num_groups = CInt(lblselno(cur_sel_selected_group).Caption)
        If num_groups > 1 Then
            num_groups = num_groups - 1
            lblselno(cur_sel_selected_group).Caption = CStr(num_groups)
        Else
            lblselindex(cur_sel_selected_group).Caption = ""
            lblsel(cur_sel_selected_group).Caption = ""
            lblselno(cur_sel_selected_group).Caption = ""
            For i = cur_sel_selected_group To MAX_GROUPS_PER_CHEM - 2
                lblselindex(i).Caption = lblselindex(i + 1).Caption
                lblsel(i).Caption = lblsel(i + 1).Caption
                lblselno(i).Caption = lblselno(i + 1).Caption
            Next i
        End If
    End If
End Sub

Private Sub cmdright_Click()

    Dim i As Integer
    Dim added As Boolean
    added = False
    If cur_selected_group > -1 And cur_selected_group < MAX_GROUPS Then    'fixed
        'add this one to the selected list if there's room
        For i = 0 To MAX_GROUPS_PER_CHEM - 1
            If Trim(lblselindex(i).Caption) = CStr(cur_selected_group + 1) Or Trim(lblselindex(i).Caption) = CStr(cur_selected_group + 1) & "." Then    'fixed
                lblselno(i).Caption = CStr(CInt(lblselno(i).Caption) + 1)
                added = True
                If cur_sel_selected_group > -1 And cur_sel_selected_group < MAX_GROUPS_PER_CHEM Then
                    lblselindex(cur_sel_selected_group).Font.Bold = False
                    lblsel(cur_sel_selected_group).Font.Bold = False
                    lblselno(cur_sel_selected_group).Font.Bold = False
                End If
                cur_sel_selected_group = i
                lblselindex(i).Font.Bold = True
                lblsel(i).Font.Bold = True
                lblselno(i).Font.Bold = True
                
                Exit For
            ElseIf Trim(lblselindex(i).Caption) = "" Then
                lblselindex(i).Caption = CStr(cur_selected_group + 1)   ' fixed
                lblsel(i).Caption = group_smiles(cur_selected_group)    'fixed
                lblselno(i).Caption = "1"
                added = True
                If cur_sel_selected_group > -1 And cur_sel_selected_group < MAX_GROUPS_PER_CHEM Then
                    lblselindex(cur_sel_selected_group).Font.Bold = False
                    lblsel(cur_sel_selected_group).Font.Bold = False
                    lblselno(cur_sel_selected_group).Font.Bold = False
                End If
                cur_sel_selected_group = i
                lblselindex(i).Font.Bold = True
                lblsel(i).Font.Bold = True
                lblselno(i).Font.Bold = True
                Exit For
            End If
        Next i
        If added = False Then
            MsgBox ("exceeded maximum groups allowed, group not added")
        End If
            
    End If
    
End Sub

Private Sub Label1_Click(Index As Integer)

    If cur_selected_group > -1 And cur_selected_group < MAX_GROUPS Then 'fixed
        frmeditgr!Label1(cur_selected_group).Font.Bold = False  'fixed
    End If
    frmeditgr!Label1(Index).Font.Bold = True
    cur_selected_group = Index  'fixed
    
End Sub

Private Sub lblsel_Click(Index As Integer)

    If cur_sel_selected_group > -1 And cur_sel_selected_group < MAX_GROUPS_PER_CHEM Then
        lblsel(cur_sel_selected_group).Font.Bold = False
        lblselindex(cur_sel_selected_group).Font.Bold = False
        lblselno(cur_sel_selected_group).Font.Bold = False
    End If
    cur_sel_selected_group = Index
    lblsel(Index).Font.Bold = True
    lblselno(Index).Font.Bold = True
    lblselindex(Index).Font.Bold = True
End Sub

Private Sub lblselindex_Click(Index As Integer)

    If cur_sel_selected_group > -1 And cur_sel_selected_group < MAX_GROUPS_PER_CHEM Then
        lblsel(cur_sel_selected_group).Font.Bold = False
        lblselindex(cur_sel_selected_group).Font.Bold = False
        lblselno(cur_sel_selected_group).Font.Bold = False
    End If
    cur_sel_selected_group = Index
    lblsel(Index).Font.Bold = True
    lblselno(Index).Font.Bold = True
    lblselindex(Index).Font.Bold = True
    
End Sub


Private Sub lblselno_Click(Index As Integer)

    If cur_sel_selected_group > -1 And cur_sel_selected_group < MAX_GROUPS_PER_CHEM Then
        lblsel(cur_sel_selected_group).Font.Bold = False
        lblselindex(cur_sel_selected_group).Font.Bold = False
        lblselno(cur_sel_selected_group).Font.Bold = False
    End If
    cur_sel_selected_group = Index
    lblsel(Index).Font.Bold = True
    lblselno(Index).Font.Bold = True
    lblselindex(Index).Font.Bold = True
End Sub

