VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmMake 
   BorderStyle     =   0  'None
   Caption         =   "Make Game"
   ClientHeight    =   9405
   ClientLeft      =   105
   ClientTop       =   210
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   627
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Bunny Stop"
      Height          =   375
      Left            =   5280
      TabIndex        =   54
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer TimBunny 
      Enabled         =   0   'False
      Interval        =   1900
      Left            =   4800
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   720
      Top             =   7440
   End
   Begin VB.CommandButton CmdAllQuests 
      Caption         =   "Quest Summary"
      Height          =   375
      Left            =   5160
      TabIndex        =   51
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   5280
      TabIndex        =   49
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   6480
      TabIndex        =   40
      Top             =   5520
      Width           =   5055
      Begin VB.Label Label11 
         Caption         =   "1,2,3 - Transport"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "G - Grow Grass "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   57
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Z - Zoom Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   53
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Q - Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   52
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "F - Fill Swamp  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "D - Destroy Rock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   46
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "A - Axe Tree"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "C - Cut Weed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "B - Bomb Wall "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1680
         TabIndex        =   43
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "S - Save Game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "L - Light Darkness"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.PictureBox Holder 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   6480
      ScaleHeight     =   3015
      ScaleWidth      =   5055
      TabIndex        =   37
      Top             =   5640
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label Choice 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   ">>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   4290
         WordWrap        =   -1  'True
      End
      Begin VB.Label Pitanje 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   ":::"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   4320
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3840
      TabIndex        =   36
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4680
      TabIndex        =   35
      Text            =   "40"
      Top             =   8760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdTest 
      Caption         =   "Speech Tester"
      Height          =   495
      Left            =   5640
      TabIndex        =   34
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdSpeech 
      Caption         =   "Add Speech"
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton CmdQuest 
      Caption         =   "ViewQuestsStatus"
      Height          =   375
      Left            =   5160
      TabIndex        =   27
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton CmdMap 
      Caption         =   "MakeMap"
      Height          =   375
      Left            =   5160
      TabIndex        =   26
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   840
      TabIndex        =   23
      Top             =   8760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtSpeech 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1935
      Left            =   6720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "FrmMake.frx":0000
      Top             =   5640
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   10440
      TabIndex        =   20
      Text            =   "190"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   10320
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Inventory"
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Read Map   and Move to  X , Y"
      Height          =   495
      Left            =   5160
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Text            =   "7"
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Text            =   "9"
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Text            =   "New"
      Top             =   0
      Width           =   1335
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Width           =   5775
      Begin VB.Label etqNBomb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4080
         TabIndex        =   29
         Top             =   480
         Width           =   135
      End
      Begin VB.Image ImgBomb2 
         Height          =   480
         Left            =   3840
         Picture         =   "FrmMake.frx":000A
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgNWood 
         Height          =   480
         Left            =   720
         Picture         =   "FrmMake.frx":0C4C
         Top             =   0
         Width           =   480
      End
      Begin VB.Label etqTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Width           =   960
      End
      Begin VB.Image imgNCoin 
         Height          =   480
         Left            =   3000
         Picture         =   "FrmMake.frx":188E
         Top             =   0
         Width           =   480
      End
      Begin VB.Label etqBacktoGame 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back to Game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   480
         TabIndex        =   8
         Top             =   4080
         Width           =   1470
      End
      Begin VB.Label etqSaveGame 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save Game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3000
         TabIndex        =   7
         Top             =   4080
         Width           =   1230
      End
      Begin VB.Image imgNMagic 
         Height          =   480
         Left            =   1440
         Picture         =   "FrmMake.frx":24D0
         Top             =   0
         Width           =   480
      End
      Begin VB.Label etqNWood 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   135
      End
      Begin VB.Label etqNCoin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   5
         Top             =   480
         Width           =   135
      End
      Begin VB.Label etqNMagic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   135
      End
      Begin VB.Image imgNTicket 
         Height          =   480
         Left            =   120
         Picture         =   "FrmMake.frx":3112
         Top             =   0
         Width           =   480
      End
      Begin VB.Label etqNTicket 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   135
      End
      Begin VB.Image imgNToast 
         Height          =   480
         Left            =   2160
         Picture         =   "FrmMake.frx":3D54
         Top             =   0
         Width           =   480
      End
      Begin VB.Label etqNToast 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   3240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Move(I J K M)"
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   4215
      Left            =   7080
      TabIndex        =   24
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ImgWGrass 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmMake.frx":4996
      Top             =   7800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblMagic 
      AutoSize        =   -1  'True
      Caption         =   "Magic ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9120
      TabIndex        =   56
      Top             =   480
      Width           =   690
   End
   Begin VB.Label LblBunny 
      AutoSize        =   -1  'True
      Caption         =   "Bunny ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   55
      Top             =   480
      Width           =   705
   End
   Begin VB.Image Imgbun 
      Height          =   480
      Left            =   360
      Picture         =   "FrmMake.frx":55D8
      Top             =   7920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   12
      Left            =   5280
      Picture         =   "FrmMake.frx":621A
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   11
      Left            =   4800
      Picture         =   "FrmMake.frx":6E5C
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   6240
      TabIndex        =   50
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblWin 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1080
      TabIndex        =   48
      Top             =   120
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image ImgPGem 
      Height          =   480
      Left            =   2640
      Picture         =   "FrmMake.frx":7A9E
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGGem 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMake.frx":86E0
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgRGem 
      Height          =   480
      Left            =   1320
      Picture         =   "FrmMake.frx":9322
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgSwamp 
      Height          =   480
      Left            =   3000
      Picture         =   "FrmMake.frx":9F64
      Top             =   8880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblChrY 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9120
      TabIndex        =   33
      Top             =   8760
      Width           =   195
   End
   Begin VB.Label LblChrX 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7440
      TabIndex        =   32
      Top             =   8760
      Width           =   90
   End
   Begin VB.Label LblWeed 
      Caption         =   "Label7"
      Height          =   255
      Left            =   6000
      TabIndex        =   31
      Top             =   8400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image ImgCandle 
      Height          =   480
      Left            =   4080
      Picture         =   "FrmMake.frx":ABA6
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblRock 
      Caption         =   "Label7"
      Height          =   255
      Left            =   7680
      TabIndex        =   30
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgShield 
      Height          =   480
      Left            =   10200
      Picture         =   "FrmMake.frx":B7E8
      Top             =   7680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgArmor 
      Height          =   480
      Left            =   10200
      Picture         =   "FrmMake.frx":C42A
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBow 
      Height          =   480
      Left            =   10200
      Picture         =   "FrmMake.frx":D06C
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBomb 
      Height          =   480
      Left            =   11040
      Picture         =   "FrmMake.frx":DCAE
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgStairs 
      Height          =   480
      Left            =   9600
      Picture         =   "FrmMake.frx":E8F0
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgSword 
      Height          =   480
      Left            =   11040
      Picture         =   "FrmMake.frx":F532
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgKing 
      Height          =   480
      Left            =   9600
      Picture         =   "FrmMake.frx":10174
      Top             =   6960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Quests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   25
      Top             =   480
      Width           =   600
   End
   Begin VB.Image ImgMap 
      Height          =   480
      Left            =   9600
      Picture         =   "FrmMake.frx":10DB6
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBBottle 
      Height          =   480
      Left            =   9600
      Picture         =   "FrmMake.frx":119F8
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgYBottle 
      Height          =   480
      Left            =   9600
      Picture         =   "FrmMake.frx":1263A
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   10
      Left            =   9720
      Picture         =   "FrmMake.frx":1327C
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   9
      Left            =   4320
      Picture         =   "FrmMake.frx":13EBE
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   8
      Left            =   3840
      Picture         =   "FrmMake.frx":14B00
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   7
      Left            =   3360
      Picture         =   "FrmMake.frx":15742
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   6
      Left            =   2880
      Picture         =   "FrmMake.frx":16384
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   5
      Left            =   2400
      Picture         =   "FrmMake.frx":16FC6
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   4
      Left            =   1920
      Picture         =   "FrmMake.frx":17C08
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   3
      Left            =   1440
      Picture         =   "FrmMake.frx":1884A
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   2
      Left            =   960
      Picture         =   "FrmMake.frx":1948C
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   1
      Left            =   480
      Picture         =   "FrmMake.frx":1A0CE
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label label5 
      AutoSize        =   -1  'True
      Caption         =   "Map Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7440
      TabIndex        =   22
      Top             =   0
      Width           =   915
   End
   Begin VB.Image ImgBlue 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "FrmMake.frx":1AD10
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem6 
      Height          =   480
      Left            =   6960
      Picture         =   "FrmMake.frx":1B952
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMagic 
      Height          =   480
      Left            =   7560
      Picture         =   "FrmMake.frx":1C594
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBook 
      Height          =   480
      Left            =   11040
      Picture         =   "FrmMake.frx":1D1D6
      Top             =   8760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGem 
      Height          =   480
      Left            =   7560
      Picture         =   "FrmMake.frx":1D67B
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgSaw 
      Height          =   480
      Left            =   8640
      Picture         =   "FrmMake.frx":1E2BD
      Top             =   5760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBucket 
      Height          =   480
      Left            =   8640
      Picture         =   "FrmMake.frx":1EEFF
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgApple 
      Height          =   480
      Left            =   8640
      Picture         =   "FrmMake.frx":1FB41
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgLamp 
      Height          =   480
      Left            =   8640
      Picture         =   "FrmMake.frx":20783
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGold 
      Height          =   480
      Left            =   8040
      Picture         =   "FrmMake.frx":213C5
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgPine 
      Height          =   480
      Left            =   8040
      Picture         =   "FrmMake.frx":22007
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgKey2 
      Height          =   480
      Left            =   8040
      Picture         =   "FrmMake.frx":22C49
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8520
      TabIndex        =   16
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7440
      TabIndex        =   15
      Top             =   5160
      Width           =   135
   End
   Begin VB.Image imgSign 
      Height          =   480
      Left            =   1080
      Picture         =   "FrmMake.frx":2388B
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgGrass 
      Height          =   480
      Left            =   3480
      Picture         =   "FrmMake.frx":244CD
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   320
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Image imgBush 
      Height          =   480
      Left            =   1080
      Picture         =   "FrmMake.frx":2510F
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlack 
      Height          =   480
      Left            =   2040
      Picture         =   "FrmMake.frx":25D51
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR 
      Height          =   480
      Left            =   2520
      Picture         =   "FrmMake.frx":26993
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL 
      Height          =   480
      Left            =   1560
      Picture         =   "FrmMake.frx":275D5
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL 
      Height          =   480
      Left            =   1560
      Picture         =   "FrmMake.frx":28217
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL 
      Height          =   480
      Left            =   120
      Picture         =   "FrmMake.frx":28E59
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR 
      Height          =   480
      Left            =   600
      Picture         =   "FrmMake.frx":29A9B
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT 
      Height          =   480
      Left            =   2040
      Picture         =   "FrmMake.frx":2A6DD
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL 
      Height          =   480
      Left            =   2520
      Picture         =   "FrmMake.frx":2B31F
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR 
      Height          =   480
      Left            =   1560
      Picture         =   "FrmMake.frx":2BF61
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR 
      Height          =   480
      Left            =   2520
      Picture         =   "FrmMake.frx":2CBA3
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOR 
      Height          =   480
      Left            =   600
      Picture         =   "FrmMake.frx":2D7E5
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOL 
      Height          =   480
      Left            =   120
      Picture         =   "FrmMake.frx":2E427
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB 
      Height          =   480
      Left            =   2040
      Picture         =   "FrmMake.frx":2F069
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgUpChar 
      Height          =   480
      Left            =   5880
      Picture         =   "FrmMake.frx":2FCAB
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLeftChar 
      Height          =   480
      Left            =   5880
      Picture         =   "FrmMake.frx":308ED
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRightChar 
      Height          =   480
      Left            =   6360
      Picture         =   "FrmMake.frx":3152F
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDownChar 
      Height          =   480
      Left            =   6360
      Picture         =   "FrmMake.frx":32171
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy1 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmMake.frx":32DB3
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen1 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmMake.frx":339F5
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWindow1 
      Height          =   480
      Left            =   4560
      Picture         =   "FrmMake.frx":34637
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBlueTop 
      Height          =   480
      Left            =   4560
      Picture         =   "FrmMake.frx":35279
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   320
      X2              =   320
      Y1              =   328
      Y2              =   8
   End
   Begin VB.Image imgNothing 
      Height          =   480
      Left            =   1080
      Picture         =   "FrmMake.frx":35EBB
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDoor1 
      Height          =   480
      Left            =   5040
      Picture         =   "FrmMake.frx":36AFD
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRedTop 
      Height          =   480
      Left            =   5040
      Picture         =   "FrmMake.frx":3773F
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB2 
      Height          =   480
      Left            =   3480
      Picture         =   "FrmMake.frx":38381
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR2 
      Height          =   480
      Left            =   3960
      Picture         =   "FrmMake.frx":38FC3
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR2 
      Height          =   480
      Left            =   3000
      Picture         =   "FrmMake.frx":39C05
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL2 
      Height          =   480
      Left            =   3960
      Picture         =   "FrmMake.frx":3A847
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT2 
      Height          =   480
      Left            =   3480
      Picture         =   "FrmMake.frx":3B489
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL2 
      Height          =   480
      Left            =   3000
      Picture         =   "FrmMake.frx":3C0CB
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL2 
      Height          =   480
      Left            =   3000
      Picture         =   "FrmMake.frx":3CD0D
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR2 
      Height          =   480
      Left            =   3960
      Picture         =   "FrmMake.frx":3D94F
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallBottom 
      Height          =   480
      Left            =   5640
      Picture         =   "FrmMake.frx":3E591
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallTop 
      Height          =   480
      Left            =   5640
      Picture         =   "FrmMake.frx":3F1D3
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFloor1 
      Height          =   480
      Left            =   1080
      Picture         =   "FrmMake.frx":3FE15
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOL2 
      Height          =   480
      Left            =   120
      Picture         =   "FrmMake.frx":40A57
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOR2 
      Height          =   480
      Left            =   600
      Picture         =   "FrmMake.frx":41699
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR2 
      Height          =   480
      Left            =   600
      Picture         =   "FrmMake.frx":422DB
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL2 
      Height          =   480
      Left            =   120
      Picture         =   "FrmMake.frx":42F1D
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRockHill 
      Height          =   480
      Left            =   1560
      Picture         =   "FrmMake.frx":43B5F
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrees 
      Height          =   480
      Left            =   2040
      Picture         =   "FrmMake.frx":447A1
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy2 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmMake.frx":453E3
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharGoodWizard 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmMake.frx":46025
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBadWizard 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmMake.frx":46C67
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeed 
      Height          =   480
      Left            =   2520
      Picture         =   "FrmMake.frx":478A9
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWall2 
      Height          =   480
      Left            =   5640
      Picture         =   "FrmMake.frx":484EB
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB3 
      Height          =   480
      Left            =   3480
      Picture         =   "FrmMake.frx":4912D
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR3 
      Height          =   480
      Left            =   3960
      Picture         =   "FrmMake.frx":49D6F
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR3 
      Height          =   495
      Left            =   3000
      Picture         =   "FrmMake.frx":4A9B1
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL3 
      Height          =   480
      Left            =   3960
      Picture         =   "FrmMake.frx":4B653
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT3 
      Height          =   480
      Left            =   3480
      Picture         =   "FrmMake.frx":4C295
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL3 
      Height          =   480
      Left            =   3000
      Picture         =   "FrmMake.frx":4CED7
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL3 
      Height          =   480
      Left            =   3000
      Picture         =   "FrmMake.frx":4DB19
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR3 
      Height          =   480
      Left            =   3960
      Picture         =   "FrmMake.frx":4E75B
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExit 
      Height          =   480
      Left            =   3480
      Picture         =   "FrmMake.frx":4F39D
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgJar 
      Height          =   480
      Left            =   4440
      Picture         =   "FrmMake.frx":4FFDF
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExitB1 
      Height          =   480
      Left            =   1560
      Picture         =   "FrmMake.frx":50C21
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExitB2 
      Height          =   480
      Left            =   2520
      Picture         =   "FrmMake.frx":51863
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTable1 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMake.frx":524A5
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTable2 
      Height          =   480
      Left            =   4920
      Picture         =   "FrmMake.frx":530E7
      Top             =   6960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBed1 
      Height          =   480
      Left            =   5400
      Picture         =   "FrmMake.frx":53D29
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBed2 
      Height          =   480
      Left            =   5400
      Picture         =   "FrmMake.frx":5496B
      Top             =   6960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen2 
      Height          =   480
      Left            =   5640
      Picture         =   "FrmMake.frx":555AD
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen3 
      Height          =   480
      Left            =   5640
      Picture         =   "FrmMake.frx":561EF
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem1 
      Height          =   480
      Left            =   6360
      Picture         =   "FrmMake.frx":56E31
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarryChar 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMake.frx":57A73
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem2 
      Height          =   480
      Left            =   6360
      Picture         =   "FrmMake.frx":586B5
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "X"
      Height          =   195
      Left            =   7080
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      Height          =   195
      Left            =   7080
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgItem3 
      Height          =   480
      Left            =   6360
      Picture         =   "FrmMake.frx":592F7
      Top             =   5640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharSoldier 
      Height          =   480
      Left            =   5640
      Picture         =   "FrmMake.frx":59F39
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStarGate 
      Height          =   480
      Left            =   720
      Picture         =   "FrmMake.frx":5AB7B
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem4 
      Height          =   480
      Left            =   6360
      Picture         =   "FrmMake.frx":5B7BD
      Top             =   6120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem5 
      Height          =   480
      Left            =   6360
      Picture         =   "FrmMake.frx":5C3FF
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FrmMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i32 As Integer
Dim ScreenWidth As Integer
Dim ScreenHeight As Integer
Dim bBunnyTimer As Boolean
Private Sub CheckMygive()
  Dim i As Integer
  For i = 0 To 9
   ImgBlue(i).BorderStyle = 0
  Next
   imgNWood.BorderStyle = 0
   imgNCoin.BorderStyle = 0
   imgNMagic.BorderStyle = 0
   imgNTicket.BorderStyle = 0
   imgNToast.BorderStyle = 0
   imgNCoin.BorderStyle = 0
   ImgBomb2.BorderStyle = 0
End Sub

Public Sub CheckNeed(strItem As String)
   Dim i As Integer
   Dim j As Integer
   Dim strMap As String
   Dim MFixX As Integer
   Dim MFixY As Integer
   Dim iQuest As Integer
   Dim iQuestNo As Integer
   For i = 0 To 5
      Choice(i).Visible = False
   Next
   For Each ClsSpeech In nSpeech
   If ClsSpeech.Name = TxtSpeech.Text And ClsSpeech.Need = strItem Then
      If iNeedQty < ClsSpeech.NeedQty Then
         Pitanje.Caption = ClsSpeech.Name & "You do not have what I need"
         Choice(0).Caption = ">>> GoodBye"
         Choice(0).Visible = True
         Choice(0).Tag = "0000"
      GoTo MyTalk
      End If
     iQuest = Val(ClsSpeech.EndQuest)
     iQuestNo = Val(ClsSpeech.QuestNo)
     If (ClsSpeech.EndQuest = "0" Or Mid(MyQuest(iQuest), 1, 1) <> " ") And _
        (ClsSpeech.QuestNo = "0" Or Mid(MyQuest(iQuestNo), 1, 1) = "N") Then
        MFixX = Val(ClsSpeech.FixX)
        MFixY = Val(ClsSpeech.FixY)
        If strItem = "WOOD" Then Wood = Wood - ClsSpeech.NeedQty
        If strItem = "COIN" Then Coin = Coin - ClsSpeech.NeedQty
        If strItem = "TICKET" Then Ticket = Ticket - ClsSpeech.NeedQty
        If strItem = "TOAST" Then Toast = Toast - ClsSpeech.NeedQty
        If MyGiveLetter <> " " Then
           strMap = ""
             For j = 1 To Len(strInventory)
              If Mid(strInventory, j, 1) <> MyGiveLetter Then strMap = strMap & Mid(strInventory, j, 1)
            Next j
            strInventory = strMap
        End If
        MyGiveLetter = " "
        Pitanje.Caption = ClsSpeech.Name & ClsSpeech.Talk
        For i = 1 To ClsSpeech.Count
          'Load Choice(Choice.Count)
         Choice(i - 1).Left = 120
         If i > 1 Then Choice(i - 1).Top = Choice(i - 2).Top + Choice(i - 2).Height + 60
         Choice(i - 1).Caption = ">>> " & Right(ClsSpeech.Question(i - 1), Len(ClsSpeech.Question(i - 1)) - 5)
         Choice(i - 1).AutoSize = True
         Choice(i - 1).Visible = True
         Choice(i - 1).Tag = Left(ClsSpeech.Question(i - 1), 4)
        Next
        If ClsSpeech.DoQuest <> " " Then
          i = Val(ClsSpeech.DoQuest)
          If MyQuest(i) = " " Then
             MyQuest(i) = "N" & Trim(ClsSpeech.QuestName)
              iQuestTot = iQuestTot + 1
            End If
        End If
'        If ClsSpeech.EndQuest <> " " Then
'          i = Val(ClsSpeech.EndQuest)
'          If Left(MyQuest(i), 1) = "N" Then MyQuest(i) = "Y" & Right(MyQuest(i), Len(MyQuest(i)) - 1)
'        End If
         If ClsSpeech.EndQuest <> " " Then
          i = Val(ClsSpeech.EndQuest)
          If Left(MyQuest(i), 1) = "N" Then
             MyQuest(i) = "Y" & Right(MyQuest(i), Len(MyQuest(i)) - 1)
               iGiveQty = ClsSpeech.BombQty
               TheyGaveToMe ("BOMB")
                 iMyScore = iMyScore + 5
                IQuestDone = IQuestDone + 1
            End If
         End If
        If ClsSpeech.Give <> " " Then
          iGiveQty = ClsSpeech.GiveQty
          TheyGaveToMe (ClsSpeech.Give)
        End If
        
         '
        ' Fix The Bridge
        '
        If MFixX > 0 And MFixY > 0 Then
           strMap = Mid(Map(MFixY), 1, MFixX - 1) & "G" & Mid(Map(MFixY), MFixX + 1, Len(Map(MFixY)) - MFixX)
            Map(MFixY) = strMap
            DrawIt
        End If
       GoTo MyTalk
     End If
   End If
        'Text1.Text = Text1.Text & cEvent & " -- " & cDesc & Chr(13) & Chr(10)
    Next
     Pitanje.Caption = TxtSpeech.Text & "I do not need that item"
     Choice(0).Caption = ">>> GoodBye"
     Choice(0).Tag = "0000"
     Choice(0).Visible = True
MyTalk:
      Frame2.Visible = False
    ' PicTXT.Visible = True
      Holder.Visible = True
End Sub
  Public Sub Speak(Newx As Integer, Newy As Integer)
    Dim i As Integer
    For i = 0 To iMsgCnt
      Mx = Val(Mid(message(i), 1, 4))
      My = Val(Mid(message(i), 5, 4))
      MFixX = Val(Mid(message(i), 9, 4))
      MFixY = Val(Mid(message(i), 13, 4))
      MGive = Val(Mid(message(i), 17, 3))
      MNeed = Val(Mid(message(i), 21, 3))
      strNeed = Mid(message(i), 24, 1)
      strGive = Mid(message(i), 20, 1)
      '
      ' Check for Need
      '
      If Mx = Newx And My = Newy Then
        If strNeed = "0" Then

         ElseIf strNeed = "1" Then 'Wood
           If MNeed > Wood Then GoTo Mnext
           Wood = Wood - MNeed
         ElseIf strNeed = "2" Then 'coin
            If MNeed > Coin Then GoTo Mnext
            Coin = Coin - MNeed
         ElseIf strNeed = "3" Then 'Magic
          If MNeed > Magic Then GoTo Mnext
            Magic = Magic - MNeed
         ElseIf strNeed = "4" Then 'ticket
          If MNeed > Ticket Then GoTo Mnext
            Ticket = Ticket - MNeed
         ElseIf strNeed = "5" Then 'Toast
           If MNeed > Toast Then GoTo Mnext
            Toast = Toast - MNeed
         ElseIf strNeed = "6" Then 'weed
           If MNeed > Weed Then GoTo Mnext
            Weed = 0
         ElseIf strNeed = "7" Then 'rock
           If MNeed > Rock Then GoTo Mnext
            Rock = 0
         Else
           If InStr(1, strInventory, strNeed) = 0 Then GoTo Mnext
             strMap = ""
            For j = 1 To Len(strInventory)
              If Mid(strInventory, j, 1) <> strNeed Then strMap = strMap & Mid(strInventory, j, 1)
            Next j
            strInventory = strMap
       End If
        'End Select
        '
      ' Check for give
      '
       If strGive = "0" Then
        
         ElseIf strGive = "1" Then 'Wood
            Wood = Wood + MGive
         ElseIf strGive = "2" Then 'coin
            Coin = Coin + MGive
         ElseIf strGive = "3" Then 'Magic
             Magic = Magic + MGive
         ElseIf strGive = "4" Then 'ticket
            Ticket = MGive
         ElseIf strGive = "5" Then 'Toast
             Toast = MGive
         ElseIf strGive = "6" Then 'weed
            Weed = MGive
         ElseIf strGive = "7" Then 'rock
            Rock = MGive
         ElseIf strGive = "X" Then 'rock
           SpellCut = True
         ElseIf strGive = "Z" Then 'rock
           SpellDestroy = True
         Else
           strInventory = strInventory & MGive
        End If
        TxtSpeech.Visible = True
        TxtSpeech.ForeColor = vbWhite
        TxtSpeech.Text = Trim(Right(message(i), Len(message(i)) - 24))
        
        '
        ' Fix The Bridge
        '
        If MFixX > 0 And MFixY > 0 Then
           strMap = Mid(Map(MFixY), 1, MFixX - 1) & "G" & Mid(Map(MFixY), MFixX + 1, Len(Map(MFixY)) - MFixX)
          GoTo MDone
       End If
      End If
Mnext:
     Next
MDone:
End Sub
Public Sub JustSpeak(Newx As Integer, Newy As Integer)
    Dim i As Integer
    For i = 0 To iMsgCnt
      Mx = Val(Mid(message(i), 1, 4))
      My = Val(Mid(message(i), 5, 4))
'      MFixX = Val(Mid(Message(i), 9, 4))
'      MFixY = Val(Mid(Message(i), 13, 4))
'      MGive = Val(Mid(Message(i), 17, 3))
'      MNeed = Val(Mid(Message(i), 21, 3))
'      strNeed = Mid(Message(i), 24, 1)
'      strGive = Mid(Message(i), 20, 1)
      '
      ' Check for Need
      '
      If Mx = Newx And My = Newy Then
        TxtSpeech.Visible = True
        TxtSpeech.ForeColor = vbWhite
        TxtSpeech.Text = Trim(Right(message(i), Len(message(i)) - 16))
        
        '
        ' Fix The Bridge
        '
'        If MFixX > 0 And MFixY > 0 Then
'           strMap = Mid(Map(MFixY), 1, MFixX - 1) & "G" & Mid(Map(MFixY), MFixX + 1, Len(Map(MFixY)) - MFixX)
'          GoTo MDone
'       End If
      End If
Mnext:
     Next
MDone:
End Sub
Public Sub CharacterMovements()
  'Character Movements.
  Select Case CharFacing
  Case Is = 1
    PaintPicture imgUpChar.Picture, 5 * i32, 5 * i32, i32, i32
  Case Is = 2
    PaintPicture imgDownChar.Picture, 5 * i32, 5 * i32, i32, i32
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * i32, 5 * i32, i32, i32
  Case Is = 4
    PaintPicture imgLeftChar.Picture, 5 * i32, 5 * i32, i32, i32
  Case Is = 5
    PaintPicture imgCarryChar.Picture, 5 * i32, 5 * i32, i32, i32
  End Select
  'Text1 = "X=" & CharX & " Y=" & CharY
End Sub
Public Sub PictureHandler()
  PassToNext = 1
  PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
End Sub

Private Sub Choice_Click(Index As Integer)
  Dim cMsg As String
  Dim i As Integer
  Dim MFixX As Integer
  Dim MFixY As Integer
  For i = 0 To 6
     Choice(i).Visible = False
     Choice(i).Enabled = True
  Next
  If Choice(Index).Tag <> "0000" Then
   cMsg = "<" & Choice(Index).Tag & ">"
   For Each ClsSpeech In nSpeech
   If ClsSpeech.RNumber = cMsg Then
     If ClsSpeech.TakeAny = "Y" Then
          Pitanje.Caption = ClsSpeech.Name & ClsSpeech.Talk & TakeAny
        Else
          Pitanje.Caption = ClsSpeech.Name & ClsSpeech.Talk
        End If
     If ClsSpeech.Status = "WIN" Then
          lblWin.Caption = "You Win!"
          lblWin.Visible = True
        End If
        If ClsSpeech.Status = "LOSE" Then
          lblWin.Caption = "You Lose....."
          lblWin.Visible = True
        End If
    For i = 1 To ClsSpeech.Count
      'Load Choice(Choice.Count)
     Choice(i - 1).Left = 120
     If i > 1 Then Choice(i - 1).Top = Choice(i - 2).Top + Choice(i - 2).Height + 60
     Choice(i - 1).Caption = ">>> " & Right(ClsSpeech.Question(i - 1), Len(ClsSpeech.Question(i - 1)) - 5)
     Choice(i - 1).AutoSize = True
     Choice(i - 1).Visible = True
     Choice(i - 1).Tag = Left(ClsSpeech.Question(i - 1), 4)
    Next
       If ClsSpeech.DoQuest <> " " Then
         i = Val(ClsSpeech.DoQuest)
         If MyQuest(i) = " " Then
             MyQuest(i) = "N" & Trim(ClsSpeech.QuestName)
              iQuestTot = iQuestTot + 1
            End If
       End If
       If ClsSpeech.Give <> " " Then
          iGiveQty = ClsSpeech.GiveQty
          TheyGaveToMe (ClsSpeech.Give)
        End If
         If ClsSpeech.EndQuest <> " " Then
         If ClsSpeech.EndQuest <> " " Then
          i = Val(ClsSpeech.EndQuest)
          If Left(MyQuest(i), 1) = "N" Then
             MyQuest(i) = "Y" & Right(MyQuest(i), Len(MyQuest(i)) - 1)
               iGiveQty = ClsSpeech.BombQty
               If iGiveQty > 0 Then TheyGaveToMe ("BOMB")
                 iMyScore = iMyScore + 5
                IQuestDone = IQuestDone + 1
            End If
         End If
        End If
        MFixX = Val(ClsSpeech.FixX)
        MFixY = Val(ClsSpeech.FixY)
        If MFixX > 0 And MFixY > 0 Then
           strMap = Mid(Map(MFixY), 1, MFixX - 1) & "G" & Mid(Map(MFixY), MFixX + 1, Len(Map(MFixY)) - MFixX)
            Map(MFixY) = strMap
            DrawIt
        End If
    End If
        'Text1.Text = Text1.Text & cEvent & " -- " & cDesc & Chr(13) & Chr(10)
    Next
'     PicTXT.Visible = True
  Else
   Holder.Visible = False
   TxtSpeech.Visible = False
   bRun = True
   Timer2.Enabled = True
   
  End If
End Sub

Private Sub Choice_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 For i = 0 To Choice.Count - 1
        Choice(i).ForeColor = vbWhite
    Next
      Choice(Index).ForeColor = vbYellow
End Sub

Private Sub CmdAllQuests_Click()
 Dim i As Integer
   Dim cMsg As String
    FlexGrid.Rows = 1
    FlexGrid.Clear
    bMyMakeMove = False
     FlexGrid.FormatString = "Num|Quest                            |Map"
      For i = 1 To 95
     If MakeQuest(i) <> " " Then
       cMsg = i & Chr(9) & MakeQuest(i) & Chr(9) & QuestMap(i)
       FlexGrid.AddItem cMsg
     End If
   Next
   Command1.SetFocus
End Sub

Private Sub CmdMap_Click()
  SAVE_MapLoaded = "B1"
  strMyForm = "Frmgame"
  FrmMap.Show
End Sub

Private Sub CmdQuest_Click()
  Dim i As Integer
   Dim cMsg As String
    FlexGrid.Rows = 1
    bMyMakeMove = False
     FlexGrid.FormatString = "Num|Done|Quest                          "
    cMsg = Chr(9) & Chr(9) & "Score= " & iMyScore & " Quest_Done= " & IQuestDone & " Total=" & iQuestTot
    FlexGrid.AddItem cMsg
   For i = 1 To 95
     If MyQuest(i) <> " " Then
       cMsg = i & Chr(9) & Left(MyQuest(i), 1) & Chr(9) & Right(MyQuest(i), Len(MyQuest(i)) - 1)
       FlexGrid.AddItem cMsg
     End If
   Next
   Command1.SetFocus
End Sub

Private Sub CmdSpeech_Click()
   FrmSpeech.Show
End Sub

Private Sub CmdTest_Click()
  SAVE_MapLoaded = "B1"
  strMyForm = "Frmgame"
  FrmTester.Show
End Sub

Private Sub Command1_Click()
  Dim i As Integer
  i = CharX
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
  If bRun Then MoveMe (KeyCode)
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
  bRun = True
  'If bMyPlayGame Then
   'MoveMe (KeyCode)
 ' End If
End Sub
Private Sub MoveMe(KeyCode As Integer)
Dim i As Integer
Dim j As Integer
Dim Newx As Integer
Dim Newy As Integer
Dim Mx As Integer
Dim My As Integer
Dim MFixX, MFixY, MGive, MNeed As Integer
Dim strNeed As String
Dim strGive As String
Dim strMap As String
If Holder.Visible Then Exit Sub
Holder.Visible = False
TxtSpeech.Visible = False
etqNWood.Caption = Wood
etqNCoin.Caption = Coin
etqNMagic.Caption = Magic
Dim strNext As String
' If PositionMap = " " Then
'      TimBunny.Enabled = False
'      Mid(Map(CharY + 2), CharX + 3, 1) = "G"
'      iBunnyCaught = iBunnyCaught + 1
'      PositionMap = "G"
'       For i = 1 To iBunny
'         If iBunnyloc(i, 1) = CharY + 2 And iBunnyloc(i, 2) = CharX + 3 Then
'             iBunnyloc(i, 3) = -1
'        End If
'       Next
'       TimBunny.Enabled = True
'    End If
'    LblBunny.Caption = "Bunny = " & iBunnyCaught & "=" & PositionMap
strNext = ""
  DrawIt
 LblChrX.Caption = ""
 LblChrY.Caption = ""
Text1.Text = SAVE_MapLoaded
Frame2.Visible = True
Select Case KeyCode
  Case vbKeyUp, vbKeyI
    'Key Up.
    PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
    strNext = PositionMap
    Newx = CharX
    Newy = CharY - 1
    CharFacing = 1
    Label1 = CharY
    Call CharacterMovements
    'Text1 = "X=" & CharX & " Y=" & CharY
    If PositionMap = " " Then
      TimBunny.Enabled = False
      Mid(Map(CharY + 2), CharX + 3, 1) = "G"
      iBunnyCaught = iBunnyCaught + 1
      PositionMap = "G"
       For i = 1 To iBunny
         If iBunnyloc(i, 1) = CharY + 2 And iBunnyloc(i, 2) = CharX + 3 Then
             iBunnyloc(i, 3) = -1
        End If
       Next
       TimBunny.Enabled = True
    End If
    LblBunny.Caption = "Bunny = " & iBunnyCaught ' & "=" & PositionMap
    If TxtSpeech.Visible = True Then TxtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Or PositionMap = "c" Or PositionMap = "g" Or _
        PositionMap = "}" Or PositionMap = "|" Or PositionMap = ":" Or PositionMap = "." Or PositionMap = "&" Or _
        PositionMap = "L" Or PositionMap = "l" Or PositionMap = "N" Or PositionMap = "n" Or _
        PositionMap = "~" Or PositionMap = "O" Or PositionMap = "x" Or PositionMap = "o" Or PositionMap = "m" Or _
        PositionMap = "=" Or PositionMap = "%" Or PositionMap = "^" Or PositionMap = "{" Then
        CharY = CharY - 1
       My = CharY + 3
       Mx = CharX + 3
       If PositionMap = "G" Or PositionMap = "C" Then
         Else
             strInventory = strInventory & PositionMap
            strMap = Mid(Map(My), 1, Mx - 1) & "G" & Mid(Map(My), Mx + 1, Len(Map(My)) - Mx)
            Map(My) = strMap
            
         End If
        'CharY = CharY - 1
        DrawIt
    End If
  Case vbKeyQ
     If bZoom Then LargeScreen (1)
     End
   Case vbKey1
      Transport (1)
      DrawIt
   Case vbKey2
      Transport (2)
      DrawIt
   Case vbKey3
      Transport (3)
      DrawIt
   Case vbKeyZ
     TimBunny.Enabled = False
     LargeScreen (2)
     Timer2.Enabled = True
  Case vbKeyW
    DrawIt
  Case vbKeyDown, vbKeyM
    'Key Down.
    PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
    strNext = PositionMap
    CharFacing = 2
    Newx = CharX
    Newy = CharY + 1
    Label1 = CharY
    Call CharacterMovements
    If PositionMap = " " Then
      TimBunny.Enabled = False
      Mid(Map(CharY + 4), CharX + 3, 1) = "G"
      iBunnyCaught = iBunnyCaught + 1
      PositionMap = "G"
       For i = 1 To iBunny
         If iBunnyloc(i, 1) = CharY + 4 And iBunnyloc(i, 2) = CharX + 3 Then
             iBunnyloc(i, 3) = -1
        End If
       Next
       TimBunny.Enabled = True
    End If
    LblBunny.Caption = "Bunny = " & iBunnyCaught '& "=" & PositionMap
    'Text1 = "X=" & CharX & " Y=" & CharY
    If TxtSpeech.Visible = True Then TxtSpeech.Visible = False
'    If PositionMap = "G" Or PositionMap = "C" Then
'      CharY = CharY + 1
'      DrawIt
'    End If
    If PositionMap = "G" Or PositionMap = "C" Or PositionMap = "c" Or PositionMap = "g" Or _
        PositionMap = "}" Or PositionMap = "|" Or PositionMap = ":" Or PositionMap = "." Or PositionMap = "&" Or _
        PositionMap = "L" Or PositionMap = "l" Or PositionMap = "N" Or PositionMap = "n" Or _
        PositionMap = "~" Or PositionMap = "O" Or PositionMap = "x" Or PositionMap = "o" Or PositionMap = "m" Or _
        PositionMap = "=" Or PositionMap = "%" Or PositionMap = "^" Or PositionMap = "{" Then
       CharY = CharY + 1
       My = CharY + 3
       Mx = CharX + 3
          If PositionMap = "G" Or PositionMap = "C" Then
         Else
             strInventory = strInventory & PositionMap
            strMap = Mid(Map(My), 1, Mx - 1) & "G" & Mid(Map(My), Mx + 1, Len(Map(My)) - Mx)
            Map(My) = strMap
         End If
        'CharY = CharY - 1
        DrawIt
    End If
    'Text1 = "X=" & CharX & " Y=" & CharY
    If PositionMap = "M" Then
      For i = 0 To iMsgCnt
      Mx = Val(Mid(message(i), 1, 4))
      My = Val(Mid(message(i), 5, 4))
      If Mx = CharX And My = CharY Then
'          strMap = Mid(Message(i), 17, 3)
'          Call door3(strMap)
'          DrawIt
'       End If
          strMap = Mid(message(i), 17, 3)
          CharX = Val(Mid(message(i), 9, 4))
          CharY = Val(Mid(message(i), 13, 4))
            TimBunny.Enabled = False
            Call door4(strMap)
            If bBunnyTimer Then TimBunny.Enabled = True
          DrawIt
       End If
     Next
    End If
  Case vbKeyRight, vbKeyK
    'Key Right.
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
     strNext = PositionMap
    CharFacing = 3
    Label2 = CharX
    Newx = CharX + 1
    Newy = CharY
    Call CharacterMovements
    If PositionMap = " " Then
      TimBunny.Enabled = False
      Mid(Map(CharY + 3), CharX + 4, 1) = "G"
      iBunnyCaught = iBunnyCaught + 1
      PositionMap = "G"
       For i = 1 To iBunny
         If iBunnyloc(i, 1) = CharY + 3 And iBunnyloc(i, 2) = CharX + 4 Then
             iBunnyloc(i, 3) = -1
        End If
       Next
       TimBunny.Enabled = True
    End If
    LblBunny.Caption = "Bunny = " & iBunnyCaught '& "=" & PositionMap
    'Text1 = "X=" & CharX & " Y=" & CharY
    If TxtSpeech.Visible = True Then TxtSpeech.Visible = False
'    If PositionMap = "G" Or PositionMap = "C" Then
'      CharX = CharX + 1
'      DrawIt
'    End If
   If PositionMap = "G" Or PositionMap = "C" Or PositionMap = "c" Or PositionMap = "g" Or _
        PositionMap = "}" Or PositionMap = "|" Or PositionMap = ":" Or PositionMap = "." Or PositionMap = "&" Or _
        PositionMap = "L" Or PositionMap = "l" Or PositionMap = "N" Or PositionMap = "n" Or _
        PositionMap = "~" Or PositionMap = "O" Or PositionMap = "x" Or PositionMap = "o" Or PositionMap = "m" Or _
        PositionMap = "=" Or PositionMap = "%" Or PositionMap = "^" Or PositionMap = "{" Then
        CharX = CharX + 1
       My = CharY + 3
       Mx = CharX + 3
       If PositionMap = "G" Or PositionMap = "C" Then
         Else
             strInventory = strInventory & PositionMap
            strMap = Mid(Map(My), 1, Mx - 1) & "G" & Mid(Map(My), Mx + 1, Len(Map(My)) - Mx)
            Map(My) = strMap
         End If
        'CharY = CharY - 1
        DrawIt
    End If
    'Text1 = "X=" & CharX & " Y=" & CharY
  Case vbKeyLeft, vbKeyJ
    'Key Left.
    PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
     strNext = PositionMap
    CharFacing = 4
    Label2 = CharX
    Newx = CharX - 1
    Newy = CharY
    Call CharacterMovements
    If PositionMap = " " Then
      TimBunny.Enabled = False
      Mid(Map(CharY + 3), CharX + 2, 1) = "G"
      iBunnyCaught = iBunnyCaught + 1
      PositionMap = "G"
       For i = 1 To iBunny
         If iBunnyloc(i, 1) = CharY + 3 And iBunnyloc(i, 2) = CharX + 2 Then
             iBunnyloc(i, 3) = -1
        End If
       Next
       TimBunny.Enabled = True
    End If
    LblBunny.Caption = "Bunny = " & iBunnyCaught '& "=" & PositionMap
    ' If txtSpeech.Visible = True Then txtSpeech.Visible = False
'    If PositionMap = "G" Or PositionMap = "C" Then
'      CharX = CharX - 1
'      DrawIt
'    End If
     If PositionMap = "G" Or PositionMap = "C" Or PositionMap = "c" Or PositionMap = "g" Or _
        PositionMap = "}" Or PositionMap = "|" Or PositionMap = ":" Or PositionMap = "." Or PositionMap = "&" Or _
        PositionMap = "L" Or PositionMap = "l" Or PositionMap = "N" Or PositionMap = "n" Or _
        PositionMap = "~" Or PositionMap = "O" Or PositionMap = "x" Or PositionMap = "o" Or PositionMap = "m" Or _
        PositionMap = "=" Or PositionMap = "%" Or PositionMap = "^" Or PositionMap = "{" Then
        CharX = CharX - 1
       My = CharY + 3
       Mx = CharX + 3
       If PositionMap = "G" Or PositionMap = "C" Then
         Else
            strInventory = strInventory & PositionMap
            strMap = Mid(Map(My), 1, Mx - 1) & "G" & Mid(Map(My), Mx + 1, Len(Map(My)) - Mx)
            Map(My) = strMap
         End If
        'CharY = CharY - 1
        DrawIt
    End If
  Case vbKeyB
     Call BombWall
      DrawIt
   Case vbKeyS
      SaveGame
     Timer1.Enabled = True
      '
      ' Spells
      '
   Case vbKeyC
    If SpellCut = True Then
      Call Make_Spell_Cut
      DrawIt
    End If
  Case vbKeyF
    If SpellWade = True Then
      Call Make_Spell_Wade
      DrawIt
    End If
   Case vbKeyG
      Call Make_GrassGrow
      DrawIt
   Case vbKeyA
    If SpellAxe = True Then
      Call Make_Spell_Axe
      DrawIt
    End If
   Case vbKeyL
    If SpellLight = True Then
      Call Make_Spell_Light
      DrawIt
    End If
  Case vbKeyD
    If SpellDestroy = True Then
      Call Make_Spell_Destroy
      DrawIt
    End If
  
  Case vbKeyEscape
   ' CloseMidi
    'UnHook
  Case vbKeyReturn
    'Call ItemCount
    'fraMenu.Visible = True
  End Select
  '
  '
  If OpenJarDoor = "'" Then Call DrawIt
  If TypeOfItem = "!" Then Call DrawIt
     ' DrawIt
  '
  ' Star Gate or House Door or Stairs
  '
   If PositionMap = "#" Or PositionMap = "@" Or PositionMap = "?" Then
      For i = 0 To iMsgCnt
      Mx = Val(Mid(message(i), 1, 4))
      My = Val(Mid(message(i), 5, 4))
      If Mx = Newx And My = Newy Then
          strMap = Mid(message(i), 17, 3)
          CharX = Val(Mid(message(i), 9, 4))
          CharY = Val(Mid(message(i), 13, 4))
            TimBunny.Enabled = False
            Call door4(strMap)
            If bBunnyTimer Then TimBunny.Enabled = True
          DrawIt
       End If

     Next
   End If
     '
     ' Inventory
     '
    Dim strMe As String
  For i = 0 To 9
      ImgBlue(i).Visible = False
    'PaintPicture ImgBlue.Picture, (i - 1) * 32, 490
  Next
  If Len(strInventory) > 0 Then
     For i = 1 To Len(strInventory)
      strMe = Mid(strInventory, i, 1)
      ImgBlue(i - 1).Visible = True
      Select Case strMe
       Case Is = "c" 'Saw
        ImgBlue(i - 1).Picture = ImgSaw.Picture ', (i - 1) * 32, 490
      Case Is = "g" 'Bucket
       ImgBlue(i - 1).Picture = ImgBucket.Picture ', (i - 1) * 32, 490
      Case Is = "O" 'Lamp
        ImgBlue(i - 1).Picture = ImgLamp.Picture ', (i - 1) * 32, 490
      Case Is = "o" 'Lamp
        ImgBlue(i - 1).Picture = ImgGold.Picture ', (i - 1) * 32, 490
      Case Is = "x" 'Purplegem
        ImgBlue(i - 1).Picture = ImgPGem.Picture ', (i - 1) * 32, 490
      Case Is = "." 'Purplegem
        ImgBlue(i - 1).Picture = ImgGGem.Picture ', (i - 1) * 32, 490
      Case Is = "&" 'Purplegem
        ImgBlue(i - 1).Picture = ImgRGem.Picture ', (i - 1) * 32, 490
      Case Is = "m" 'Apple
        ImgBlue(i - 1).Picture = ImgApple.Picture ', (i - 1) * 32, 490
      Case Is = "L" 'YKey
        ImgBlue(i - 1).Picture = imgItem6.Picture ', (i - 1) * 32, 490
      Case Is = "l" 'Bottle
        ImgBlue(i - 1).Picture = ImgMagic.Picture ', (i - 1) * 32, 490
      Case Is = "~" 'RKey
        ImgBlue(i - 1).Picture = ImgKey2.Picture ', (i - 1) * 32, 490
      Case Is = "N" 'Book
       ImgBlue(i - 1).Picture = ImgBook.Picture ', (i - 1) * 32, 490
      Case Is = "n" 'Gem
       ImgBlue(i - 1).Picture = ImgGem.Picture ', (i - 1) * 32, 490
     Case Is = "=" 'Map
        ImgBlue(i - 1).Picture = ImgMap.Picture ', (i - 1) * 32, 490
     Case Is = "%" 'BBottle
       ImgBlue(i - 1).Picture = ImgBBottle.Picture ', (i - 1) * 32, 490
     Case Is = "^" 'YBottle
       ImgBlue(i - 1).Picture = ImgYBottle.Picture ', (i - 1) * 32, 490
     Case Is = "{" 'Sword
       ImgBlue(i - 1).Picture = ImgSword.Picture ', (i - 1) * 32, 490
     Case Is = "}" 'Bow
       ImgBlue(i - 1).Picture = ImgBow.Picture ', (i - 1) * 32, 490
     Case Is = "|" 'Armor
       ImgBlue(i - 1).Picture = ImgArmor.Picture ', (i - 1) * 32, 490
     Case Is = ":" 'Shield
       ImgBlue(i - 1).Picture = imgShield.Picture ', (i - 1) * 32, 490
     End Select
    Next
  End If
  '
  '
  '
  If strNext = "#" Or strNext = "@" Or strNext = "?" Then
     LblChrX.Caption = "DoorX " & Newx
     LblChrY.Caption = "DoorY " & Newy
  End If
  If strNext >= "0" And strNext <= "9" Then
   LblChrX.Caption = "CharX " & Newx
   LblChrY.Caption = "CharY " & Newy
     Call JustSpeak(Newx, Newy)
     Command7_Click
  End If
  If Weed > 0 Then Weed = Weed - 1
  If Rock > 0 Then Rock = Rock - 1
  'label5.Caption = Weed
  etqNWood.Caption = Wood
  etqNCoin.Caption = Coin
  etqNMagic.Caption = Magic
  etqNTicket.Caption = Ticket
  etqNToast.Caption = Toast
  etqNBomb.Caption = Bomb
  Text2.Text = CharX
  Text3.Text = CharY
  If bMyPlayGame Then ViewQuests
  Text1.Text = SAVE_MapLoaded
  LblRock.Caption = Rock
  LblWeed.Caption = Weed
   If bMyPlayGame Then
     '
     ' Determine Magic Points
     '
     iMagicPoints = 5 + iBunnyCaught * 5
     If SpellCut Then iMagicPoints = iMagicPoints + 5
     If SpellDestroy Then iMagicPoints = iMagicPoints + 5
     If SpellWade Then iMagicPoints = iMagicPoints + 5
     If SpellLight Then iMagicPoints = iMagicPoints + 5
     If SpellAxe Then iMagicPoints = iMagicPoints + 5
     'LblMagic.Caption = "Magic = " & iMagicPoints
   End If
    ShowLevel
 End Sub
Private Sub ShowLevel()
  If iMagicPoints < 15 Then
    LblMagic.Caption = "Level = None (" & iMagicPoints & ")"
  ElseIf iMagicPoints > 14 And iMagicPoints < 25 Then
    LblMagic.Caption = "Level = Thoughts (" & iMagicPoints & ")"
  ElseIf iMagicPoints > 24 And iMagicPoints < 35 Then
    LblMagic.Caption = "Level = Pence (" & iMagicPoints & ")"
  ElseIf iMagicPoints > 34 And iMagicPoints < 45 Then
    LblMagic.Caption = "Level = G - GrowGrass (" & iMagicPoints & ")"
  ElseIf iMagicPoints > 44 And iMagicPoints < 55 Then
    LblMagic.Caption = "Level = 1 - Transport 1 (" & iMagicPoints & ")"
  ElseIf iMagicPoints > 54 And iMagicPoints < 65 Then
    LblMagic.Caption = "Level = 2 - Transport 2 (" & iMagicPoints & ")"
  Else 'If iMagicPoints > 34 And iMagicPoints < 135 Then
    LblMagic.Caption = "Level = 3 - Transport 3 (" & iMagicPoints & ")"
  End If
End Sub

Private Sub Command2_Click()
  i32 = Text5.Text
  DrawIt
End Sub

Private Sub Command3_Click()
   Dim i As Integer
   Dim cMsg As String
   CharX = Text2.Text
   CharY = Text3.Text
   MeLoaded = Text1.Text
   TimBunny.Enabled = False
   ReadMapFile (MeLoaded)
   If bBunnyTimer Then TimBunny.Enabled = True
   'SAVE_MapLoaded = MeLoaded
   DrawIt
   Command1.SetFocus
   Text1.Text = MeLoaded
   FlexGrid.Rows = 1
   FlexGrid.FormatString = "    X|    Y|Name        "
   For i = 0 To iMsgCnt
     cMsg = Mid(message(i), 1, 4) & Chr(9) & Mid(message(i), 5, 4) & Chr(9) & Trim(Right(message(i), Len(message(i)) - 16))
     FlexGrid.AddItem cMsg
   Next
   i = CharX
   bMyMakeMove = True
End Sub

Private Sub Command4_Click()
    Toast = 10
    Coin = 55
    Wood = 1
    Magic = 10
    Bomb = 3
    iMagicPoints = iMagicPoints + 5
    ShowLevel
    etqNWood.Caption = Wood
    etqNCoin.Caption = Coin
   etqNMagic.Caption = Magic
   etqNTicket.Caption = Ticket
   etqNToast.Caption = Toast
   SpellCut = True
   SpellDestroy = True
   SpellWade = True
   SpellAxe = True
   SpellLight = True
   Command1.SetFocus
   strInventory = "L"
End Sub

Private Sub Command5_Click()
  'PaintPicture imgGrass.Picture, (Text2.Text) * 32, (Text3.Text) * 32
  '.fraMenu
  Dim strMe As String
  Dim i As Integer
  Dim j As Integer
  For i = 0 To 10
    PaintPicture ImgBlue(0).Picture, (i - 1) * 32, 490
  Next
  If Len(strInventory) > 0 Then
      j = Len(strInventory)
      If j > 10 Then j = 10
     For i = 1 To j
      strMe = Mid(strInventory, i, 1)
      Select Case strMe
       Case Is = "c" 'Saw
        PaintPicture ImgSaw.Picture, (i - 1) * 32, 490
      Case Is = "g" 'Bucket
        PaintPicture ImgBucket.Picture, (i - 1) * 32, 490
      Case Is = "O" 'Lamp
        PaintPicture ImgLamp.Picture, (i - 1) * 32, 490
      Case Is = "o" 'Lamp
        PaintPicture ImgGold.Picture, (i - 1) * 32, 490
      Case Is = "x" 'purple gem
        PaintPicture ImgPGem.Picture, (i - 1) * 32, 490
      Case Is = "." 'green gem
        PaintPicture ImgGGem.Picture, (i - 1) * 32, 490
      Case Is = "&" 'red gem
        PaintPicture ImgRGem.Picture, (i - 1) * 32, 490
      Case Is = "m" 'Apple
        PaintPicture ImgApple.Picture, (i - 1) * 32, 490
      Case Is = "L" 'Key
        PaintPicture imgItem6.Picture, (i - 1) * 32, 490
      Case Is = "l" 'Key
        PaintPicture ImgMagic.Picture, (i - 1) * 32, 490
      Case Is = "~" 'Key
        PaintPicture ImgKey2.Picture, (i - 1) * 32, 490
      Case Is = "N" 'Key
        PaintPicture ImgBook.Picture, (i - 1) * 32, 490
      Case Is = "n" 'Key
        PaintPicture ImgGem.Picture, (i - 1) * 32, 490
      End Select
    Next
  End If
End Sub



Private Sub Command6_Click()
 Label7.Caption = iJarCnt
  lblWin.Visible = True
End Sub

Private Sub Command7_Click()
 Dim cMsg As String
 Dim iQuest As Integer
 Dim iQuestYes As Integer
 Dim iQuestNo As Integer
 Dim i As Integer
 Dim iAdd As Integer
 Dim MFixX As Integer
 Dim MFixY As Integer
 Dim bPence As Boolean
 Dim bWaitP As Boolean
 Dim bWaitT As Boolean
 Dim strPence As String
 Dim strPenceSay As String
 Frame2.Visible = False
 bPence = False
 bWaitP = False
 bWaitT = False
 strPence = ""
 strPenceSay = ""
 '
 ' Find Pence
 '
  For Each ClsThought In nThought
  If UCase(ClsThought.Person) = UCase(TxtSpeech.Text) Then
      If iMagicPoints < 25 And ClsThought.Waitp Then bWaitP = True
      If iMagicPoints < 15 And ClsThought.Waitt Then bWaitT = True
      If iMagicPoints > 24 Then
        If ClsThought.Pence <> "" Then
          strPenceSay = ":::Pence " & ClsThought.Pence
          bPence = True
         End If
      End If
  End If
  Next
   '
 For i = 0 To 6
  Choice(i).Enabled = True
  Choice(i).Visible = False
 Next
 bRun = False
 Holder.Visible = True
 Pitanje.Visible = True
 iQuest = 0
 iAdd = 0
 iQuestYes = 0
 cMsg = "<0001>"
 '
 '
 '
 If bWaitT Or bWaitP Then GoTo MyTalk
  For Each ClsSpeech In nSpeech
   If ClsSpeech.Name = TxtSpeech.Text Then
     iQuest = Val(ClsSpeech.EndQuest)
     iQuestYes = Val(ClsSpeech.QuestYes)
     iQuestNo = Val(ClsSpeech.QuestNo)
     If UCase(ClsSpeech.Need) = "WEED" And Weed < ClsSpeech.NeedQty Then GoTo MySkipTalk
     If UCase(ClsSpeech.Need) = "ROCK" And Rock < ClsSpeech.NeedQty Then GoTo MySkipTalk
     If bPence Then
        strPence = Mid(ClsSpeech.RNumber, 2, 4)
        GoTo MyTalk
     End If
     If UCase(ClsSpeech.Need) = "WEED" Then
       Weed = 0
     End If
     If UCase(ClsSpeech.Need) = "ROCK" Then Rock = 0
     If (ClsSpeech.EndQuest = "0" Or Mid(MyQuest(iQuest), 1, 1) <> " ") And ClsSpeech.SayOnce <> "D" And _
        (ClsSpeech.QuestYes = "0" Or Mid(MyQuest(iQuestYes), 1, 1) = "Y") And _
        (ClsSpeech.QuestNo = "0" Or Mid(MyQuest(iQuestNo), 1, 1) = "N") Then
        If ClsSpeech.TakeAny = "Y" Then
          Pitanje.Caption = ClsSpeech.Name & ClsSpeech.Talk & TakeAny
        Else
          Pitanje.Caption = ClsSpeech.Name & ClsSpeech.Talk
        End If
        If ClsSpeech.SayOnce = "Y" Then
            ClsSpeech.SayOnce = "D"
        End If
        iAdd = ClsSpeech.Count
        For i = 1 To ClsSpeech.Count
          'Load Choice(Choice.Count)
         Choice(i - 1).Left = 120
         If i > 1 Then Choice(i - 1).Top = Choice(i - 2).Top + Choice(i - 2).Height + 60
         Choice(i - 1).Caption = ">>> " & Right(ClsSpeech.Question(i - 1), Len(ClsSpeech.Question(i - 1)) - 5)
         Choice(i - 1).AutoSize = True
         Choice(i - 1).Visible = True
         Choice(i - 1).Tag = Left(ClsSpeech.Question(i - 1), 4)
        Next
        If ClsSpeech.DoQuest <> " " Then
          i = Val(ClsSpeech.DoQuest)
          If MyQuest(i) = " " Then
             MyQuest(i) = "N" & Trim(ClsSpeech.QuestName)
              iQuestTot = iQuestTot + 1
            End If
        End If
        If ClsSpeech.EndQuest <> " " Then
          i = Val(ClsSpeech.EndQuest)
          If Left(MyQuest(i), 1) = "N" Then
             MyQuest(i) = "Y" & Right(MyQuest(i), Len(MyQuest(i)) - 1)
               iGiveQty = ClsSpeech.BombQty
               If iGiveQty > 0 Then TheyGaveToMe ("BOMB")
                iMyScore = iMyScore + 5
                 IQuestDone = IQuestDone + 1
            End If
         End If
        If ClsSpeech.Status = "WIN" Then
          lblWin.Caption = "You Win!"
          lblWin.Visible = True
        End If
        If ClsSpeech.Status = "LOSE" Then
          lblWin.Caption = "You Lose....."
          lblWin.Visible = True
        End If
        If ClsSpeech.Give <> " " Then
          iGiveQty = ClsSpeech.GiveQty
          TheyGaveToMe (ClsSpeech.Give)
        End If
        MFixX = Val(ClsSpeech.FixX)
        MFixY = Val(ClsSpeech.FixY)
        If MFixX > 0 And MFixY > 0 Then
           strMap = Mid(Map(MFixY), 1, MFixX - 1) & "G" & Mid(Map(MFixY), MFixX + 1, Len(Map(MFixY)) - MFixX)
            Map(MFixY) = strMap
            DrawIt
        End If
       GoTo MyTalk
     End If
MySkipTalk:
   End If
        'Text1.Text = Text1.Text & cEvent & " -- " & cDesc & Chr(13) & Chr(10)
    Next
MyTalk:
  '
  ' Show Thought and pence
  '
  If bWaitT Or bWaitP Then
         Pitanje.Caption = ClsSpeech.Name & " I can not talk to you now"
         Choice(iAdd).Caption = ">>>Bye "
         If iAdd > 0 Then Choice(iAdd).Top = Choice(iAdd - 1).Top + Choice(iAdd - 1).Height + 60
         Choice(iAdd).Visible = True
         Choice(iAdd).Left = 120
         Choice(iAdd).Tag = "0000"
         iAdd = iAdd + 1
  Else
   If iMagicPoints > 14 Then
     For Each ClsThought In nThought
       If UCase(ClsThought.Person) = UCase(TxtSpeech.Text) Then
         If bPence Then
           cMsg = ""
          If ClsThought.Givep <> "" Then
             cMsg = " Here is a " & ClsThought.Givep
             TheyGaveToMe (ClsThought.Givep)
             ClsThought.Givep = ""
           End If
            Pitanje.Caption = ClsSpeech.Name
            Choice(iAdd).Caption = ">>>Pence(click to talk to me) " & ClsThought.Pence & cMsg
            If iAdd > 0 Then Choice(iAdd).Top = Choice(iAdd - 1).Top + Choice(iAdd - 1).Height + 60
            Choice(iAdd).Visible = True
            Choice(iAdd).Left = 120
            Choice(iAdd).Tag = strPence
            iAdd = iAdd + 1
             Choice(iAdd).Caption = ">>>Bye "
            If iAdd > 0 Then Choice(iAdd).Top = Choice(iAdd - 1).Top + Choice(iAdd - 1).Height + 60
            Choice(iAdd).Visible = True
            Choice(iAdd).Left = 120
            Choice(iAdd).Tag = "0000"
            iAdd = iAdd + 1
         End If
         If ClsThought.HideThought And bPence = False Then
               Choice(iAdd).Caption = ":::Thought ?????"
         Else
           cMsg = ""
          If ClsThought.Givet <> "" Then
             cMsg = " Here is a " & ClsThought.Givet
             TheyGaveToMe (ClsThought.Givet)
             ClsThought.Givet = ""
           End If
           Choice(iAdd).Caption = ":::Thought " & ClsThought.Thought & cMsg
         End If
         Choice(iAdd).Enabled = False
         If iAdd > 0 Then Choice(iAdd).Top = Choice(iAdd - 1).Top + Choice(iAdd - 1).Height + 60
         Choice(iAdd).Visible = True
         Choice(iAdd).Left = 120
       End If
     Next
   End If
  End If
    Holder.Visible = True
End Sub




Private Sub Command8_Click()
   If TimBunny.Enabled Then
         bBunnyTimer = True
        TimBunny.Enabled = True
        Command8.Caption = "Bunny move"
   Else
       TimBunny.Enabled = False
       Command8.Caption = "Bunny stop"
       bBunnyTimer = False
   End If
'  Dim i As Integer
'  Dim j As Integer
'  j = 1
'  i = iBunnyloc(1, 3)
'  Select Case i
'    Case 1
'    'right
'        If Mid(Map(iBunnyloc(1, 1)), iBunnyloc(1, 2) + 1, 1) = "G" Then
'          Mid(Map(iBunnyloc(1, 1)), iBunnyloc(1, 2) + 1, 1) = " "
'          Mid(Map(iBunnyloc(1, 1)), iBunnyloc(1, 2), 1) = "G"
'        Else
'         iBunnyloc(j, 3) = 2
'        End If
'     Case 2
'        If Mid(Map(iBunnyloc(1, 1) - 1), iBunnyloc(1, 2), 1) = "G" Then
'          '
'          'up
'          Mid(Map(iBunnyloc(1, 1) - 1), iBunnyloc(1, 2), 1) = " "
'          Mid(Map(iBunnyloc(1, 1)), iBunnyloc(1, 2), 1) = "G"
'        Else
'         iBunnyloc(j, 3) = 3
'        End If
'    Case 3
'      'left
'        If Mid(Map(iBunnyloc(1, 1)), iBunnyloc(1, 2) - 1, 1) = "G" Then
'           Mid(Map(iBunnyloc(1, 1)), iBunnyloc(1, 2) - 1, 1) = " "
'           Mid(Map(iBunnyloc(1, 1)), iBunnyloc(1, 2), 1) = "G"
'        Else
'         iBunnyloc(j, 3) = 4
'        End If
'    Case 4
'    'down
'      If Mid(Map(iBunnyloc(1, 1) + 1), iBunnyloc(1, 2), 1) = "G" Then
'       Mid(Map(iBunnyloc(1, 1) + 1), iBunnyloc(1, 2), 1) = " "
'       Mid(Map(iBunnyloc(1, 1)), iBunnyloc(1, 2), 1) = "G"
'      Else
'         iBunnyloc(j, 3) = 1
'       End If
' End Select
'  DrawIt
End Sub

Private Sub FlexGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bMyMakeMove = False Then Exit Sub
   If FlexGrid.MouseRow <> 0 Then
     FlexGrid.Col = 0
     Text2.Text = FlexGrid.Text
     FlexGrid.Col = 1
     Text3.Text = FlexGrid.Text + 1
     Command3_Click
   End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If bMyPlayGame Then
      If bRun Then MoveMe (KeyCode)
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    bRun = True
'   If bMyPlayGame Then
'      MoveMe (KeyCode)
'  End If
End Sub
Private Sub ViewQuests()
   Dim i As Integer
   Dim cMsg As String
   FlexGrid.Rows = 1
   iMyScore = 0
    iQuestTot = 0
     IQuestDone = 0
     For i = 1 To 95
     If Mid(MyQuest(i), 1, 1) <> " " Then iQuestTot = iQuestTot + 1
     If Mid(MyQuest(i), 1, 1) = "Y" Then
       IQuestDone = IQuestDone + 1
       iMyScore = iMyScore + 5
     End If
   Next
      cMsg = Chr(9) & "Score= " & iMyScore & " Done= " & IQuestDone & " Total=" & iQuestTot
    FlexGrid.AddItem cMsg
   For i = 1 To 95
     If Mid(MyQuest(i), 1, 1) = "N" Then
        cMsg = Left(MyQuest(i), 1) & Chr(9) & Right(MyQuest(i), Len(MyQuest(i)) - 1)
       FlexGrid.AddItem cMsg
     End If
   Next
   For i = 1 To 95
     If Mid(MyQuest(i), 1, 1) = "Y" Then
       cMsg = Left(MyQuest(i), 1) & Chr(9) & Right(MyQuest(i), Len(MyQuest(i)) - 1)
       FlexGrid.AddItem cMsg
     End If
   Next
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Randomize
    bBunnyTimer = False
    i32 = 32
    Timer1.Enabled = False
    TxtSpeech.Enabled = False
    Wood = SAVE_Wood
    Coin = SAVE_Coin
    Magic = SAVE_Magic
    Bomb = SAVE_Bomb
    bRun = True
    LblBunny.Caption = "Bunny = " & iBunnyCaught
    bZoom = False
    Holder.Visible = False
    Mygive = ""
    iMyScore = 0
    iQuestTot = 0
     IQuestDone = 0
    'strInventory = ""
    bMyMakeMove = False
    strMyForm = "FrmMake"
    i = CharX
    ItemHouseFound1 = False
    ItemHouseFound2 = False
    ItemHouseFound3 = False
    ItemHouseFound4 = False
   For i = 1 To 6
      Load Choice(Choice.Count)
   Next
'   For i = 0 To 99
'    MyQuest(i) = " "
'   Next
   'iSaveMapcnt = 0
   'iSaveMapLast = 0
   'iSaveMapLoc = 0
   strSaveMapName(0, 1) = " "
   strSaveMapName(0, 2) = " "
   strSaveMapName(0, 3) = " "
   If bMyPlayGame Then
      i32 = 40
      bBunnyTimer = True
      
      Holder.Top = 48
      Holder.Left = 36
      FlexGrid.Left = 400
      Label6.Left = 400
      LblBunny.Left = LblBunny.Left - 75
      LblMagic.Left = LblMagic.Left - 75
      Label5.Left = 400
      Text1.Left = 500
      Timer2.Enabled = True
      LblChrX.Visible = False
     LblChrY.Visible = False
     LblRock.Visible = False
     LblWeed.Visible = False
     Command3.Visible = False
     Command1.Visible = False
     Command4.Visible = False
     Command8.Visible = False
     CmdSpeech.Visible = False
     CmdMap.Visible = False
     CmdQuest.Visible = False
     CmdAllQuests.Visible = False
     Text2.Enabled = False
     Text3.Enabled = False
     Text1.Enabled = False
     MeLoaded = SAVE_MapLoaded
     ReadMapFile (MeLoaded)
     'DrawIt
     CharX = SAVE_CharX
     CharY = SAVE_CharY
     Call ViewQuests
     Timer1.Interval = 200
     Timer1.Enabled = True
     TimBunny.Enabled = True
     Me.Caption = "Play Game"
     FlexGrid.Enabled = False
     FlexGrid.Rows = 1
     FlexGrid.FormatString = "Done|Quest                                "
    Else
      For i = 0 To 99
         MyQuest(i) = " "
         MakeQuest(i) = " "
      Next
      iSaveMapcnt = 0
     iSaveMapLast = 0
      iSaveMapLoc = 0
   End If
   ShowLevel
End Sub



  Public Sub DrawIt()
     For Y = -3 To 6
    For X = -3 To 6
      'If the result to Paint is 0 then it will get error.
      'This will prevent this.
      PassToNext = 0
        If Y + CharY + 0 < 1 Then PictureHandler
        If X + CharX + 0 < 1 Then PictureHandler
        If X + CharX + 0 > Len(Map(1)) Then PictureHandler
        If Y + CharY + 0 > 99 Then PictureHandler
      If PassToNext = 0 Then PositionMap = Mid(Map(Y + CharY + 1), (X + CharX + 1), 1)
      'If X = 0 And Y = 0 Then GoTo skip:
      Select Case PositionMap
      Case Is = "`" 'green gem
        PaintPicture ImgSwamp.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = " " 'bunny
        PaintPicture Imgbun.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
'        iBunny = iBunny + 1
'        iBunnyloc(iBunny, 1) = Y + CharY + 1
'        iBunnyloc(iBunny, 2) = X + CharX + 1
      Case Is = "." 'green gem
        PaintPicture ImgGGem.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = Chr(34) 'green gem
        PaintPicture ImgWGrass.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = " " 'bunny
        PaintPicture Imgbun.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "'" 'green gem
        PaintPicture imgJar.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "&" 'red gem
        PaintPicture ImgRGem.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "$" 'Candle
        PaintPicture ImgCandle.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "+" 'Swamp
       PaintPicture ImgSwamp.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "?" 'Grass
        PaintPicture ImgStairs.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "}" 'Bow
        PaintPicture ImgBow.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "|" 'Armor
        PaintPicture ImgArmor.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = ":" 'Shield
        PaintPicture imgShield.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "P" 'Wall-Door
        PaintPicture imgWallBottom.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "p" 'Weed-Door
        PaintPicture imgWeed.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "{" 'Sword
        PaintPicture ImgSword.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "c" 'Saw
        PaintPicture ImgSaw.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "=" 'Map
        PaintPicture ImgMap.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "%" 'BBottle
        PaintPicture ImgBBottle.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "^" 'YBottle
        PaintPicture ImgYBottle.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "g" 'Bucket
        PaintPicture ImgBucket.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "O" 'Lamp
        PaintPicture ImgLamp.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "o" 'Lamp
        PaintPicture ImgGold.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "x" 'Purple gem
        PaintPicture ImgPGem.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "m" 'Saw
        PaintPicture ImgApple.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "L" 'Key
        PaintPicture imgItem6.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "l" 'Key
        PaintPicture ImgMagic.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "~" 'Key
        PaintPicture ImgKey2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "N" 'Key
        PaintPicture ImgBook.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "n" 'Blue gem
        PaintPicture ImgGem.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "b" '2 bush
        PaintPicture imgTrees.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "<" 'Weed
        PaintPicture imgWeed.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = ">" 'Rock Hill
        PaintPicture imgRockHill.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "0" 'Cero Sign
        PaintPicture imgSign.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      
      Case Is = "Q" 'Water
        PaintPicture imgTOL.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "q" 'grass
        PaintPicture imgTOL2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "A" 'Water
        PaintPicture imgBOL.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "a" 'grass
        PaintPicture imgBOL2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "W" 'Water
        PaintPicture imgTOR.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "w" 'grass
        PaintPicture imgTOR2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "S" 'Water
        PaintPicture imgBOR.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "s" 'grass
        PaintPicture imgBOR2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      
      Case Is = "E" 'Border Left water
        PaintPicture ImgIL.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "e" 'Border Left grass
        PaintPicture ImgIL2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "R" 'Border Right water
        PaintPicture ImgIR.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "r" 'Border Right grass
        PaintPicture ImgIR2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "D" 'Border Top water
        PaintPicture ImgIT.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "d" 'Border Top grass
        PaintPicture ImgIT2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "F" 'Border Bottom water
        PaintPicture ImgIB.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "f" 'Border Bottom grass
        PaintPicture ImgIB2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      
      Case Is = "T" 'Border Bottom water
        PaintPicture ImgTIL.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "t" 'Border Bottom grass
        PaintPicture ImgTIL2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "Y" 'Border Bottom water
        PaintPicture ImgTIR.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "y" 'Border Bottom grass
        PaintPicture ImgTIR2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "V" 'Border Bottom water
        PaintPicture ImgBIL.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "v" 'Border Bottom grass
        PaintPicture ImgBIL2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "H" 'Border Bottom water
        PaintPicture ImgBIR.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "h" 'Border Bottom grass
        PaintPicture ImgBIR2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      
      Case Is = "U" 'Water
        PaintPicture ImgTIL3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "u" 'grass
        PaintPicture ImgIR3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "I" 'Water
        PaintPicture ImgTIR3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "i" 'grass
        PaintPicture ImgIB3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "J" 'Water
        PaintPicture ImgBIL3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "j" 'grass
        PaintPicture ImgIL3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "K" 'Water
        PaintPicture ImgBIR3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "k" 'grass
        PaintPicture ImgIT3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "," 'grass
        PaintPicture imgHouseExitB1.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = ";" 'grass
        PaintPicture imgHouseExitB2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "M" 'grass
        PaintPicture imgHouseExit.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      
      Case Is = "-" 'Water
        PaintPicture ImgBlack.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "_" 'Nothing
        PaintPicture imgNothing.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      
      Case Is = "Z" 'Wall bottom
        PaintPicture imgWallBottom.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "X" 'Wall top
        PaintPicture imgWallTop.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "C" 'Floor Blue 1
        PaintPicture imgFloor1.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "z" 'Floor Tronco
        PaintPicture imgWall2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
            
      Case Is = "1" 'Laddy
        PaintPicture imgCharWomen1.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "2" 'Boy 1
        PaintPicture imgCharBoy1.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "3" 'Boy 2
        PaintPicture imgCharBoy2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "4" 'Good Wizard
        PaintPicture imgCharGoodWizard.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "5" 'Bad Wizard
        PaintPicture imgCharBadWizard.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "6" 'Lady
        PaintPicture imgCharWomen2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "7" 'Lady
        PaintPicture imgCharWomen3.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "8" 'Soldier
        PaintPicture imgCharSoldier.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "9" 'King
        PaintPicture ImgKing.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      
      
      Case Is = "/"
        PaintPicture imgBlueTop.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "\"
        PaintPicture imgRedTop.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "*"
        PaintPicture imgWindow1.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "#" '3
        'PaintPicture imgNothing.Picture, (X + 3) * i32, (Y + 3) * i32, i32 , i32
        PaintPicture imgDoor1.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "@" '3
        PaintPicture imgStarGate.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
    
      Case Is = "!" 'Jar
        PaintPicture imgJar.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "(" 'Table
        PaintPicture imgTable1.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = ")" 'Table2
        PaintPicture imgTable2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "[" 'Bed
        PaintPicture imgBed1.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
      Case Is = "]" 'Bed2
        PaintPicture imgBed2.Picture, (X + 3) * i32, (Y + 3) * i32, i32, i32
    
    
      End Select
skip:
    Next
  Next
  

  'Character Movements.
  Select Case CharFacing
  Case Is = 1
    PaintPicture imgUpChar.Picture, 5 * i32, 5 * i32, i32, i32
  Case Is = 2
    PaintPicture imgDownChar.Picture, 5 * i32, 5 * i32, i32, i32
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * i32, 5 * i32, i32, i32
  Case Is = 4
    PaintPicture imgLeftChar.Picture, 5 * i32, 5 * i32, i32, i32
  Case Is = 5
    PaintPicture imgCarryChar.Picture, 5 * i32, 5 * i32, i32, i32
  End Select
  
  Select Case Item
  Case Is = 1
    PaintPicture imgItem1.Picture, 5 * i32, 4 * i32, i32, i32
    Item = 0
  Case Is = 2
    PaintPicture imgItem2.Picture, 5 * i32, 4 * i32, i32, i32
    Item = 0
  Case Is = 3
    PaintPicture imgItem3.Picture, 5 * i32, 4 * i32, i32, i32
    Item = 0
  End Select
End Sub
Public Sub DrawItold()
  For Y = -3 To 6
    For X = -3 To 6
      'If the result to Paint is 0 then it will get error.
      'This will prevent this.
      PassToNext = 0
        If Y + CharY + 0 < 1 Then PictureHandler
        If X + CharX + 0 < 1 Then PictureHandler
        If X + CharX + 0 > Len(Map(1)) Then PictureHandler
        If Y + CharY + 0 > 99 Then PictureHandler
      If PassToNext = 0 Then PositionMap = Mid(Map(Y + CharY + 1), (X + CharX + 1), 1)
      'If X = 0 And Y = 0 Then GoTo skip:
      Select Case PositionMap
      Case Is = "`" 'green gem
        PaintPicture ImgSwamp.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "." 'green gem
        PaintPicture ImgGGem.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "'" 'green gem
        PaintPicture imgJar.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "&" 'red gem
        PaintPicture ImgRGem.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "$" 'Candle
        PaintPicture ImgCandle.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "+" 'Swamp
       PaintPicture ImgSwamp.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "?" 'Grass
        PaintPicture ImgStairs.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "}" 'Bow
        PaintPicture ImgBow.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "|" 'Armor
        PaintPicture ImgArmor.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ":" 'Shield
        PaintPicture imgShield.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "P" 'Wall-Door
        PaintPicture imgWallBottom.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "p" 'Weed-Door
        PaintPicture imgWeed.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "{" 'Sword
        PaintPicture ImgSword.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "c" 'Saw
        PaintPicture ImgSaw.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "=" 'Map
        PaintPicture ImgMap.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "%" 'BBottle
        PaintPicture ImgBBottle.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "^" 'YBottle
        PaintPicture ImgYBottle.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "g" 'Bucket
        PaintPicture ImgBucket.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "O" 'Lamp
        PaintPicture ImgLamp.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "o" 'Lamp
        PaintPicture ImgGold.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "x" 'Purple gem
        PaintPicture ImgPGem.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "m" 'Saw
        PaintPicture ImgApple.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "L" 'Key
        PaintPicture imgItem6.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "l" 'Key
        PaintPicture ImgMagic.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "~" 'Key
        PaintPicture ImgKey2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "N" 'Key
        PaintPicture ImgBook.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "n" 'Blue gem
        PaintPicture ImgGem.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "b" '2 bush
        PaintPicture imgTrees.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "<" 'Weed
        PaintPicture imgWeed.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ">" 'Rock Hill
        PaintPicture imgRockHill.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "0" 'Cero Sign
        PaintPicture imgSign.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "Q" 'Water
        PaintPicture imgTOL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "q" 'grass
        PaintPicture imgTOL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "A" 'Water
        PaintPicture imgBOL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "a" 'grass
        PaintPicture imgBOL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "W" 'Water
        PaintPicture imgTOR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "w" 'grass
        PaintPicture imgTOR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "S" 'Water
        PaintPicture imgBOR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "s" 'grass
        PaintPicture imgBOR2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "E" 'Border Left water
        PaintPicture ImgIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "e" 'Border Left grass
        PaintPicture ImgIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "R" 'Border Right water
        PaintPicture ImgIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "r" 'Border Right grass
        PaintPicture ImgIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "D" 'Border Top water
        PaintPicture ImgIT.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "d" 'Border Top grass
        PaintPicture ImgIT2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "F" 'Border Bottom water
        PaintPicture ImgIB.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "f" 'Border Bottom grass
        PaintPicture ImgIB2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "T" 'Border Bottom water
        PaintPicture ImgTIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "t" 'Border Bottom grass
        PaintPicture ImgTIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "Y" 'Border Bottom water
        PaintPicture ImgTIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "y" 'Border Bottom grass
        PaintPicture ImgTIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "V" 'Border Bottom water
        PaintPicture ImgBIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "v" 'Border Bottom grass
        PaintPicture ImgBIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "H" 'Border Bottom water
        PaintPicture ImgBIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "h" 'Border Bottom grass
        PaintPicture ImgBIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "U" 'Water
        PaintPicture ImgTIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "u" 'grass
        PaintPicture ImgIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "I" 'Water
        PaintPicture ImgTIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "i" 'grass
        PaintPicture ImgIB3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "J" 'Water
        PaintPicture ImgBIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "j" 'grass
        PaintPicture ImgIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "K" 'Water
        PaintPicture ImgBIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "k" 'grass
        PaintPicture ImgIT3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "," 'grass
        PaintPicture imgHouseExitB1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ";" 'grass
        PaintPicture imgHouseExitB2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "M" 'grass
        PaintPicture imgHouseExit.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "-" 'Water
        PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "_" 'Nothing
        PaintPicture imgNothing.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "Z" 'Wall bottom
        PaintPicture imgWallBottom.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "X" 'Wall top
        PaintPicture imgWallTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "C" 'Floor Blue 1
        PaintPicture imgFloor1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "z" 'Floor Tronco
        PaintPicture imgWall2.Picture, (X + 3) * 32, (Y + 3) * 32
            
      Case Is = "1" 'Laddy
        PaintPicture imgCharWomen1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "2" 'Boy 1
        PaintPicture imgCharBoy1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "3" 'Boy 2
        PaintPicture imgCharBoy2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "4" 'Good Wizard
        PaintPicture imgCharGoodWizard.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "5" 'Bad Wizard
        PaintPicture imgCharBadWizard.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "6" 'Lady
        PaintPicture imgCharWomen2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "7" 'Lady
        PaintPicture imgCharWomen3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "8" 'Soldier
        PaintPicture imgCharSoldier.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "9" 'King
        PaintPicture ImgKing.Picture, (X + 3) * 32, (Y + 3) * 32
      
      
      Case Is = "/"
        PaintPicture imgBlueTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "\"
        PaintPicture imgRedTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "*"
        PaintPicture imgWindow1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "#" '3
        'PaintPicture imgNothing.Picture, (X + 3) * 32, (Y + 3) * 32
        PaintPicture imgDoor1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "@" '3
        PaintPicture imgStarGate.Picture, (X + 3) * 32, (Y + 3) * 32
    
      Case Is = "!" 'Jar
        PaintPicture imgJar.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "(" 'Table
        PaintPicture imgTable1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ")" 'Table2
        PaintPicture imgTable2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "[" 'Bed
        PaintPicture imgBed1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "]" 'Bed2
        PaintPicture imgBed2.Picture, (X + 3) * 32, (Y + 3) * 32
    
    
      End Select
skip:
    Next
  Next
  

  'Character Movements.
  Select Case CharFacing
  Case Is = 1
    PaintPicture imgUpChar.Picture, 5 * 32, 5 * 32
  Case Is = 2
    PaintPicture imgDownChar.Picture, 5 * 32, 5 * 32
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * 32, 5 * 32
  Case Is = 4
    PaintPicture imgLeftChar.Picture, 5 * 32, 5 * 32
  Case Is = 5
    PaintPicture imgCarryChar.Picture, 5 * 32, 5 * 32
  End Select
  
  Select Case Item
  Case Is = 1
    PaintPicture imgItem1.Picture, 5 * 32, 4 * 32
    Item = 0
  Case Is = 2
    PaintPicture imgItem2.Picture, 5 * 32, 4 * 32
    Item = 0
  Case Is = 3
    PaintPicture imgItem3.Picture, 5 * 32, 4 * 32
    Item = 0
  End Select

End Sub

Private Sub Form_Resize()
  DrawIt
End Sub

Private Sub Form_Terminate()
  If bZoom Then LargeScreen (1)
End Sub
Private Sub LargeScreen(iSize As Integer)
'Code:
Dim message As String
Dim typDevM As typDevMODE
Dim lngResult As Long
Dim intAns    As Integer
ScreenWidth = 1024
ScreenHeight = 768
If iSize = 2 Then
   ScreenWidth = 640
   ScreenHeight = 480
End If
' Retrieve info about the current graphics mode
' on the current display device.
lngResult = EnumDisplaySettings(0, 0, typDevM)

' Set the new resolution. Don't change the color
' depth so a restart is not necessary.
With typDevM
    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    .dmPelsWidth = ScreenWidth  'ScreenWidth (640,800,1024, etc)
    .dmPelsHeight = ScreenHeight 'ScreenHeight (480,600,768, etc)
End With

' Change the display settings to the specified graphics mode.
lngResult = ChangeDisplaySettings(typDevM, CDS_TEST)
Select Case lngResult
    Case DISP_CHANGE_RESTART
        intAns = MsgBox("You must restart your computer to apply these changes." & _
            vbCrLf & vbCrLf & "Do you want to restart now?", _
            vbYesNo + vbSystemModal, "Screen Resolution")
        If intAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
    Case DISP_CHANGE_SUCCESSFUL
        Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
        message = MsgBox("Screen resolution changed", vbInformation, "Resolution Changed ")
        bZoom = True
    Case Else
        message = MsgBox("Mode not supported", vbSystemModal, "Error")
End Select
End Sub

Private Sub ImgBlue_Click(Index As Integer)
   Dim cMsg As String
   Dim strMe As String
   If ImgBlue(Index).BorderStyle = 1 Then
     ImgBlue(Index).BorderStyle = 0
   Else
     CheckMygive
     ImgBlue(Index).BorderStyle = 1
     Mygive = "imgblue" & Index
     iNeedQty = 1
     strMe = Mid(strInventory, Index + 1, 1)
     MyGiveLetter = strMe
     Select Case strMe
      Case Is = "c" 'Saw
        cMsg = "SAW"
      Case Is = "g" 'Bucket
       cMsg = "BUCKET"
      Case Is = "O" 'Lamp
        cMsg = "LAMP"
      Case Is = "o" 'GOLD
        cMsg = "GOLD"
      Case Is = "x" 'Purplegem
       cMsg = "PURPLEGEM"
      Case Is = "." 'Purplegem
       cMsg = "GREENGEM"
      Case Is = "&" 'Purplegem
       cMsg = "REDGEM"
      Case Is = "m" 'Apple
        cMsg = "APPLE"
      Case Is = "L" 'YKey
       cMsg = "YKEY"
      Case Is = "l" 'Bottle
       cMsg = "BOTTLE"
      Case Is = "~" 'RKey
        cMsg = "RKEY"
      Case Is = "N" 'Book
       cMsg = "BOOK"
      Case Is = "n" 'Gem
       cMsg = "BLUEGEM"
      Case Is = "=" 'Map
        cMsg = "MAP"
      Case Is = "%" 'BBottle
       cMsg = "BBOTTLE"
      Case Is = "^" 'YBottle
       cMsg = "YBOTTLE"
      Case Is = "{" 'Sword
       cMsg = "SWORD"
       Case Is = "}" 'Bow
       cMsg = "BOW"
      Case Is = "|" 'Armor
       cMsg = "ARMOR"
      Case Is = ":" 'Shield
       cMsg = "SHIELD"
       Case Is = "$" 'Candle
       cMsg = "CANDLE"
      End Select
     CheckNeed (cMsg)
   End If
End Sub

Private Sub ImgBomb2_Click()
 If Bomb = 0 Then Exit Sub
 If ImgBomb2.BorderStyle = 1 Then
    ImgBomb2.BorderStyle = 0
  Else
    CheckMygive
    ImgBomb2.BorderStyle = 1
    iNeedQty = Bomb2
    CheckNeed ("BOMB")
  End If
End Sub

Private Sub imgNCoin_Click()
 If Coin = 0 Then Exit Sub
 If imgNCoin.BorderStyle = 1 Then
    imgNCoin.BorderStyle = 0
   Else
    CheckMygive
    imgNCoin.BorderStyle = 1
    iNeedQty = Coin
    CheckNeed ("COIN")
   End If
End Sub

Private Sub imgNMagic_Click()
  If Magic = 0 Then Exit Sub
 If imgNMagic.BorderStyle = 1 Then
    imgNMagic.BorderStyle = 0
   Else
    CheckMygive
    imgNMagic.BorderStyle = 1
    iNeedQty = Magic
    CheckNeed ("MAGIC")
   End If
End Sub


Private Sub imgNTicket_Click()
 If Ticket = 0 Then Exit Sub
 If imgNTicket.BorderStyle = 1 Then
    imgNTicket.BorderStyle = 0
  Else
  CheckMygive
    imgNTicket.BorderStyle = 1
    Mygive = "imgnticket"
     iNeedQty = Ticket
     CheckNeed ("TICKET")
  End If
End Sub


Private Sub imgNToast_Click()
 If Toast = 0 Then Exit Sub
 If imgNToast.BorderStyle = 1 Then
    imgNToast.BorderStyle = 0
  Else
    CheckMygive
    imgNToast.BorderStyle = 1
      iNeedQty = Toast
     CheckNeed ("TOAST")
  End If
End Sub


Private Sub imgNWood_Click()
  If Wood = 0 Then Exit Sub
  If imgNWood.BorderStyle = 1 Then
    imgNWood.BorderStyle = 0
   Else
    CheckMygive
    imgNWood.BorderStyle = 1
     iNeedQty = Wood
    CheckNeed ("WOOD")
   End If
End Sub


Private Sub TimBunny_Timer()
  Dim i As Integer
  Dim j As Integer
  If iBunny > 0 Then
      For j = 1 To iBunny
         If iBunnyloc(j, 3) > 0 Then
            If iBunnyloc(j, 3) = 1 Then
           'right
                If Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2) + 1, 1) = "G" Then
                  Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2) + 1, 1) = " "
                  Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2), 1) = "G"
                  iBunnyloc(j, 2) = iBunnyloc(j, 2) + 1
                Else
                 iBunnyloc(j, 3) = 2
                End If
            End If
            If iBunnyloc(j, 3) = 2 Then
               If Mid(Map(iBunnyloc(j, 1) - 1), iBunnyloc(j, 2), 1) = "G" Then
                 '
                 'up
                 Mid(Map(iBunnyloc(j, 1) - 1), iBunnyloc(j, 2), 1) = " "
                 Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2), 1) = "G"
                 iBunnyloc(j, 1) = iBunnyloc(j, 1) - 1
               Else
                  iBunnyloc(j, 3) = 3
               End If
            End If
            If iBunnyloc(j, 3) = 3 Then
             'left
               If Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2) - 1, 1) = "G" Then
                  Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2) - 1, 1) = " "
                  Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2), 1) = "G"
                   iBunnyloc(j, 2) = iBunnyloc(j, 2) - 1
               Else
                  iBunnyloc(j, 3) = 4
               End If
           End If
            If iBunnyloc(j, 3) = 4 Then
           'down
             If Mid(Map(iBunnyloc(j, 1) + 1), iBunnyloc(j, 2), 1) = "G" Then
                Mid(Map(iBunnyloc(j, 1) + 1), iBunnyloc(j, 2), 1) = " "
                Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2), 1) = "G"
                iBunnyloc(j, 1) = iBunnyloc(j, 1) + 1
                
               Else
                     iBunnyloc(j, 3) = 1
                     'right
                    If Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2) + 1, 1) = "G" Then
                      Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2) + 1, 1) = " "
                      Mid(Map(iBunnyloc(j, 1)), iBunnyloc(j, 2), 1) = "G"
                      iBunnyloc(j, 2) = iBunnyloc(j, 2) + 1
                    Else
                     iBunnyloc(j, 3) = 2
                    End If
            
              End If
        End If
        End If
       Next
     DrawIt
 End If
End Sub

Private Sub Timer1_Timer()
   DrawIt
   Timer1.Enabled = False
   ViewQuests
End Sub

Private Sub Timer2_Timer()
  DrawIt
  Timer2.Enabled = False
  If bBunnyTimer And TimBunny.Enabled = False Then TimBunny.Enabled = True
End Sub


