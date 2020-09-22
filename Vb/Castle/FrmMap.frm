VERSION 5.00
Begin VB.Form FrmMap 
   Caption         =   "MakeMap"
   ClientHeight    =   9855
   ClientLeft      =   2100
   ClientTop       =   900
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   657
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   Begin VB.CommandButton Command5 
      Caption         =   "Grass"
      Height          =   375
      Left            =   9960
      TabIndex        =   16
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox TxtGive 
      Height          =   495
      Left            =   9960
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton ChrX 
      Caption         =   "X"
      Height          =   495
      Left            =   10320
      TabIndex        =   14
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton Cmdy 
      Caption         =   "Y"
      Height          =   495
      Left            =   10320
      TabIndex        =   13
      Top             =   9000
      Width           =   735
   End
   Begin VB.TextBox TxtId 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Black"
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Move"
      Height          =   495
      Left            =   9720
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   735
      Left            =   9600
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox TxtMap 
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Text            =   "New"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TxtY 
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Text            =   "3"
      Top             =   8880
      Width           =   855
   End
   Begin VB.TextBox TxtX 
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Text            =   "3"
      Top             =   8880
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   9240
      Width           =   735
   End
   Begin VB.CommandButton CmdBlank 
      Caption         =   "Water"
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read"
      Height          =   735
      Left            =   9600
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5640
      Top             =   5400
   End
   Begin VB.Image ImgWGrass 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMap.frx":0000
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image Imgbun 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMap.frx":0C42
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image ImgJDoor 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMap.frx":1884
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image ImgSDoor 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMap.frx":24C6
      Top             =   6960
      Width           =   480
   End
   Begin VB.Image ImgGGem 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMap.frx":3108
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image ImgRGem 
      Height          =   480
      Left            =   960
      Picture         =   "FrmMap.frx":3D4A
      Top             =   9000
      Width           =   480
   End
   Begin VB.Label Label8 
      Caption         =   "---Spells---"
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
      Left            =   7320
      TabIndex        =   22
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Cut - Weed"
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
      Left            =   7320
      TabIndex        =   21
      Top             =   7740
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Light - Darkness"
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
      Left            =   7320
      TabIndex        =   20
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Fill - Swamp  "
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
      Left            =   7320
      TabIndex        =   19
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Image ImgSwamp 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMap.frx":498C
      Top             =   6960
      Width           =   480
   End
   Begin VB.Label Label4 
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
      Left            =   9480
      TabIndex        =   18
      Top             =   8520
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "X"
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
      Left            =   7560
      TabIndex        =   17
      Top             =   8520
      Width           =   195
   End
   Begin VB.Image ImgCandle 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMap.frx":55CE
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image ImgKing 
      Height          =   480
      Left            =   9360
      Picture         =   "FrmMap.frx":6210
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image ImgArmor 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMap.frx":6E52
      Top             =   9240
      Width           =   480
   End
   Begin VB.Image imgBrick2 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMap.frx":7A94
      Top             =   8760
      Width           =   480
   End
   Begin VB.Image imgShield 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMap.frx":86D6
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image ImgBow 
      Height          =   480
      Left            =   10560
      Picture         =   "FrmMap.frx":9318
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image ImgSword 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMap.frx":9F5A
      Top             =   9240
      Width           =   480
   End
   Begin VB.Image ImgBrick 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMap.frx":AB9C
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image imgBomb 
      Height          =   480
      Left            =   10680
      Picture         =   "FrmMap.frx":B7DE
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMap 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMap.frx":C420
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image ImgBBottle 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMap.frx":D062
      Top             =   7920
      Width           =   480
   End
   Begin VB.Image ImgYBottle 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMap.frx":DCA4
      Top             =   7920
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Destroy - Rock"
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
      Left            =   7320
      TabIndex        =   12
      Top             =   7140
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Axe - Tree"
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
      Left            =   7320
      TabIndex        =   11
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image imgItem6 
      Height          =   480
      Left            =   9960
      Picture         =   "FrmMap.frx":E8E6
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image ImgMagic 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMap.frx":F528
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image ImgBook 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMap.frx":1016A
      Top             =   7800
      Width           =   480
   End
   Begin VB.Image ImgGem 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMap.frx":1060F
      Top             =   9480
      Width           =   480
   End
   Begin VB.Image ImgSaw 
      Height          =   480
      Left            =   9960
      Picture         =   "FrmMap.frx":11251
      Top             =   7200
      Width           =   480
   End
   Begin VB.Image ImgBucket 
      Height          =   480
      Left            =   9960
      Picture         =   "FrmMap.frx":11E93
      Top             =   7800
      Width           =   480
   End
   Begin VB.Image ImgApple 
      Height          =   480
      Left            =   8040
      Picture         =   "FrmMap.frx":12AD5
      Top             =   8640
      Width           =   480
   End
   Begin VB.Image ImgLamp 
      Height          =   480
      Left            =   8040
      Picture         =   "FrmMap.frx":13717
      Top             =   9120
      Width           =   480
   End
   Begin VB.Image ImgGold 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMap.frx":14359
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image ImgPGem 
      Height          =   480
      Left            =   960
      Picture         =   "FrmMap.frx":14F9B
      Top             =   9480
      Width           =   480
   End
   Begin VB.Image ImgKey2 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMap.frx":15BDD
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   9600
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Image imgSign 
      Height          =   480
      Left            =   960
      Picture         =   "FrmMap.frx":1681F
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image imgGrass 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMap.frx":17461
      Top             =   7560
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
      Left            =   960
      Picture         =   "FrmMap.frx":180A3
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image ImgBlack 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMap.frx":18CE5
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image ImgBIR 
      Height          =   480
      Left            =   2400
      Picture         =   "FrmMap.frx":19927
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image ImgTIL 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMap.frx":1A569
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image ImgBIL 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMap.frx":1B1AB
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image imgTOL 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMap.frx":1BDED
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image imgTOR 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMap.frx":1CA2F
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image ImgIT 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMap.frx":1D671
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image ImgIL 
      Height          =   480
      Left            =   2400
      Picture         =   "FrmMap.frx":1E2B3
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image ImgIR 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMap.frx":1EEF5
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image ImgTIR 
      Height          =   480
      Left            =   2400
      Picture         =   "FrmMap.frx":1FB37
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image imgBOR 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMap.frx":20779
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image imgBOL 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMap.frx":213BB
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image ImgIB 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMap.frx":21FFD
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image imgUpChar 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmMap.frx":22C3F
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLeftChar 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmMap.frx":23881
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRightChar 
      Height          =   480
      Left            =   5640
      Picture         =   "FrmMap.frx":244C3
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDownChar 
      Height          =   480
      Left            =   5640
      Picture         =   "FrmMap.frx":25105
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy1 
      Height          =   480
      Left            =   8880
      Picture         =   "FrmMap.frx":25D47
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image imgCharWomen1 
      Height          =   480
      Left            =   8880
      Picture         =   "FrmMap.frx":26989
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image imgWindow1 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMap.frx":275CB
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image imgBlueTop 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMap.frx":2820D
      Top             =   7080
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
      Left            =   960
      Picture         =   "FrmMap.frx":28E4F
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image imgDoor1 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMap.frx":29A91
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image imgRedTop 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMap.frx":2A6D3
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image ImgIB2 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMap.frx":2B315
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image ImgTIR2 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMap.frx":2BF57
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image ImgIR2 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMap.frx":2CB99
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image ImgIL2 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMap.frx":2D7DB
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image ImgIT2 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMap.frx":2E41D
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image ImgBIL2 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMap.frx":2F05F
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image ImgTIL2 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMap.frx":2FCA1
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image ImgBIR2 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMap.frx":308E3
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image imgWallBottom 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMap.frx":31525
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image imgWallTop 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMap.frx":32167
      Top             =   6960
      Width           =   480
   End
   Begin VB.Image imgFloor1 
      Height          =   480
      Left            =   960
      Picture         =   "FrmMap.frx":32DA9
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgBOL2 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMap.frx":339EB
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgBOR2 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMap.frx":3462D
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgTOR2 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMap.frx":3526F
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image imgTOL2 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMap.frx":35EB1
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image imgRockHill 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMap.frx":36AF3
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgTrees 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMap.frx":37735
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgCharBoy2 
      Height          =   480
      Left            =   8880
      Picture         =   "FrmMap.frx":38377
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image imgCharGoodWizard 
      Height          =   480
      Left            =   8880
      Picture         =   "FrmMap.frx":38FB9
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image imgCharBadWizard 
      Height          =   480
      Left            =   8520
      Picture         =   "FrmMap.frx":39BFB
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgWeed 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMap.frx":3A83D
      Top             =   8760
      Width           =   480
   End
   Begin VB.Image imgWall2 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMap.frx":3B47F
      Top             =   7920
      Width           =   480
   End
   Begin VB.Image ImgIB3 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMap.frx":3C0C1
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image ImgTIR3 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMap.frx":3CD03
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image ImgIR3 
      Height          =   495
      Left            =   2880
      Picture         =   "FrmMap.frx":3D945
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image ImgIL3 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMap.frx":3E5E7
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image ImgIT3 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMap.frx":3F229
      Top             =   9480
      Width           =   480
   End
   Begin VB.Image ImgBIL3 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMap.frx":3FE6B
      Top             =   9480
      Width           =   480
   End
   Begin VB.Image ImgTIL3 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMap.frx":40AAD
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image ImgBIR3 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMap.frx":416EF
      Top             =   9480
      Width           =   480
   End
   Begin VB.Image imgHouseExit 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMap.frx":42331
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image imgJar 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMap.frx":42F73
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image imgHouseExitB1 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMap.frx":43BB5
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image imgHouseExitB2 
      Height          =   480
      Left            =   2400
      Picture         =   "FrmMap.frx":447F7
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image imgTable1 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMap.frx":45439
      Top             =   8760
      Width           =   480
   End
   Begin VB.Image imgTable2 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMap.frx":4607B
      Top             =   8760
      Width           =   480
   End
   Begin VB.Image imgBed1 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMap.frx":46CBD
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image imgBed2 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMap.frx":478FF
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image imgCharWomen2 
      Height          =   480
      Left            =   9360
      Picture         =   "FrmMap.frx":48541
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image imgCharWomen3 
      Height          =   480
      Left            =   9360
      Picture         =   "FrmMap.frx":49183
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image imgItem1 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMap.frx":49DC5
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image imgCarryChar 
      Height          =   480
      Left            =   6120
      Picture         =   "FrmMap.frx":4AA07
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem2 
      Height          =   480
      Left            =   4200
      Picture         =   "FrmMap.frx":4B649
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem3 
      Height          =   480
      Left            =   4680
      Picture         =   "FrmMap.frx":4C28B
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharSoldier 
      Height          =   480
      Left            =   9360
      Picture         =   "FrmMap.frx":4CECD
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image imgStarGate 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMap.frx":4DB0F
      Top             =   9000
      Width           =   480
   End
   Begin VB.Image imgItem4 
      Height          =   480
      Left            =   3240
      Picture         =   "FrmMap.frx":4E751
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem5 
      Height          =   480
      Left            =   3720
      Picture         =   "FrmMap.frx":4F393
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FrmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strMapChr As String
Public Sub PictureHandler()
  PassToNext = 1
  PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
End Sub
Public Sub DrawIt()
  For Y = -3 To 10
    For X = -3 To 14
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
      Case Is = "?" 'Grass
        PaintPicture ImgStairs.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "=" 'Map
        PaintPicture ImgMap.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "%" 'BBottle
        PaintPicture ImgBBottle.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "^" 'YBottle
        PaintPicture ImgYBottle.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "c" 'Bucket
        PaintPicture ImgSaw.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "g" 'Saw
        PaintPicture ImgBucket.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "O" 'Lamp
        PaintPicture ImgLamp.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "o" 'Lamp
        PaintPicture ImgGold.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "x" 'Lamp
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
      Case Is = "n" 'Key
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
      Case Is = "6" 'Laddy
        PaintPicture imgCharWomen2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "7" 'Laddy
        PaintPicture imgCharWomen3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "8" 'Soldier
        PaintPicture imgCharSoldier.Picture, (X + 3) * 32, (Y + 3) * 32
      
      
      Case Is = "/"
        PaintPicture imgBlueTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "\"
        PaintPicture imgRedTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "*"
        PaintPicture imgWindow1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "#" '3
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
Public Sub DrawIt2()
  For Y = -3 To 14
    For X = -3 To 20
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
      Case Is = "." 'green gem
        PaintPicture ImgGGem.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = Chr(34) 'Wgrass
        PaintPicture ImgWGrass.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "'" 'green gem
        PaintPicture ImgJDoor.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = " " 'bunny gem
        PaintPicture Imgbun.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "`" 'green gem
        PaintPicture ImgSDoor.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "&" 'red gem
        PaintPicture ImgRGem.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "$" 'Candle
        PaintPicture ImgCandle.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "+" 'Swamp
        PaintPicture ImgSwamp.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "}" 'Bow
        PaintPicture ImgBow.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "|" 'Armor
        PaintPicture ImgArmor.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = ":" 'Shield
        PaintPicture imgShield.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "P" 'Grass
        PaintPicture ImgBrick.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "p" 'Grass
        PaintPicture imgBrick2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "{" 'Sword
        PaintPicture ImgSword.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "=" 'Map
        PaintPicture ImgMap.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "%" 'BBottle
        PaintPicture ImgBBottle.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "^" 'YBottle
        PaintPicture ImgYBottle.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "c" 'Saw
         PaintPicture ImgSaw.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "g" 'Saw
        PaintPicture ImgBucket.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "O" 'Lamp
        PaintPicture ImgLamp.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "o" 'Lamp
        PaintPicture ImgGold.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "x" 'Lamp
        PaintPicture ImgPGem.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "m" 'Saw
        PaintPicture ImgApple.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "L" 'Key
        PaintPicture imgItem6.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "l" 'Key
        PaintPicture ImgMagic.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "~" 'Key
        PaintPicture ImgKey2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "N" 'Key
        PaintPicture ImgBook.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "n" 'BLUE GEM
        PaintPicture ImgGem.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "b" '2 bush
        PaintPicture imgTrees.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "<" 'Weed
        PaintPicture imgWeed.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = ">" 'Rock Hill
        PaintPicture imgRockHill.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "0" 'Cero Sign
        PaintPicture imgSign.Picture, (X + 3) * 24, (Y + 3) * 24
      
      Case Is = "Q" 'Water
        PaintPicture imgTOL.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "q" 'grass
        PaintPicture imgTOL2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "A" 'Water
        PaintPicture imgBOL.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "a" 'grass
        PaintPicture imgBOL2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "W" 'Water
        PaintPicture imgTOR.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "w" 'grass
        PaintPicture imgTOR2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "S" 'Water
        PaintPicture imgBOR.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "s" 'grass
        PaintPicture imgBOR2.Picture, (X + 3) * 24, (Y + 3) * 24
      
      Case Is = "E" 'Border Left water
        PaintPicture ImgIL.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "e" 'Border Left grass
        PaintPicture ImgIL2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "R" 'Border Right water
        PaintPicture ImgIR.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "r" 'Border Right grass
        PaintPicture ImgIR2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "D" 'Border Top water
        PaintPicture ImgIT.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "d" 'Border Top grass
        PaintPicture ImgIT2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "F" 'Border Bottom water
        PaintPicture ImgIB.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "f" 'Border Bottom grass
        PaintPicture ImgIB2.Picture, (X + 3) * 24, (Y + 3) * 24
      
      Case Is = "T" 'Border Bottom water
        PaintPicture ImgTIL.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "t" 'Border Bottom grass
        PaintPicture ImgTIL2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "Y" 'Border Bottom water
        PaintPicture ImgTIR.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "y" 'Border Bottom grass
        PaintPicture ImgTIR2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "V" 'Border Bottom water
        PaintPicture ImgBIL.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "v" 'Border Bottom grass
        PaintPicture ImgBIL2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "H" 'Border Bottom water
        PaintPicture ImgBIR.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "h" 'Border Bottom grass
        PaintPicture ImgBIR2.Picture, (X + 3) * 24, (Y + 3) * 24
      
      Case Is = "U" 'Water
        PaintPicture ImgTIL3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "u" 'grass
        PaintPicture ImgIR3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "I" 'Water
        PaintPicture ImgTIR3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "i" 'grass
        PaintPicture ImgIB3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "J" 'Water
        PaintPicture ImgBIL3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "j" 'grass
        PaintPicture ImgIL3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "K" 'Water
        PaintPicture ImgBIR3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "k" 'grass
        PaintPicture ImgIT3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "," 'grass
        PaintPicture imgHouseExitB1.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = ";" 'grass
        PaintPicture imgHouseExitB2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "M" 'grass
        PaintPicture imgHouseExit.Picture, (X + 3) * 24, (Y + 3) * 24
      
      Case Is = "-" 'Water
        PaintPicture ImgBlack.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "_" 'Nothing
        PaintPicture imgNothing.Picture, (X + 3) * 24, (Y + 3) * 24
      
      Case Is = "Z" 'Wall bottom
        PaintPicture imgWallBottom.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "X" 'Wall top
        PaintPicture imgWallTop.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "C" 'Floor Blue 1
        PaintPicture imgFloor1.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "z" 'Floor Tronco
        PaintPicture imgWall2.Picture, (X + 3) * 24, (Y + 3) * 24
            
      Case Is = "1" 'Laddy
        PaintPicture imgCharWomen1.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "2" 'Boy 1
        PaintPicture imgCharBoy1.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "3" 'Boy 2
        PaintPicture imgCharBoy2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "4" 'Good Wizard
        PaintPicture imgCharGoodWizard.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "5" 'Bad Wizard
        PaintPicture imgCharBadWizard.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "6" 'Laddy
        PaintPicture imgCharWomen2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "7" 'Laddy
        PaintPicture imgCharWomen3.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "8" 'Soldier
        PaintPicture imgCharSoldier.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "9" 'KING
        PaintPicture ImgKing.Picture, (X + 3) * 24, (Y + 3) * 24
      
      Case Is = "/"
        PaintPicture imgBlueTop.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "\"
        PaintPicture imgRedTop.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "*"
        PaintPicture imgWindow1.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "#" '3
        PaintPicture imgDoor1.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "@" '3
        PaintPicture imgStarGate.Picture, (X + 3) * 24, (Y + 3) * 24
    
      Case Is = "!" 'Jar
        PaintPicture imgJar.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "(" 'Table
        PaintPicture imgTable1.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = ")" 'Table2
        PaintPicture imgTable2.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "[" 'Bed
        PaintPicture imgBed1.Picture, (X + 3) * 24, (Y + 3) * 24
      Case Is = "]" 'Bed2
        PaintPicture imgBed2.Picture, (X + 3) * 24, (Y + 3) * 24
    
    
      End Select
skip:
    Next
  Next
  
'
'  'Character Movements.
'  Select Case CharFacing
'  Case Is = 1
'    PaintPicture imgUpChar.Picture, 5 * 24, 5 * 24
'  Case Is = 2
'    PaintPicture imgDownChar.Picture, 5 * 24, 5 * 24
'  Case Is = 3
'    PaintPicture imgRightChar.Picture, 5 * 24, 5 * 24
'  Case Is = 4
'    PaintPicture imgLeftChar.Picture, 5 * 24, 5 * 24
'  Case Is = 5
'    PaintPicture imgCarryChar.Picture, 5 * 24, 5 * 24
'  End Select
'
'  Select Case Item
'  Case Is = 1
'    PaintPicture imgItem1.Picture, 5 * 24, 4 * 24
'    Item = 0
'  Case Is = 2
'    PaintPicture imgItem2.Picture, 5 * 24, 4 * 24
'    Item = 0
'  Case Is = 3
'    PaintPicture imgItem3.Picture, 5 * 24, 4 * 24
'    Item = 0
'  End Select

End Sub

Private Sub ChrX_Click()
  TxtX.Text = TxtX.Text + 12
    CharX = 2 + TxtX
   CharY = 2 + TxtY
   CharFacing = 1
  DrawIt2
End Sub

Private Sub CmdBlank_Click()
   Map(0) = "-------------------------------------------------------------------------------"
   Map(0) = Map(0) & "-------------------------------------------------------------------------------"
   
   Dim i As Integer
    For i = 1 To 90
       Map(i) = Map(0)
   Next
   CharX = 5
   CharY = 5
   CharFacing = 1
  DrawIt
End Sub

Private Sub Cmdy_Click()
TxtY.Text = TxtY.Text + 10
    CharX = 2 + TxtX
   CharY = 2 + TxtY
   CharFacing = 1
  DrawIt2
End Sub

Private Sub Command1_Click()
   Dim i  As Integer
   On Error GoTo myexit
   Dim cFile
   Dim strMap
   cFile = "c:\Kids\Quest\" & Trim(TxtMap.Text) & ".txt"
    Open cFile For Input As #1 ' Open file for input.
    i = 0
   Do While Not EOF(1) ' Loop until end of file.
   Line Input #1, strMap
   If i < 99 Then
     Map(i) = strMap
     i = i + 1
    End If
    Loop
   Close #1    ' Close file.
   CharX = 2 + TxtX
   CharY = 2 + TxtY
   CharFacing = 1
  DrawIt2
myexit:
'TxtMap.Text = Error(Err.Number)
 
End Sub

Private Sub Command2_Click()
  SaveMapFile (TxtMap.Text)
End Sub

Private Sub Command3_Click()
   CharX = 2 + TxtX
   CharY = 2 + TxtY
   CharFacing = 1
  DrawIt2
End Sub

Private Sub Command4_Click()
  Map(0) = "____________________________________________________________________"
   Map(0) = Map(0) & "____________________________________________________________________"
   
   Dim i As Integer
    For i = 1 To 90
       Map(i) = Map(0)
   Next
   CharX = 5
   CharY = 5
   CharFacing = 1
  DrawIt2
End Sub

Private Sub Command5_Click()
    Map(0) = "-------------------------------------------------------------------------------"
   Map(0) = Map(0) & "-------------------------------------------------------------------------------"
     
   Dim i As Integer
    For i = 1 To 90
       Map(i) = Map(0)
   Next
      Map(8) = "-------GGGGGGGGGGGGGGGGGGGGGGGGGGGGGG------------------------------"
    For i = 8 To 16
     Map(i) = Map(8)
   Next
   CharX = 5
   CharY = 5
   CharFacing = 1
  DrawIt2
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Integer
   Dim j As Integer
   i = TxtX + Int(X / 24)
   j = TxtY + Int(Y / 24)
   Text1.Text = i
   Text2.Text = j
   strMap = Mid(Map(j), 1, i - 1) & strMapChr & Mid(Map(j), i + 1, Len(Map(j)) - i)
   Map(j) = strMap
   DrawIt2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim i As Integer
   Dim j As Integer
   i = TxtX + Int(X / 24)
   j = TxtY + Int(Y / 24)
   Text1.Text = i - 3
   Text2.Text = j - 3
End Sub

Private Sub Image2_Click()
  strMapChr = " "
   Image1.Picture = imgItem6.Picture
   TxtId.Text = strMapChr
   
End Sub

Private Sub Imbun_Click()
   strMapChr = " "
   Image1.Picture = imgItem6.Picture
   TxtId.Text = strMapChr
End Sub

Private Sub ImgApple_Click()
  strMapChr = "m"
  Image1.Picture = ImgApple.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Apple"
End Sub

Private Sub ImgArmor_Click()
   strMapChr = "|"
  Image1.Picture = ImgArmor.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Armor"
End Sub

Private Sub ImgBBottle_Click()
  strMapChr = "%"
  Image1.Picture = ImgBBottle.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "BBottle"
End Sub

Private Sub imgBed1_Click()
  strMapChr = "["
   Image1.Picture = imgBed1.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgBed2_Click()
 strMapChr = "]"
   Image1.Picture = imgBed2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgBIL_Click()
  strMapChr = "V"
   Image1.Picture = ImgBIL.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgBIL2_Click()
   strMapChr = "v"
   Image1.Picture = ImgBIL2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgBIL3_Click()
   strMapChr = "J"
  Image1.Picture = ImgBIL3.Picture
   TxtId.Text = "0"
   TxtGive.Text = ""
End Sub

Private Sub ImgBIR_Click()
   strMapChr = "H"
   Image1.Picture = ImgBIR.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgBIR2_Click()
   strMapChr = "h"
   Image1.Picture = ImgBIR2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgBIR3_Click()
  strMapChr = "K"
   Image1.Picture = ImgBIR3.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgBlack_Click()
   strMapChr = "-"
   Image1.Picture = ImgBlack.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgBlueTop_Click()
  strMapChr = "/"
   Image1.Picture = imgBlueTop.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgBOL_Click()
   strMapChr = "A"
   Image1.Picture = imgBOL.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgBOL2_Click()
    strMapChr = "a"
   Image1.Picture = imgBOL2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgBook_Click()
  strMapChr = "N"
  Image1.Picture = ImgBook.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Book"
End Sub

Private Sub imgBOR_Click()
   strMapChr = "S"
  Image1.Picture = imgBOR.Picture
   TxtId.Text = "0"
   TxtGive.Text = ""
End Sub

Private Sub imgBOR2_Click()
  strMapChr = "s"
   Image1.Picture = imgBOR2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgBow_Click()
   strMapChr = "}"
   Image1.Picture = ImgBow.Picture
    TxtId.Text = "}"
    TxtGive.Text = "Bow"
End Sub

Private Sub ImgBrick_Click()
   strMapChr = "P"
   Image1.Picture = ImgBrick.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgBrick2_Click()
   strMapChr = "p"
   Image1.Picture = imgBrick2.Picture
   TxtId.Text = "0"
   TxtGive.Text = ""
End Sub

Private Sub ImgBucket_Click()
 strMapChr = "g"
  Image1.Picture = ImgBucket.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Bucket"
End Sub

Private Sub Imgbun_Click()
   strMapChr = " "
   Image1.Picture = Imgbun.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = ""
End Sub

Private Sub imgBush_Click()
   strMapChr = "B"
   Image1.Picture = imgBush.Picture
    TxtId.Text = "B"
    TxtGive.Text = "Tree"
End Sub

Private Sub ImgCandle_Click()
   strMapChr = "$"
  Image1.Picture = ImgCandle.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Candle"
End Sub

Private Sub imgCharBadWizard_Click()
   strMapChr = "5"
   Image1.Picture = imgCharBadWizard.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgCharBoy1_Click()
   strMapChr = "2"
   Image1.Picture = imgCharBoy1.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgCharBoy2_Click()
 strMapChr = "3"
   Image1.Picture = imgCharBoy2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgCharGoodWizard_Click()
   strMapChr = "4"
   Image1.Picture = imgCharGoodWizard.Picture
    TxtId.Text = "0"
    TxtGive.Text = "Wizard"
End Sub

Private Sub imgCharSoldier_Click()
  strMapChr = "8"
   Image1.Picture = imgCharSoldier.Picture
    TxtId.Text = "0"
    TxtGive.Text = "Soldier"
End Sub

Private Sub imgCharWomen1_Click()
   strMapChr = "1"
   Image1.Picture = imgCharWomen1.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgCharWomen2_Click()
  strMapChr = "6"
   Image1.Picture = imgCharWomen2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgCharWomen3_Click()
  strMapChr = "7"
   Image1.Picture = imgCharWomen3.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgDoor1_Click()
   strMapChr = "#"
   Image1.Picture = imgDoor1.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgFloor1_Click()
    strMapChr = "C"
   Image1.Picture = imgFloor1.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgGem_Click()
   strMapChr = "n"
  Image1.Picture = ImgGem.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "BlueGem"
End Sub

Private Sub ImgGGem_Click()
   strMapChr = "."
   Image1.Picture = ImgGGem.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "GreenGem"
End Sub

Private Sub ImgGold_Click()
 strMapChr = "o"
  Image1.Picture = ImgGold.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Gold"
End Sub

Private Sub imgGrass_Click()
  strMapChr = "G"
   Image1.Picture = imgGrass.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgHouseExit_Click()
   strMapChr = "M"
   Image1.Picture = imgHouseExit.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgHouseExitB1_Click()
  strMapChr = ","
   Image1.Picture = imgHouseExitB1.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgHouseExitB2_Click()
   strMapChr = ";"
   Image1.Picture = imgHouseExitB2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIB_Click()
   strMapChr = "F"
   Image1.Picture = ImgIB.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIB2_Click()
   strMapChr = "f"
   Image1.Picture = ImgIB2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIB3_Click()
  strMapChr = "i"
   Image1.Picture = ImgIB3.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIL_Click()
    strMapChr = "E"
   Image1.Picture = ImgIL.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIL2_Click()
    strMapChr = "e"
   Image1.Picture = ImgIL2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIL3_Click()
  strMapChr = "j"
   Image1.Picture = ImgIL3.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIR_Click()
   strMapChr = "R"
   Image1.Picture = ImgIR.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIR2_Click()
    strMapChr = "r"
   Image1.Picture = ImgIR2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIR3_Click()
  strMapChr = "u"
  Image1.Picture = ImgIR3.Picture
  TxtId.Text = "0"
  TxtGive.Text = ""
End Sub

Private Sub ImgIT_Click()
  strMapChr = "D"
   Image1.Picture = ImgIT.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgIT2_Click()
  strMapChr = "d"
   Image1.Picture = ImgIT2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""

End Sub

Private Sub ImgIT3_Click()
   strMapChr = "k"
   Image1.Picture = ImgIT3.Picture
   TxtId.Text = "0"
   TxtGive.Text = ""
End Sub

Private Sub imgItem1_Click()
 TxtId.Text = "1"
 TxtGive.Text = "Wood"
End Sub

Private Sub imgItem2_Click()
 TxtId.Text = "2"
 TxtGive.Text = "Coin"
End Sub

Private Sub imgItem3_Click()
 TxtId.Text = "3"
 TxtGive.Text = "Magic"
End Sub

Private Sub imgItem4_Click()
 TxtId.Text = "4"
 TxtGive.Text = "Ticket"
End Sub

Private Sub imgItem5_Click()
  TxtId.Text = "5"
  TxtGive.Text = "Toast"
End Sub

Private Sub imgItem6_Click()
    strMapChr = "L"
   Image1.Picture = imgItem6.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "YKey"
End Sub

Private Sub imgJar_Click()
   strMapChr = "!"
   Image1.Picture = imgJar.Picture
   TxtId.Text = "0"
   TxtGive.Text = ""
End Sub

Private Sub ImgJDoor_Click()
   strMapChr = "'"
   Image1.Picture = ImgJDoor.Picture
   TxtId.Text = "'"
   TxtGive.Text = ""
End Sub

Private Sub ImgKey2_Click()
 strMapChr = "~"
  Image1.Picture = ImgKey2.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "RKey"
End Sub

Private Sub ImgKing_Click()
  strMapChr = "9"
   Image1.Picture = ImgKing.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgLamp_Click()
  strMapChr = "O"
  Image1.Picture = ImgLamp.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Lamp"
End Sub

Private Sub ImgMagic_Click()
  strMapChr = "l"
  Image1.Picture = ImgMagic.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Bottle"
End Sub

Private Sub ImgMap_Click()
  strMapChr = "="
  Image1.Picture = ImgMap.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Map"
End Sub

Private Sub imgNothing_Click()
    strMapChr = "_"
   Image1.Picture = imgNothing.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgPGem_Click()
  strMapChr = "x"
  Image1.Picture = ImgPGem.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "PurpleGem"

End Sub

Private Sub imgRedTop_Click()
   strMapChr = "\"
   Image1.Picture = imgRedTop.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgRGem_Click()
  strMapChr = "&"
  Image1.Picture = ImgRGem.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "RedGem"
End Sub

Private Sub imgRockHill_Click()
  strMapChr = ">"
   Image1.Picture = imgRockHill.Picture
    TxtId.Text = "7"
    TxtGive.Text = "Rock"
End Sub

Private Sub ImgSaw_Click()
   strMapChr = "c"
  Image1.Picture = ImgSaw.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Saw"
End Sub

Private Sub ImgSDoor_Click()
  strMapChr = "`"
  Image1.Picture = ImgSDoor.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = ""
End Sub

Private Sub imgShield_Click()
  strMapChr = ":"
  Image1.Picture = imgShield.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "Shield"
End Sub

Private Sub imgSign_Click()
   strMapChr = "0"
  Image1.Picture = imgSign.Picture
   TxtId.Text = "0"
   TxtGive.Text = "Sign"
End Sub

Private Sub imgStarGate_Click()
  strMapChr = "@"
   Image1.Picture = imgStarGate.Picture
    TxtId.Text = "0"
    TxtGive.Text = "StarGate"
End Sub

Private Sub ImgSwamp_Click()
    strMapChr = "+"
   Image1.Picture = ImgSwamp.Picture
   TxtId.Text = "+"
    TxtGive.Text = "Swamp"
End Sub

Private Sub ImgSword_Click()
   strMapChr = "{"
   Image1.Picture = ImgSword.Picture
   TxtId.Text = "{"
    TxtGive.Text = "Sword"
End Sub

Private Sub imgTable1_Click()
  strMapChr = "("
   Image1.Picture = imgTable1.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgTable2_Click()
strMapChr = ")"
   Image1.Picture = imgTable2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgTIL_Click()
    strMapChr = "T"
   Image1.Picture = ImgTIL.Picture
    TxtId.Text = "T"
    TxtGive.Text = ""
End Sub

Private Sub ImgTIL2_Click()
    strMapChr = "t"
   Image1.Picture = ImgTIL2.Picture
    TxtId.Text = "t"
    TxtGive.Text = ""
End Sub

Private Sub ImgTIL3_Click()
  strMapChr = "U"
   Image1.Picture = ImgTIL3.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgTIR_Click()
   strMapChr = "Y"
   Image1.Picture = ImgTIR.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgTIR2_Click()
   strMapChr = "y"
   Image1.Picture = ImgTIR2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgTIR3_Click()
   strMapChr = "I"
   Image1.Picture = ImgTIR3.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgTOL_Click()
   strMapChr = "Q"
   Image1.Picture = imgTOL.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgTOL2_Click()
   strMapChr = "q"
   Image1.Picture = imgTOL2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgTOR_Click()
  strMapChr = "W"
  Image1.Picture = imgTOR.Picture
   TxtId.Text = "0"
   TxtGive.Text = ""
End Sub

Private Sub imgTOR2_Click()
    strMapChr = "w"
   Image1.Picture = imgTOR2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgTrees_Click()
    strMapChr = "b"
   Image1.Picture = imgTrees.Picture
    TxtId.Text = "b"
    TxtGive.Text = "Bush"
    TxtGive.Text = ""
End Sub

Private Sub imgWall2_Click()
   strMapChr = "z"
   Image1.Picture = imgWall2.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgWallBottom_Click()
   strMapChr = "Z"
   Image1.Picture = imgWallBottom.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgWallTop_Click()
    strMapChr = "X"
   Image1.Picture = imgWallTop.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub imgWeed_Click()
    strMapChr = "<"
   Image1.Picture = imgWeed.Picture
    TxtId.Text = "<"
    TxtGive.Text = "Weed"
End Sub

Private Sub ImgWGrass_Click()
   strMapChr = Chr(34)
   Image1.Picture = ImgWGrass.Picture
    TxtId.Text = Chr(34)
    TxtGive.Text = ""
End Sub

Private Sub imgWindow1_Click()
  strMapChr = "*"
   Image1.Picture = imgWindow1.Picture
    TxtId.Text = "0"
    TxtGive.Text = ""
End Sub

Private Sub ImgYBottle_Click()
  strMapChr = "^"
  Image1.Picture = ImgYBottle.Picture
   TxtId.Text = strMapChr
   TxtGive.Text = "YBottle"
End Sub

Private Sub Label1_Click()
  TxtGive.Text = "AXE"
End Sub

Private Sub Label2_Click()
 TxtGive.Text = "DESTROY"
End Sub

Private Sub Label5_Click()
  TxtGive.Text = "FILL"
End Sub

Private Sub Label6_Click()
 TxtGive.Text = "LIGHT"
End Sub

Private Sub Label7_Click()
 TxtGive.Text = "CUT"
End Sub
