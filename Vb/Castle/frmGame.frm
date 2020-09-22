VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quest"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtSpeech 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
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
      Height          =   2175
      Left            =   2280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmGame.frx":0000
      Top             =   3600
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5880
      Top             =   3240
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00C0C000&
      ForeColor       =   &H00C0FFFF&
      Height          =   4935
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
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
         Left            =   2400
         TabIndex        =   11
         Top             =   2280
         Width           =   135
      End
      Begin VB.Image imgNToast 
         Height          =   480
         Left            =   1920
         Picture         =   "frmGame.frx":000A
         Top             =   2040
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
         Left            =   840
         TabIndex        =   10
         Top             =   2280
         Width           =   135
      End
      Begin VB.Image imgNTicket 
         Height          =   480
         Left            =   360
         Picture         =   "frmGame.frx":0C4C
         Top             =   2040
         Width           =   480
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   480
         X2              =   4200
         Y1              =   1800
         Y2              =   1800
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
         Left            =   3960
         TabIndex        =   9
         Top             =   1200
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
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   135
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
         TabIndex        =   7
         Top             =   1200
         Width           =   135
      End
      Begin VB.Image imgNMagic 
         Height          =   480
         Left            =   3480
         Picture         =   "frmGame.frx":188E
         Top             =   960
         Width           =   480
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
         TabIndex        =   6
         Top             =   4080
         Width           =   1230
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
         TabIndex        =   5
         Top             =   4080
         Width           =   1470
      End
      Begin VB.Image imgNCoin 
         Height          =   480
         Left            =   1920
         Picture         =   "frmGame.frx":24D0
         Top             =   960
         Width           =   480
      End
      Begin VB.Label etqTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Menu"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   1245
      End
      Begin VB.Image imgNWood 
         Height          =   480
         Left            =   360
         Picture         =   "frmGame.frx":3112
         Top             =   960
         Width           =   480
      End
   End
   Begin VB.Image ImgMap 
      Height          =   480
      Left            =   5280
      Picture         =   "frmGame.frx":3D54
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image ImgBBottle 
      Height          =   480
      Left            =   4080
      Picture         =   "frmGame.frx":4996
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image ImgYBottle 
      Height          =   480
      Left            =   4680
      Picture         =   "frmGame.frx":55D8
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image ImgKey2 
      Height          =   480
      Left            =   7680
      Picture         =   "frmGame.frx":621A
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgPine 
      Height          =   480
      Left            =   7680
      Picture         =   "frmGame.frx":6E5C
      Top             =   7800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGold 
      Height          =   480
      Left            =   7680
      Picture         =   "frmGame.frx":7A9E
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgLamp 
      Height          =   480
      Left            =   8280
      Picture         =   "frmGame.frx":86E0
      Top             =   7800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgApple 
      Height          =   480
      Left            =   8280
      Picture         =   "frmGame.frx":9322
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBucket 
      Height          =   480
      Left            =   8280
      Picture         =   "frmGame.frx":9F64
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgSaw 
      Height          =   480
      Left            =   8280
      Picture         =   "frmGame.frx":ABA6
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGem 
      Height          =   480
      Left            =   8280
      Picture         =   "frmGame.frx":B7E8
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBook 
      Height          =   480
      Left            =   8280
      Picture         =   "frmGame.frx":C42A
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMagic 
      Height          =   480
      Left            =   8280
      Picture         =   "frmGame.frx":C8CF
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem6 
      Height          =   480
      Left            =   8280
      Picture         =   "frmGame.frx":D511
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   8160
      TabIndex        =   15
      Top             =   840
      Width           =   735
   End
   Begin VB.Image imgItem5 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":E153
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem4 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":ED95
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStarGate 
      Height          =   480
      Left            =   720
      Picture         =   "frmGame.frx":F9D7
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharSoldier 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":10619
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem3 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":1125B
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      Height          =   195
      Left            =   4320
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "X"
      Height          =   195
      Left            =   4320
      TabIndex        =   2
      Top             =   8880
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgItem2 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":11E9D
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarryChar 
      Height          =   480
      Left            =   6120
      Picture         =   "frmGame.frx":12ADF
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem1 
      Height          =   540
      Left            =   6120
      Picture         =   "frmGame.frx":13721
      Top             =   4440
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgCharWomen3 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":14693
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen2 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":152D5
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBed2 
      Height          =   480
      Left            =   5400
      Picture         =   "frmGame.frx":15F17
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBed1 
      Height          =   480
      Left            =   5400
      Picture         =   "frmGame.frx":16B59
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTable2 
      Height          =   480
      Left            =   4920
      Picture         =   "frmGame.frx":1779B
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTable1 
      Height          =   480
      Left            =   4440
      Picture         =   "frmGame.frx":183DD
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExitB2 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":1901F
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExitB1 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":19C61
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgJar 
      Height          =   480
      Left            =   4440
      Picture         =   "frmGame.frx":1A8A3
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExit 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":1B4E5
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":1C127
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL3 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":1CD69
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL3 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":1D9AB
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT3 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":1E5ED
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":1F22F
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR3 
      Height          =   495
      Left            =   3000
      Picture         =   "frmGame.frx":1FE71
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":20B13
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB3 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":21755
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWall2 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":22397
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeed 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":22FD9
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBadWizard 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":23C1B
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharGoodWizard 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":2485D
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy2 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":2549F
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrees 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":260E1
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRockHill 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":26D23
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL2 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":27965
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR2 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":285A7
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOR2 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":291E9
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOL2 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":29E2B
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFloor1 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":2AA6D
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallTop 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":2B6AF
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallBottom 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":2C2F1
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":2CF33
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":2DB75
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":2E7B7
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT2 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":2F3F9
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":3003B
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":30C7D
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":318BF
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB2 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":32501
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRedTop 
      Height          =   480
      Left            =   5040
      Picture         =   "frmGame.frx":33143
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDoor1 
      Height          =   480
      Left            =   5040
      Picture         =   "frmGame.frx":33D85
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgNothing 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":349C7
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   320
      X2              =   320
      Y1              =   320
      Y2              =   0
   End
   Begin VB.Image imgBlueTop 
      Height          =   480
      Left            =   4560
      Picture         =   "frmGame.frx":35609
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWindow1 
      Height          =   480
      Left            =   4560
      Picture         =   "frmGame.frx":3624B
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen1 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":36E8D
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy1 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":37ACF
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDownChar 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":38711
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRightChar 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":39353
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLeftChar 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":39F95
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgUpChar 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":3ABD7
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":3B819
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOL 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":3C45B
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOR 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":3D09D
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":3DCDF
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":3E921
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":3F563
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":401A5
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":40DE7
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":41A29
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":4266B
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":432AD
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":43EEF
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlack 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":44B31
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBush 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":45773
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   320
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Image imgGrass 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":463B5
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSign 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":46FF7
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'******************************************************
'Andres Zacarias and Bill Perkett
'Quest Game
'Game Type: RPG
'
'This game is an idea from my favourite game : Zelda64 Ocarina
'of time.
'
'If you have any idea of how to improve the speech and the item finding please
'email Bill Perkett
'Im really not comenting this thing.
'
'NOTE: You must first press Escape Key before closing the game or
'else vb will crash.
'******************************************************
'******************************************************


Option Explicit


Private Sub Command3_Click()
     Text3.Enabled = False
     Text2.Enabled = False
     Command3.Enabled = False
     CharX = Text2.Text
      CharY = Text3.Text
     DrawIt
End Sub

Private Sub etqBacktoGame_Click()
  fraMenu.Visible = False
End Sub

Private Sub etqBacktoGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  etqBacktoGame.ForeColor = &HFF&
  etqSaveGame.ForeColor = &HFFFFFF
End Sub

Private Sub etqSaveGame_Click()
  Call SaveGame
End Sub

Private Sub etqSaveGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  etqBacktoGame.ForeColor = &HFFFFFF
  etqSaveGame.ForeColor = &HFF&
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyV
     Text3.Enabled = True
     Text2.Enabled = True
     Command3.Enabled = True
  Case vbKeyUp
    'Key Up.
    PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
    CharFacing = 1
    Label1 = CharY
    Call CharacterMovements
    If TxtSpeech.Visible = True Then TxtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharY = CharY - 1
      DrawIt
    End If
    If PositionMap = "#" Or PositionMap = "@" Then
      Call SelectPlace
      DrawIt
    End If
  Case vbKeyDown
    'Key Down.
    PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
    CharFacing = 2
    Label1 = CharY
    Call CharacterMovements
    If TxtSpeech.Visible = True Then TxtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharY = CharY + 1
      DrawIt
    End If
    If PositionMap = "M" Then
      Call SelectPlace
      DrawIt
    End If
  Case vbKeyRight
    'Key Right.
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    CharFacing = 3
    Label2 = CharX
    Call CharacterMovements
    If TxtSpeech.Visible = True Then TxtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharX = CharX + 1
      DrawIt
    End If
  Case vbKeyLeft
    'Key Left.
    PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    CharFacing = 4
    Label2 = CharX
    Call CharacterMovements
    If TxtSpeech.Visible = True Then TxtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharX = CharX - 1
      DrawIt
    End If
  Case vbKeySpace
    'Key Space.
    Call TypeOfItem
    DrawIt
    Call InitializeSpeech
  Case vbKeyZ
    If SpellCut = True Then
      Call Make_Spell_Cut
      DrawIt
    End If
  Case vbKeyX
    If SpellDestroy = True Then
      Call Make_Spell_Destroy
      DrawIt
    End If
  
  Case vbKeyEscape
    CloseMidi
    UnHook
  Case vbKeyReturn
    Call ItemCount
    fraMenu.Visible = True
  End Select
   Text2.Text = CharX
  Text3.Text = CharY
End Sub

Private Sub Form_Load()
  'frmGame.Height = 5175
  'frmGame.Width = 4890
  'char
  Call InitGame
  CharX = 19
  CharY = 11
  DrawIt
  Midi = SAVE_Midi
  Speech_B1
  gHW = frmGame.hWnd
 ' Call InitMusic
End Sub

Public Sub InitGame()
  
  'Character Position.
    MapLoaded = SAVE_MapLoaded
    SpeechLoaded = SAVE_SpeechLoaded
    CharX = SAVE_CharX
    CharY = SAVE_CharY
    CharFacing = SAVE_CharFacing
    Wood = SAVE_Wood
    Coin = SAVE_Coin
    Magic = SAVE_Magic
    SpellCut = SAVE_SpellCut
    SpellDestroy = SAVE_SpellDestroy
  'Starting the first Map.
  If MapLoaded = "A1" Then Call Map_A1
  If MapLoaded = "A2" Then Call Map_A2
  If MapLoaded = "A3" Then Call Map_A3
  If MapLoaded = "A4" Then Call Map_A4
  If MapLoaded = "B1" Then Call Map_B1
  If MapLoaded = "B2" Then Call Map_B2
  If MapLoaded = "B3" Then Call Map_B3
  If MapLoaded = "B4" Then Call Map_B4
  If MapLoaded = "B5" Then Call Map_B5
  'Load the Speech for the Map.
  If SpeechLoaded = "A1" Then Call Speech_A1
  If SpeechLoaded = "B1" Then Call Speech_B1
  NewLine = Chr(13) + Chr(10)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub



Private Sub Timer1_Timer()
  DrawIt
  Timer1.Enabled = False
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
        If Y + CharY + 0 > 51 Then PictureHandler
      If PassToNext = 0 Then PositionMap = Mid(Map(Y + CharY + 1), (X + CharX + 1), 1)
      'If X = 0 And Y = 0 Then GoTo skip:
      Select Case PositionMap
      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "c" 'Saw
        PaintPicture ImgSaw.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "g" 'Saw
        PaintPicture ImgBucket.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "O" 'Lamp
        PaintPicture ImgLamp.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "o" 'Lamp
        PaintPicture ImgGold.Picture, (X + 3) * 32, (Y + 3) * 32
      'Case Is = "x" 'Lamp
       ' PaintPicture ImgPGem.Picture, (X + 3) * 32, (Y + 3) * 32
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

Public Sub PictureHandler()
  PassToNext = 1
  PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
End Sub

Public Sub CharacterMovements()
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
  End Select
End Sub

