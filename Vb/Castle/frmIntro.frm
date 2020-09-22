VERSION 5.00
Begin VB.Form frmIntro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quest"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   5040
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4880
      Left            =   5040
      Top             =   120
   End
   Begin VB.TextBox txtStory 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
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
      Height          =   5175
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmIntro.frx":0000
      Top             =   -240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label etqHelp 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   3600
      Width           =   570
   End
   Begin VB.Image imgRedTop 
      Height          =   480
      Left            =   2760
      Picture         =   "frmIntro.frx":0006
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDoor1 
      Height          =   480
      Left            =   1320
      Picture         =   "frmIntro.frx":0C48
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBlueTop 
      Height          =   480
      Left            =   840
      Picture         =   "frmIntro.frx":188A
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWindow1 
      Height          =   480
      Left            =   840
      Picture         =   "frmIntro.frx":24CC
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label etqMakeGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MakeGame"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   1380
   End
   Begin VB.Image ImgBlack 
      Height          =   480
      Left            =   2400
      Picture         =   "frmIntro.frx":310E
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRightChar 
      Height          =   480
      Left            =   1920
      Picture         =   "frmIntro.frx":3D50
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgGrass 
      Height          =   480
      Left            =   1440
      Picture         =   "frmIntro.frx":4992
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBush 
      Height          =   480
      Left            =   1920
      Picture         =   "frmIntro.frx":55D4
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL 
      Height          =   480
      Left            =   480
      Picture         =   "frmIntro.frx":6216
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR 
      Height          =   480
      Left            =   960
      Picture         =   "frmIntro.frx":6E58
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT 
      Height          =   480
      Left            =   1440
      Picture         =   "frmIntro.frx":7A9A
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL 
      Height          =   480
      Left            =   480
      Picture         =   "frmIntro.frx":86DC
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR 
      Height          =   480
      Left            =   960
      Picture         =   "frmIntro.frx":931E
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label etqLoadGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Load Current Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   2385
   End
   Begin VB.Label etqNewGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Start New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   1080
      Picture         =   "frmIntro.frx":9F60
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2670
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   320
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   320
      X2              =   320
      Y1              =   320
      Y2              =   0
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub etqHelp_Click()
  Form1.Show
End Sub

Private Sub etqHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 etqNewGame.ForeColor = &HFFFFFF
  etqHelp.ForeColor = &HFF&
  etqMakeGame.ForeColor = &HFFFFFF
  etqLoadGame.ForeColor = &HFFFFFF
End Sub

Private Sub etqLoadGame_Click()
  strChrName = InputBox("Enter your name ", "Character Name", "Sheik")
  Timer2.Enabled = False
  bMyPlayGame = True
  Call LoadGame
  FrmMake.Show
  'iBunnyCaught = 0
  Unload Me
End Sub

Private Sub etqLoadGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  etqNewGame.ForeColor = &HFFFFFF
  etqHelp.ForeColor = &HFFFFFF
  etqMakeGame.ForeColor = &HFFFFFF
  etqLoadGame.ForeColor = &HFF&
End Sub

Private Sub etqMakeGame_Click()
  strChrName = "Goblin"
  bMyPlayGame = False
  SAVE_MapLoaded = "A1"
  SAVE_Midi = "Town.mid"
  SAVE_OutsideMidi = "Town.mid"
  SAVE_SpeechLoaded = "A1"
  SAVE_CharX = 10
  SAVE_CharY = 7
  SAVE_CharFacing = 2
  SAVE_MapLoaded = "A1"
  SAVE_SpellCut = False
  SAVE_SpellDestroy = False
  strInventory = ""
  FrmMake.Show
  iBunnyCaught = 0
  Timer1.Enabled = False
  Unload Me
End Sub

Private Sub etqMakeGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'strChrName = InputBox("Enter your name ", "Character Name", Default)
   etqNewGame.ForeColor = &HFFFFFF
   etqLoadGame.ForeColor = &HFFFFFF
    etqHelp.ForeColor = &HFFFFFF
   etqMakeGame.ForeColor = &HFF&
End Sub

Private Sub etqNewGame_Click()
  Dim i As Integer
  strChrName = InputBox("Enter your name ", "Character Name", "Sheik")
  Timer1.Enabled = True
  Timer2.Enabled = False
  txtStory.Visible = True
  frmIntro.Cls
  frmIntro.BackColor = &H0&
  For i = 0 To 99
         MyQuest(i) = " "
   Next
  '  Midi = "Bad.mid"
'  gHW = frmIntro.hWnd
'  Call InitMusic
   iBunnyCaught = 0
  Call LoadStory
End Sub

Private Sub etqNewGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  etqNewGame.ForeColor = &HFF&
  etqLoadGame.ForeColor = &HFFFFFF
   etqHelp.ForeColor = &HFFFFFF
  etqMakeGame.ForeColor = &HFFFFFF
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  
  Select Case KeyCode
    Case vbKeyEscape
    CloseMidi
    UnHook
  End Select

End Sub

Private Sub Form_Load()
  frmIntro.Height = 5175
  frmIntro.Width = 4890
  iMagicPoints = 0
'  Midi = "intro.mid"
'  gHW = frmIntro.hWnd
'  Call InitMusic
  Randomize
  SAVE_SpellWade = False
  Call InitIntro
  'Call DrawIt2
End Sub

Public Sub InitIntro()
  'Starting the first Map.
  Call Map_Intro
  'Character Position.
  CharX = 4
  CharY = 5
  'Facing Position.
  CharFacing = 3
End Sub

Public Sub LoadStory()
  Dim Story(5) As String
  Dim DSpace As String
  DSpace = vbCrLf & vbCrLf
  Story(0) = "Far away, in the kigndom of Hyrule, there was a legend of that the chosen one will save the kingdom from a great dark force."
  Story(1) = "The time has arrived for our hero sheik to do his job."
  Story(2) = "But the quest will not be easy, sheik must learn spells " & DSpace
  Story(2) = Story(2) & "and to get the different items."
  Story(3) = "The future of the kingdom depends on you. Good Luck..."
  txtStory.Text = DSpace & DSpace & Story(0) & DSpace & Story(1) & DSpace & Story(2) & DSpace & Story(3)
End Sub

Private Sub Label1_Click()
   Form1.Show
End Sub

Private Sub Timer1_Timer()
  SAVE_MapLoaded = "New"
  SAVE_Midi = "Town.mid"
  SAVE_OutsideMidi = "Town.mid"
  SAVE_SpeechLoaded = "New"
  SAVE_CharX = 9
  SAVE_CharY = 7
  SAVE_CharFacing = 2
  SAVE_MapLoaded = "New"
  iSaveMapcnt = 0
  iSaveMapLast = 0
  iSaveMapLoc = 0
  SAVE_Wood = 0
  SAVE_Coin = 0
  SAVE_Magic = 0
  SAVE_Ticket = 0
  SAVE_Toast = 0
  SAVE_Bomb = 0
  strInventory = ""
  SpellCut = False
  SpellDestroy = False
  SpellWade = False
  SpellLight = False
  SpellAxe = False
  Item = 0
  ItemTownFound1 = False
  bMyPlayGame = True
  FrmMake.Show
  Timer1.Enabled = False
  Unload Me
End Sub

Public Sub IntroMov()
    
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    'PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    DrawIt2

End Sub


Private Sub Timer2_Timer()
  CharX = CharX + 1
  If CharX = 21 Then CharX = 10
  Call IntroMov
  DrawIt2
  'Timer2.Enabled = False
End Sub

Public Sub DrawIt2()
  For Y = -3 To 6
    For X = -3 To 6
      'If the result to Paint is 0 then it will get error.
      'This will prevent this.
      PassToNext = 0
        If Y + CharY + 0 < 1 Then PictureHandler2
        If X + CharX + 0 < 1 Then PictureHandler2
        If X + CharX + 0 > Len(Map(1)) Then PictureHandler2
        If Y + CharY + 0 > 51 Then PictureHandler2
      If PassToNext = 0 Then PositionMap = Mid(Map(Y + CharY + 1), (X + CharX + 1), 1)
      'If X = 0 And Y = 0 Then GoTo skip:
      Select Case PositionMap
         
      Case Is = "/"
        PaintPicture imgBlueTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "\"
        PaintPicture imgRedTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "*"
        PaintPicture imgWindow1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "#" '3
        'PaintPicture imgNothing.Picture, (X + 3) * 32, (Y + 3) * 32
        PaintPicture imgDoor1.Picture, (X + 3) * 32, (Y + 3) * 32

      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "Q" 'Water
        PaintPicture imgTOL.Picture, (X + 3) * 32, (Y + 3) * 32
      'Case Is = "A" 'Water
      '  PaintPicture imgBOL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "W" 'Water
        PaintPicture imgTOR.Picture, (X + 3) * 32, (Y + 3) * 32
      'Case Is = "S" 'Water
      '  PaintPicture imgBOR.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "E" 'Border Left water
        PaintPicture ImgIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "R" 'Border Right water
        PaintPicture ImgIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "D" 'Border Top water
        PaintPicture ImgIT.Picture, (X + 3) * 32, (Y + 3) * 32
      'Case Is = "F" 'Border Bottom water
      '  PaintPicture ImgIB.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "-" 'Water
        PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
                
      End Select
skip:
    Next
  Next
  

  'Character Movements.
  Select Case CharFacing
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * 32, 5 * 32
  End Select
  
End Sub

Public Sub PictureHandler2()
  PassToNext = 1
  PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
End Sub

