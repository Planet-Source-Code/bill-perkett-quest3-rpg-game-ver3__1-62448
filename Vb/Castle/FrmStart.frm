VERSION 5.00
Begin VB.Form FrmStart 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CmdMap 
      Caption         =   "Map"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Game"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdMap_Click()
 SAVE_MapLoaded = "B1"
  strMyForm = "Frmgame"
  FrmMap.Show
End Sub

Private Sub Command1_Click()
  strMyForm = "FrmMake"
  FrmMake.Show
End Sub

Private Sub Command2_Click()
  SAVE_MapLoaded = "B1"
  strMyForm = "Frmgame"
  frmGame.Show
End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
  Label1.Caption = KeyAscii
End Sub

Private Sub Form_Load()
  SAVE_MapLoaded = "B1"
  SAVE_Midi = "Town.mid"
  SAVE_OutsideMidi = "Town.mid"
  SAVE_SpeechLoaded = "A1"
  SAVE_CharX = 10
  SAVE_CharY = 7
  SAVE_CharFacing = 2
  SAVE_MapLoaded = "A1"
  SAVE_SpellCut = False
  SAVE_SpellDestroy = False
End Sub
