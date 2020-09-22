VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmQuests 
   Caption         =   "Quests"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   3255
      Begin VB.Label Label7 
         Caption         =   "H - Hack through Swamp  "
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
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label8 
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
         Left            =   600
         TabIndex        =   7
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label9 
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
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
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
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label12 
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
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   520
         Width           =   1575
      End
      Begin VB.Label Label13 
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
         Left            =   600
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label11 
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
         Left            =   600
         TabIndex        =   2
         Top             =   1740
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7646
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim i As Integer
   Dim cMsg As String
   FlexGrid.Rows = 1
   FlexGrid.FormatString = "Num|Done|Quest                                                   "
   For i = 1 To 15
     If MyQuest(i) <> " " Then
       cMsg = i & Chr(9) & Left(MyQuest(i), 1) & Chr(9) & Right(MyQuest(i), Len(MyQuest(i)) - 1)
       FlexGrid.AddItem cMsg
     End If
   Next
End Sub
