VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmSpeech 
   Caption         =   "Speech"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ComSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9340
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmSpeech.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   4815
      Left            =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8493
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "FrmSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim Tag As String
  Dim cFile As String
  Text1.Text = SAVE_MapLoaded
  cFile = "c:\Kids\Quest\" & SAVE_MapLoaded & "s.txt"
  Open cFile For Input As #1 ' Open file for input.
   i = 0
   RichTextBox1.Text = ""
   Do While Not EOF(1) ' Loop until end of file.
     Line Input #1, Tag ' Read data into two variables.
     If i = 0 Then
        RichTextBox1.Text = Tag
     Else
        RichTextBox1.Text = RichTextBox1.Text & vbCr & vbLf & Tag
     End If
     i = 1
   Loop
   Close #1
   ComSave.Visible = True
End Sub

Private Sub ComSave_Click()
  Dim cFile As String
  cFile = "c:\Kids\Quest\" & SAVE_MapLoaded & "s.txt"
   Open cFile For Output As #1 ' Open file for output.
   Print #1, RichTextBox1.Text  '& Chr(10) & Chr(13)
  Close #1    ' Close file.
  Label1.Caption = Time
End Sub

Private Sub Form_Load()
  FlexGrid.FormatString = "Command         |Description                  "
  FlexGrid.AddItem "0008000600000000" & Chr(9) & "Item Location"
  FlexGrid.AddItem "Mother:         " & Chr(9) & "Name"
  FlexGrid.AddItem "<THOUGHT>       " & Chr(9) & "Thought Start"
  FlexGrid.AddItem "Name!Mother:    " & Chr(9) & "Name"
  FlexGrid.AddItem "Thought!Hello   " & Chr(9) & "Thought"
  FlexGrid.AddItem "Pence!I can Help" & Chr(9) & "Feeling"
  FlexGrid.AddItem "HideThought!    " & Chr(9) & "Hide Thought until you can pence"
  FlexGrid.AddItem "WaitT!          " & Chr(9) & "Talk only if you read thought"
  FlexGrid.AddItem "WaitP!          " & Chr(9) & "Talk only if you can pence"
  'FlexGrid.AddItem "GiftT!Candle    " & Chr(9) & "Thought gift"
  'FlexGrid.AddItem "GiveP!Sword     " & Chr(9) & "Pence gift"
  FlexGrid.AddItem "<SPEECH>        " & Chr(9) & "Speech Start"
  FlexGrid.AddItem "<0001>          " & Chr(9) & "Number 1"
  FlexGrid.AddItem "Mother: Hello   " & Chr(9) & "Name"
  FlexGrid.AddItem "0003=Any advice?" & Chr(9) & "Answer = Number 3"
  FlexGrid.AddItem "0000=Goodby     " & Chr(9) & "End Speech"
  FlexGrid.AddItem "<0003>          " & Chr(9) & "Number 3"
  FlexGrid.AddItem "SayOnce!        " & Chr(9) & "Say only one time"
  FlexGrid.AddItem "TakeAny!        " & Chr(9) & "Take any inventory item"
  FlexGrid.AddItem "Win!            " & Chr(9) & "You win the game"
  FlexGrid.AddItem "Lose!           " & Chr(9) & "You lose the game"
  FlexGrid.AddItem "DoQuest!0001    " & Chr(9) & "Start quest 1"
  FlexGrid.AddItem "QuestName!Talk brother" & Chr(9) & "Quest Name"
  FlexGrid.AddItem "EndQuest!0001   " & Chr(9) & "End quest 1"
  FlexGrid.AddItem "QuestNo!0001    " & Chr(9) & "Only if Quest 1 is started"
  FlexGrid.AddItem "QuestYes!0001   " & Chr(9) & "Only if Quest 1 is done"
  FlexGrid.AddItem "Give!Ticket     " & Chr(9) & "Items given to you"
  FlexGrid.AddItem "Giveqty!01      " & Chr(9) & "Give Qty"
  FlexGrid.AddItem "NeedItem!Wood   " & Chr(9) & "Item you give away"
  FlexGrid.AddItem "NeedQty!01      " & Chr(9) & "Need Qty"
  FlexGrid.AddItem "BombQty!01      " & Chr(9) & "Bombs given to you"
  FlexGrid.AddItem "FixX!0010       " & Chr(9) & "X Pos - Make grass"
  FlexGrid.AddItem "FixY!0022       " & Chr(9) & "Y Pos - Make grass"
End Sub
