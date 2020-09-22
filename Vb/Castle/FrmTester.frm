VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmTester 
   Caption         =   "FrmSpeechTester"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form2"
   ScaleHeight     =   7110
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1935
      Left            =   360
      TabIndex        =   7
      Top             =   4800
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmTester.frx":0000
   End
   Begin VB.TextBox TxtSpeech 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Text            =   "A1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Read Map  "
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Holder 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
      Begin VB.Label Pitanje 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   ":::"
         ForeColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4320
         WordWrap        =   -1  'True
      End
      Begin VB.Label Choice 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   ">>>"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   4290
         WordWrap        =   -1  'True
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   3615
      Left            =   5640
      TabIndex        =   3
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6376
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
End
Attribute VB_Name = "FrmTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Choice_Click(Index As Integer)
 Dim cMsg As String
  Dim i As Integer
  Dim MFixX As Integer
  Dim MFixY As Integer
  For i = 0 To 5
     Choice(i).Visible = False
  Next
  If Choice(Index).Tag <> "0000" Then
   cMsg = "<" & Choice(Index).Tag & ">"
   For Each ClsSpeech In nSpeech
   If ClsSpeech.RNumber = cMsg Then
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
         If MyQuest(i) = " " Then MyQuest(i) = "N" & Trim(ClsSpeech.QuestName)
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
            End If
         End If
        End If
        MFixX = Val(ClsSpeech.FixX)
        MFixY = Val(ClsSpeech.FixY)
        If MFixX > 0 And MFixY > 0 Then
           strMap = Mid(Map(MFixY), 1, MFixX - 1) & "G" & Mid(Map(MFixY), MFixX + 1, Len(Map(MFixY)) - MFixX)
            Map(MFixY) = strMap
            'DrawIt
        End If
    End If
        'Text1.Text = Text1.Text & cEvent & " -- " & cDesc & Chr(13) & Chr(10)
    Next
     'PicTXT.Visible = True
  Else
   'PicTXT.Visible = False
   'TxtSpeech.Visible = False
  End If
End Sub

Private Sub Command3_Click()
  Dim i As Integer
   Dim cMsg As String
   MeLoaded = Text1.Text
   ReadMapFile (MeLoaded)
   'SAVE_MapLoaded = MeLoaded
    FlexGrid.Rows = 1
   FlexGrid.FormatString = "    X|    Y|Name        "
   For i = 0 To iMsgCnt
     cMsg = Mid(Message(i), 1, 4) & Chr(9) & Mid(Message(i), 5, 4) & Chr(9) & Trim(Right(Message(i), Len(Message(i)) - 16))
     FlexGrid.AddItem cMsg
   Next
   
End Sub

Private Sub Command7_Click()
 Dim cMsg As String
 Dim iQuest As Integer
 Dim iQuestYes As Integer
 Dim iQuestNo As Integer
 Dim i As Integer
 Dim MFixX As Integer
 Dim MFixY As Integer
 'Frame1.Visible = False
 For i = 0 To 5
  Choice(i).Visible = False
 Next
 
 Pitanje.Visible = True
 iQuest = 0
 iQuestYes = 0
  cMsg = "<0001>"
  RichTextBox1.Text = ""
  For Each ClsSpeech In nSpeech
   If ClsSpeech.Name = TxtSpeech.Text Then
     iQuest = Val(ClsSpeech.EndQuest)
     iQuestYes = Val(ClsSpeech.QuestYes)
     iQuestNo = Val(ClsSpeech.QuestNo)
     If UCase(ClsSpeech.Need) = "WEED" And Weed < ClsSpeech.NeedQty Then GoTo MySkipTalk
     If UCase(ClsSpeech.Need) = "ROCK" And Rock < ClsSpeech.NeedQty Then GoTo MySkipTalk
     If UCase(ClsSpeech.Need) = "WEED" Then
       Weed = 0
     End If
     If ClsSpeech.Need = "ROCK" Then Rock = 0
     RichTextBox1.Text = RichTextBox1.Text & ClsSpeech.RNumber & " " & ClsSpeech.Name & ClsSpeech.Talk & cbcr & vbLf
     RichTextBox1.Text = RichTextBox1.Text & "End = " & ClsSpeech.EndQuest & cbcr & vbLf
     RichTextBox1.Text = RichTextBox1.Text & "Yes = " & ClsSpeech.QuestYes & cbcr & vbLf
     RichTextBox1.Text = RichTextBox1.Text & "No = " & ClsSpeech.QuestNo & cbcr & vbLf
     If (ClsSpeech.EndQuest = "0" Or Mid(MyQuest(iQuest), 1, 1) <> " ") And ClsSpeech.SayOnce <> "D" And _
        (ClsSpeech.QuestYes = "0" Or Mid(MyQuest(iQuestYes), 1, 1) = "Y") And _
        (ClsSpeech.QuestNo = "0" Or Mid(MyQuest(iQuestNo), 1, 1) = "N") Then
        Pitanje.Caption = ClsSpeech.Name & ClsSpeech.Talk
        If ClsSpeech.SayOnce = "Y" Then ClsSpeech.SayOnce = "D"
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
          If MyQuest(i) = " " Then MyQuest(i) = "N" & Trim(ClsSpeech.QuestName)
        End If
        If ClsSpeech.EndQuest <> " " Then
          i = Val(ClsSpeech.EndQuest)
          If Left(MyQuest(i), 1) = "N" Then
             MyQuest(i) = "Y" & Right(MyQuest(i), Len(MyQuest(i)) - 1)
               iGiveQty = ClsSpeech.BombQty
               If iGiveQty > 0 Then TheyGaveToMe ("BOMB")
            End If
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
            'DrawIt
        End If
       GoTo MyTalk
     End If
MySkipTalk:
   End If
        'Text1.Text = Text1.Text & cEvent & " -- " & cDesc & Chr(13) & Chr(10)
    Next
MyTalk:
     'PicTXT.Visible = True
End Sub

Private Sub FlexGrid_Click()
   If FlexGrid.MouseRow <> 0 Then
     FlexGrid.Col = 2
     TxtSpeech.Text = FlexGrid.Text
     If Right(TxtSpeech.Text, 1) = ":" Then Command7_Click
   End If
End Sub

Private Sub Form_Load()
  Dim i As Integer
  For i = 1 To 5
      Load Choice(Choice.Count)
   Next
End Sub
