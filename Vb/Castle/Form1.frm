VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "frmHelp"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "Close"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox TxtFile 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   840
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "Quest Instructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdEnd_Click()
  Form1.Hide
End Sub

Private Sub Command2_Click()
    Printer.Print "  "
     Printer.Print "  "
     Printer.Print TxtFile.Text
End Sub


Private Sub Form_Load()
   Dim cFile As String
    cFile = App.Path & "\readQuest.txt"
    TxtFile.Text = ""
    Open cFile For Input As #1 ' Open file for input.
    Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, MyString ' Read data into two variables.
      TxtFile.Text = TxtFile.Text & MyString & Chr(13) & Chr(10) '& " " & Chr(13) & Chr(10)
   Loop
   Close #1    ' Close file.
End Sub
