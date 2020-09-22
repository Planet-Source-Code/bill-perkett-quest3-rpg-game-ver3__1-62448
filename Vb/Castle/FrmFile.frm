VERSION 5.00
Begin VB.Form FrmFile 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "FrmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim cMysource As String
   Dim MyFile As String
   Dim cDestination As String
   MyFile = Dir("c:\kids", vbDirectory)
   If Len(MyFile) < 1 Then MkDir "c:\kids"
   MyFile = Dir("c:\kids\Quest", vbDirectory)
   If Len(MyFile) < 1 Then
       MkDir "c:\kids\Quest"
        File1.Path = App.Path & "\Quest\"
         File1.Pattern = "*.txt"
        Dim i As Integer
        For i = 0 To File1.ListCount - 1
            cMysource = App.Path & "\Quest\" & File1.List(i)
            cDestination = "c:\kids\Quest\" & File1.List(i)
            FileCopy cMysource, cDestination
        Next
   End If
   '
   '
   '
   frmIntro.Show
   Unload Me
End Sub
