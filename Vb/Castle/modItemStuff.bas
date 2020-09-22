Attribute VB_Name = "modItemStuff"

Public Function TypeOfItem() As String
  
  If CharFacing = 1 Then
    PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
    If PositionMap = "!" Then Call Jar(CharY + 2, CharX + 3)
   End If
  If CharFacing = 2 Then
    PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
    If PositionMap = "!" Then Call Jar(CharY + 4, CharX + 3)
  End If
  If CharFacing = 3 Then
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    If PositionMap = "!" Then Call Jar(CharY + 3, CharX + 4)
  End If
  If CharFacing = 4 Then
    PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    If PositionMap = "!" Then Call Jar(CharY + 3, CharX + 2)
  End If
  TypeOfItem = PositionMap
End Function
Public Sub BombWall()
 Dim MFixX As Integer
 If Bomb > 0 Then
  Bomb = Bomb - 1
  If CharFacing = 1 Then
    PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
    If PositionMap = "P" Then
      MFixX = CharX + 3
      strMap = Mid(Map(CharY + 2), 1, MFixX - 1) & "?" & Mid(Map(CharY + 2), MFixX + 1, Len(Map(MFixY)) - MFixX)
      Map(CharY + 2) = strMap
    End If
'    If PositionMap = "p" Then
'      MFixX = CharX + 3
'      strMap = Mid(Map(CharY + 2), 1, MFixX - 1) & "?" & Mid(Map(CharY + 2), MFixX + 1, Len(Map(MFixY)) - MFixX)
'      Map(CharY + 2) = strMap
'    End If
  End If
  '
  '
  '
  If CharFacing = 2 Then
    PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
    'If PositionMap = "!" Then Call Jar
    If PositionMap = "P" Then
      MFixX = CharX + 3
      strMap = Mid(Map(CharY + 4), 1, MFixX - 1) & "?" & Mid(Map(CharY + 4), MFixX + 1, Len(Map(MFixY)) - MFixX)
      Map(CharY + 4) = strMap
    End If
'    If PositionMap = "p" Then
'      MFixX = CharX + 3
'      strMap = Mid(Map(CharY + 4), 1, MFixX - 1) & "?" & Mid(Map(CharY + 4), MFixX + 1, Len(Map(MFixY)) - MFixX)
'      Map(CharY + 4) = strMap
'    End If
  End If
  '
  '
  '
  If CharFacing = 3 Then
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    'If PositionMap = "!" Then Call Jar
    If PositionMap = "P" Then
      MFixX = CharX + 4
      strMap = Mid(Map(CharY + 3), 1, MFixX - 1) & "?" & Mid(Map(CharY + 3), MFixX + 1, Len(Map(MFixY)) - MFixX)
      Map(CharY + 3) = strMap
    End If
'    If PositionMap = "p" Then
'      MFixX = CharX + 4
'      strMap = Mid(Map(CharY + 3), 1, MFixX - 1) & "?" & Mid(Map(CharY + 3), MFixX + 1, Len(Map(MFixY)) - MFixX)
'      Map(CharY + 3) = strMap
'    End If
  End If
  '
  '
  '
  If CharFacing = 4 Then
    PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    'If PositionMap = "!" Then Call Jar
    If PositionMap = "P" Then
      MFixX = CharX + 2
      strMap = Mid(Map(CharY + 3), 1, MFixX - 1) & "?" & Mid(Map(CharY + 3), MFixX + 1, Len(Map(MFixY)) - MFixX)
      Map(CharY + 3) = strMap
    End If
'    If PositionMap = "p" Then
'      MFixX = CharX + 2
'      strMap = Mid(Map(CharY + 3), 1, MFixX - 1) & "?" & Mid(Map(CharY + 3), MFixX + 1, Len(Map(MFixY)) - MFixX)
'      Map(CharY + 3) = strMap
'    End If
  End If
End If
End Sub
Public Sub ItemCount()
  frmGame.etqNWood = Wood
  frmGame.etqNCoin = Coin
  frmGame.etqNMagic = Magic
  frmGame.etqNTicket = Ticket
  frmGame.etqNToast = Toast
End Sub

Public Sub Jar(iX As Integer, iY As Integer)
  Dim j As Integer
  Dim RandomNumber As Integer
  'Find Wood
  'If MapLoaded = "A2" Or MapLoaded = "A3" Or MapLoaded = "A4" Or MapLoaded = "B2" Or MapLoaded = "B3" Or MapLoaded = "B4" Then
    Randomize Timer
    RandomNumber = Int(Rnd * 3.999)
    bRun = False
    FrmMake.Frame2.Visible = False
    If iJarCnt < 0 Then
        FrmMake.txtSpeech.Visible = True
        FrmMake.txtSpeech.Text = "Jar Empty"
        Exit Sub
    End If
     For j = 1 To iJarLast
      If iJar(1, j) = iX And iJar(2, j) = iY Then
        FrmMake.txtSpeech.Visible = True
        FrmMake.txtSpeech.Text = "Jar is now Empty"
        Exit Sub
      End If
    Next j
    iJar(1, iJarLast) = iX
    iJar(2, iJarLast) = iY
    iJarLast = iJarLast + 1
    If RandomNumber = 0 Then
        FrmMake.txtSpeech.Visible = True
        FrmMake.txtSpeech.Text = "Jar Empty"
    End If
    If RandomNumber = 1 Then
        If ItemHouseFound1 = False Then
            CharFacing = 5
            Item = 1
            Wood = Wood + 1
            FrmMake.Frame2.Visible = False
            FrmMake.txtSpeech.Visible = True
            FrmMake.txtSpeech.Text = "You found some wood"
           ItemHouseFound1 = True
           iJarCnt = iJarCnt - 1
       Else
          FrmMake.Frame2.Visible = False
          FrmMake.txtSpeech.Visible = True
          FrmMake.txtSpeech.Text = "Jar Empty"
      End If
    End If
    '
    '
    If RandomNumber = 2 Then
        If ItemHouseFound2 = False Then
            CharFacing = 5
            Item = 2
            Coin = Coin + 1
            FrmMake.Frame2.Visible = False
            FrmMake.txtSpeech.Visible = True
            FrmMake.txtSpeech.Text = "You found a coin"
            ItemHouseFound2 = True
            iJarCnt = iJarCnt - 1
         Else
         FrmMake.Frame2.Visible = False
          FrmMake.txtSpeech.Visible = True
          FrmMake.txtSpeech.Text = "Jar Empty"
       End If
     End If
     '
     '
     '
    If RandomNumber = 3 Then
      If ItemHouseFound3 = False Then
        CharFacing = 5
        Item = 3
        Magic = Magic + 10
        FrmMake.Frame2.Visible = False
        FrmMake.txtSpeech.Visible = True
        FrmMake.txtSpeech.Text = "You found some magic"
        ItemHouseFound3 = True
        iJarCnt = iJarCnt - 1
     Else
         FrmMake.Frame2.Visible = False
          FrmMake.txtSpeech.Visible = True
          FrmMake.txtSpeech.Text = "Jar Empty"
     End If
   End If
   'DrawIt
 End Sub

