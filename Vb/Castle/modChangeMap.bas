Attribute VB_Name = "modChangeMap"
Public iMsgCnt As Integer
Public strMyForm As String
Public bMainMapLoaded As Boolean
Public strMainMap As String
Public strInventory As String
Public MeLoaded As String
Public nSpeech As Collection
Public ClsSpeech As ClsSpeech
Public nThought As Collection
Public ClsThought As ClsThought
Public MyQuest(100) As String
Public MakeQuest(100) As String
Public QuestMap(100) As String
Public Mygive As String
Public iNeedQty As Integer
Public iGiveQty As Integer
Public MyGiveLetter As String
Public bMyPlayGame As Boolean
Public bMyMakeMove As Boolean
Public bZoom As Boolean
Public iBunny As Integer
Public iBunnyloc(50, 3) As Integer
Public iBunnyCaught As Integer
Public iMagicPoints As Integer
Public strChrName As String
'
'
'
Public strSaveMap(9000) As String
Public strSaveMapName(500, 3) As String
Public iSaveMapcnt As Integer
Public iSaveMapLast As Integer
Public iSaveMapLoc As Integer
Public iJar(2, 100) As Integer
Public iJarCnt As Integer
Public iJarLast As Integer
Public bRun As Boolean
Public iMyScore As Integer
Public iQuestTot As Integer
Public IQuestDone As Integer


Public Sub TheyGaveToMe(strGive As String)
  Dim cMsg As String
  Dim j As Integer
  If strGive = "TICKET" Then Ticket = iGiveQty
  If strGive = "COIN" Then Coin = Coin + iGiveQty
  If strGive = "CUT" Then SpellCut = True
  If strGive = "DESTROY" Then SpellDestroy = True
  If strGive = "FILL" Then SpellWade = True
  If strGive = "AXE" Then SpellAxe = True
  If strGive = "LIGHT" Then SpellLight = True
  If strGive = "TOAST" Then Toast = Toast + iGiveQty
  If strGive = "BOMB" Then Bomb = Bomb + iGiveQty
  If Bomb > 8 Then Bomb = 8
  cMsg = " "
   Select Case strGive
    Case "SAW" 'Saw
        cMsg = "c" '"SAW"
     Case "BUCKET" '"g" 'Bucket
       cMsg = "g"
     Case "LAMP"  '"O" 'Lamp
        cMsg = "O"
     Case "GOLD" '"o" 'GOLD
        cMsg = "o"
     Case "PURPLEGEM"      'Is = "x" 'PURPLEGEM
       cMsg = "x"
     Case "GREENGEM"      'Is = "x" 'PURPLEGEM
       cMsg = "."
       Case "REDGEM"      'Is = "x" 'PURPLEGEM
       cMsg = "&"
      Case "APPLE" 'Is = "m" 'Apple
        cMsg = "m"
      Case "YKEY" 'Is = "L" 'YKey
       cMsg = "L"
     Case "BOTTLE" '"l" 'Bottle
       cMsg = "l"
     Case "RKEY" ' "~" 'RKey
        cMsg = "~"
     Case "BOOK" 'Is = "N" 'Book
       cMsg = "N"
     Case "BLUEGEM" 'Is = "n" 'Gem
       cMsg = "n"
      Case Is = "MAP" '"=" 'Map
        cMsg = "="
      Case Is = "BBOTTLE" '"%" 'BBottle
       cMsg = "%"
      Case Is = "YBOTTLE" ' "^" 'YBottle
       cMsg = "^"
     Case Is = "SWORD" ' "^" 'Sword
       cMsg = "{"
     Case Is = "BOW" '"%" 'Bow
       cMsg = "}"
    Case Is = "ARMOR" ' "^" 'Armor
       cMsg = "|"
     Case Is = "SHIELD" ' "^" 'Shield
       cMsg = ":"
     Case Is = "CANDLE" ' "^" 'Shield
       cMsg = "$"
     Case 1
       cMsg = " "
     End Select
     If cMsg <> " " Then
       j = InStr(1, strInventory, cMsg)
       If j = 0 Then strInventory = strInventory + cMsg
     End If
End Sub
Public Function TakeAny() As String
   Dim cMsg As String
   Dim strMe As String
   Dim strMap As String
   Dim j As Integer
   TakeAny = ""
   If Len(strInventory) = 0 Then Exit Function
      i = Rnd * Len(strInventory) + 1
      If i < 1 Or i > Len(strInventory) Then i = Len(strInventory)
      strMe = Mid(strInventory, i, 1)
      strMap = ""
      For j = 1 To Len(strInventory)
           If Mid(strInventory, j, 1) <> strMe Then strMap = strMap & Mid(strInventory, j, 1)
          Next j
        strInventory = strMap
     Select Case strMe
      Case Is = "c" 'Saw
        cMsg = "SAW"
      Case Is = "g" 'Bucket
       cMsg = "BUCKET"
      Case Is = "O" 'Lamp
        cMsg = "LAMP"
      Case Is = "o" 'GOLD
        cMsg = "GOLD"
      Case Is = "x" 'Purplegem
       cMsg = "PURPLEGEM"
      Case Is = "." 'Purplegem
       cMsg = "GREENGEM"
      Case Is = "&" 'Purplegem
       cMsg = "REDGEM"
      Case Is = "m" 'Apple
        cMsg = "APPLE"
      Case Is = "L" 'YKey
       cMsg = "YKEY"
      Case Is = "l" 'Bottle
       cMsg = "BOTTLE"
      Case Is = "~" 'RKey
        cMsg = "RKEY"
      Case Is = "N" 'Book
       cMsg = "BOOK"
      Case Is = "n" 'Gem
       cMsg = "BLUEGEM"
      Case Is = "=" 'Map
        cMsg = "MAP"
      Case Is = "%" 'BBottle
       cMsg = "BBOTTLE"
      Case Is = "^" 'YBottle
       cMsg = "YBOTTLE"
      Case Is = "{" 'Sword
       cMsg = "SWORD"
       Case Is = "}" 'Bow
       cMsg = "BOW"
      Case Is = "|" 'Armor
       cMsg = "ARMOR"
      Case Is = ":" 'Shield
       cMsg = "SHIELD"
       Case Is = "$" 'Candle
       cMsg = "CANDLE"
      End Select
       TakeAny = " Thanks for the " & cMsg
 End Function
Public Sub ReadMapFile(cMapfile As String)
   Dim i  As Integer
   Dim j  As Integer
   Dim jBun As Integer
   Dim iMap  As Integer
   Dim strMap As String
   Dim cFile
   Dim jquest As Integer
   ItemHouseFound1 = False
   ItemHouseFound2 = False
   ItemHouseFound3 = False
   ItemHouseFound4 = False
   iMap = 0
   iBunny = 0
   SAVE_MapLoaded = cMapfile
   For j = 1 To 50
    iBunnyloc(j, 3) = 1
   Next
   For j = 0 To iSaveMapcnt
      If strSaveMapName(j, 1) = cMapfile Then iMap = j
   Next
   If iSaveMapLast > 0 Then
      i = 0
     For j = strSaveMapName(iSaveMapLast, 2) To strSaveMapName(iSaveMapLast, 3)
      strSaveMap(j) = Map(i)
      i = i + 1
     Next
   End If
   '
   ' Read the map file - Map.txt
   '
   If iMap = 0 Then
        cFile = "c:\Kids\Quest\" & cMapfile & ".txt"
         Open cFile For Input As #1 ' Open file for input.
        i = 0
        iSaveMapcnt = iSaveMapcnt + 1
        iSaveMapLast = iSaveMapcnt
        strSaveMapName(iSaveMapcnt, 1) = cMapfile
        strSaveMapName(iSaveMapcnt, 2) = iSaveMapLoc
        strSaveMapName(iSaveMapcnt, 3) = iSaveMapLoc
        Do While Not EOF(1) ' Loop until end of file.
         Line Input #1, strMap
         If i < 99 And strMap <> "" Then
            Map(i) = Trim(strMap) ' Read data into two variables.
            jBun = 0
GetNextBun:
            jBun = InStr(jBun + 1, Map(i), " ")
            If jBun > 0 Then
                iBunny = iBunny + 1
               iBunnyloc(iBunny, 1) = i
               iBunnyloc(iBunny, 2) = jBun
               GoTo GetNextBun
            End If
            strSaveMapName(iSaveMapcnt, 3) = iSaveMapLoc
            strSaveMap(iSaveMapLoc) = strMap
            iSaveMapLoc = iSaveMapLoc + 1
              i = i + 1
         End If
         Loop
        Close #1    ' Close file.
  Else
    i = 0
     iSaveMapLast = iMap
    For j = strSaveMapName(iMap, 2) To strSaveMapName(iMap, 3)
      Map(i) = strSaveMap(j)
      jBun = 0
GetNextBun2:
       jBun = InStr(jBun + 1, Map(i), " ")
         If jBun > 0 Then
            iBunny = iBunny + 1
           iBunnyloc(iBunny, 1) = i
           iBunnyloc(iBunny, 2) = jBun
         GoTo GetNextBun2
        End If
     i = i + 1
    Next
  End If
  '
  ' Read the Maps.txt file for speech
  '
  For j = 1 To 90
   iJar(1, j) = 0
   iJar(2, j) = 0
  Next
  iJarLast = 1
  j = Rnd * 400
  jquest = 0
  iJarCnt = 1 + (j Mod 4)
  Dim Tag As String
  Dim bThought As Boolean
  Dim bSpeech As Boolean
  Dim bADD As Boolean
  Dim iCnt As Integer
  bSpeech = False
  bThought = False
  bADD = False
  Set nThought = New Collection
  Set nSpeech = New Collection
  cFile = "c:\Kids\Quest\" & cMapfile & "s.txt"
  Open cFile For Input As #1 ' Open file for input.
   i = 0
   Line Input #1, Tag
   Do While Not EOF(1) ' Loop until end of file.
     Line Input #1, Tag ' Read data into two variables.
     '
     ' Replace # with character name
     '
     Tag = Replace(Tag, "#", strChrName)
     If bSpeech Then
       If Mid(Tag, 1, 1) = "<" Then
         If bADD Then
            nSpeech.Add ClsSpeech
            If jquest > 0 Then
               MakeQuest(jquest) = ClsSpeech.QuestName
               QuestMap(jquest) = cMapfile
           End If
         End If
         Set ClsSpeech = New ClsSpeech
         jquest = 0
         ClsSpeech.Clear
         ClsSpeech.RNumber = Tag
         ClsSpeech.DoQuest = " "
         ClsSpeech.EndQuest = "0"
         ClsSpeech.QuestYes = "0"
         ClsSpeech.QuestNo = "0"
         ClsSpeech.Need = " "
         ClsSpeech.Give = " "
         ClsSpeech.FixX = "0"
         ClsSpeech.FixY = "0"
         ClsSpeech.BombQty = 0
         ClsSpeech.GiveQty = 0
         ClsSpeech.NeedQty = 0
         ClsSpeech.SayOnce = "N"
         ClsSpeech.TakeAny = "N"
         ClsSpeech.Status = ""
         iCnt = 0
         bADD = True
       Else
        If InStr(1, Tag, ":") > 0 Then
           j = InStr(1, Tag, ":")
          ClsSpeech.Name = Trim(Mid(Tag, 1, j))
          ClsSpeech.Talk = Mid(Tag, j + 1, Len(Tag) - j)
        ElseIf InStr(1, Tag, "=") Then
           j = Len(ClsSpeech.Name)
          'Tag = Replace(Tag, "@", Mid(ClsSpeech.Name, 1, j - 1))
          ClsSpeech.Question(iCnt) = Tag
          iCnt = iCnt + 1
        ElseIf UCase(Mid(Tag, 1, 8)) = "DOQUEST!" Then
          ClsSpeech.DoQuest = Right(Tag, Len(Tag) - 8)
          jquest = Val(ClsSpeech.DoQuest)
        ElseIf UCase(Mid(Tag, 1, 4)) = "WIN!" Then
          ClsSpeech.Status = "WIN"
        ElseIf UCase(Mid(Tag, 1, 5)) = "LOSE!" Then
          ClsSpeech.Status = "LOSE"
        ElseIf UCase(Mid(Tag, 1, 9)) = "ENDQUEST!" Then
          ClsSpeech.EndQuest = Right(Tag, Len(Tag) - 9)
        ElseIf UCase(Mid(Tag, 1, 5)) = "FIXX!" Then
          j = Val(Right(Tag, Len(Tag) - 5)) + 3
          ClsSpeech.FixX = j
        ElseIf UCase(Mid(Tag, 1, 5)) = "FIXY!" Then
          j = Val(Right(Tag, Len(Tag) - 5)) + 3
          ClsSpeech.FixY = j
        ElseIf UCase(Mid(Tag, 1, 8)) = "SAYONCE!" Then
           ClsSpeech.SayOnce = "Y"
       ElseIf UCase(Mid(Tag, 1, 8)) = "TAKEANY!" Then
           ClsSpeech.TakeAny = "Y"
        ElseIf UCase(Mid(Tag, 1, 8)) = "NEEDQTY!" Then
          ClsSpeech.NeedQty = Val(Right(Tag, Len(Tag) - 8))
        ElseIf UCase(Mid(Tag, 1, 9)) = "NEEDITEM!" Then
          ClsSpeech.Need = UCase(Right(Tag, Len(Tag) - 9))
        ElseIf UCase(Mid(Tag, 1, 10)) = "QUESTNAME!" Then
          ClsSpeech.QuestName = UCase(Right(Tag, Len(Tag) - 10))
        ElseIf UCase(Mid(Tag, 1, 5)) = "GIVE!" Then
          ClsSpeech.Give = UCase(Right(Tag, Len(Tag) - 5))
        ElseIf UCase(Mid(Tag, 1, 8)) = "GIVEQTY!" Then
          ClsSpeech.GiveQty = Val(Right(Tag, Len(Tag) - 8))
        ElseIf UCase(Mid(Tag, 1, 9)) = "QUESTYES!" Then
          ClsSpeech.QuestYes = Val(Right(Tag, Len(Tag) - 9))
        ElseIf UCase(Mid(Tag, 1, 8)) = "QUESTNO!" Then
          ClsSpeech.QuestNo = Val(Right(Tag, Len(Tag) - 8))
        ElseIf UCase(Mid(Tag, 1, 8)) = "BOMBQTY!" Then
          ClsSpeech.BombQty = Val(Right(Tag, Len(Tag) - 8))
        Else
        End If
       End If
    Else
       If bThought Then
         If UCase(Tag) = "<SPEECH>" Then
             bSpeech = True
             If bADD Then nThought.Add ClsThought
             bADD = False
          ElseIf UCase(Mid(Tag, 1, 5)) = "NAME!" Then
               If bADD Then
                  nThought.Add ClsThought
                  Set ClsThought = New ClsThought
                   ClsThought.Pence = ""
                    ClsThought.Person = ""
                    ClsThought.Thought = ""
                    ClsThought.HideThought = False
                    ClsThought.Waitt = False
                    ClsThought.Waitp = False
                    ClsThought.Givep = ""
                    ClsThought.Givet = ""
              End If
               bADD = True
              ClsThought.Person = Trim(Right(Tag, Len(Tag) - 5))
            ElseIf UCase(Mid(Tag, 1, 6)) = "PENCE!" Then
              ClsThought.Pence = Right(Tag, Len(Tag) - 6)
            ElseIf UCase(Mid(Tag, 1, 8)) = "THOUGHT!" Then
              ClsThought.Thought = Right(Tag, Len(Tag) - 8)
            ElseIf UCase(Mid(Tag, 1, 12)) = "HIDETHOUGHT!" Then
              ClsThought.HideThought = True
            ElseIf UCase(Mid(Tag, 1, 6)) = "WAITT!" Then
              ClsThought.Waitt = True
            ElseIf UCase(Mid(Tag, 1, 6)) = "WAITP!" Then
              ClsThought.Waitp = True
'            ElseIf UCase(Mid(Tag, 1, 6)) = "GIVET!" Then
'              ClsThought.Givet = UCase(Right(Tag, Len(Tag) - 6))
'            ElseIf UCase(Mid(Tag, 1, 6)) = "GIVEP!" Then
'              ClsThought.Givep = UCase(Right(Tag, Len(Tag) - 6))
            End If
       Else
       If UCase(Tag) = "<SPEECH>" Then
         bSpeech = True
         If bADD Then nThought.Add ClsThought
         bADD = False
       ElseIf UCase(Tag) = "<THOUGHT>" Then
         bThought = True
         Set ClsThought = New ClsThought
          ClsThought.Pence = ""
          ClsThought.Person = ""
          ClsThought.Thought = ""
          ClsThought.Waitt = False
          ClsThought.HideThought = False
          ClsThought.Waitt = False
          ClsThought.Waitp = False
          ClsThought.Givep = ""
          ClsThought.Givet = ""
       Else
        message(i) = Tag
        i = i + 1
       End If
       End If
    End If
'      If Tag = "<SPEECH>" Then bSpeech = True
    Loop
    'Message(i) = "0000000000000000aaa"
   Close #1
   iMsgCnt = i - 1
   If bADD Then
        nSpeech.Add ClsSpeech
         If jquest > 0 Then
           MakeQuest(jquest) = ClsSpeech.QuestName
          QuestMap(jquest) = cMapfile
         End If
   End If
End Sub
Public Sub SaveMapFile(cMapfile As String)
'
' Save MapFile
'
   Dim i As Integer
   Dim MyFile As String
   Dim Myfilein As String
   Myfilein = "c:\Kids\Quest\" & cMapfile & "s.txt"
   cMapfile = "c:\Kids\Quest\" & cMapfile & ".txt"
   Open cMapfile For Output As #1 ' Open file for output.
   For i = 0 To 100
      Print #1, Map(i) '& Chr(13) & Chr(10)
   Next
   Close #1    ' Close file.
 '
 'If needed create a speech file
 '
 MyFile = Dir(Myfilein)
  If Len(MyFile) < 4 Then
     Open Myfilein For Output As #1 ' Open file for output.
      Print #1, "  " '& Chr(13) & Chr(10)
      Close #1
   End If
 End Sub
Public Sub SelectPlace()
  'I really dont have idea of improving this.
  
  'Enter
  If PositionMap = "#" Then
    Call Door1
    OutsideMidi = Midi
    Midi = "House.mid"
    Call InitMusic
  End If
  
  'Exit
  If PositionMap = "M" Then
    Call Exit1
    Midi = OutsideMidi
    Call InitMusic
    ItemTownFound1 = False
    ItemTownFound2 = False
    ItemTownFound3 = False
  End If
  
  'Star Gate
  If PositionMap = "@" Then
    Call StarGate1
    Midi = "Town2.mid"
    Call InitMusic
  End If
  
End Sub
Public Sub door3(strHouse As String)
   ' Dim i As Integer
     Dim j As Integer
    'ItemHouseFound1 = False
    'ItemHouseFound2 = False
    'ItemHouseFound3 = False
    strHouse = UCase(strHouse)
    SAVE_MapLoaded = strHouse
    strMainMap = SAVE_MapLoaded
    j = InStr(1, strHouse, "H")
    
    'If bMainMapLoaded = True Then
      ' bMainMapLoaded = False
       '
       ' House Loc
       '
    If j > 0 Then
       MeLoaded = strHouse
       ReadMapFile (strHouse)
       CharY = CharY + 10
    Else
      bMainMapLoaded = True
      MeLoaded = strMainMap
      ReadMapFile (SAVE_MapLoaded)
      CharY = CharY - 10
    End If
End Sub
Public Sub door4(strHouse As String)
   ' Dim i As Integer
      SAVE_MapLoaded = strHouse
      bMainMapLoaded = True
      strMainMap = SAVE_MapLoaded
      MeLoaded = strMainMap
      ReadMapFile (SAVE_MapLoaded)
      'CharY = CharY - 3
    'End If
End Sub
Public Sub Door1()

  If MapLoaded = "A1" Then
    If CharY = 7 Then
      If CharX = 10 Then
        Call Map_A2
        MapLoaded = "A2"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 7 Then
      If CharX = 29 Then
        Call Map_A3
        MapLoaded = "A3"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 7 Then
      If CharX = 35 Then
        Call Map_A4
        MapLoaded = "A4"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
  End If
  
  If MapLoaded = "B1" Then
    If CharY = 7 Then
      If CharX = 7 Then
        Call Map_B2
        MapLoaded = "B2"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 10 Then
      If CharX = 8 Then
        Call Map_B3
        MapLoaded = "B3"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 8 Then
      If CharX = 36 Then
        Call Map_B4
        MapLoaded = "B4"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 9 Then
      If CharX = 21 Then
        Call Map_B5
        MapLoaded = "B5"
        CharY = 20
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
  End If
  
  If MapLoaded = "B1" Then
    If CharY = 23 Then
      If CharX = 59 Then
        Dim Final(10) As String
        Final(1) = "Guess what theres no temple!! This is the last version of this game."
        Final(2) = "Now you can make your RPG game."
        Final(3) = "And one last thing. If you make a game and you got inspired by this one"
        Final(4) = "try putting me on your credits :P"
        Final(5) = "Thanks for downloading this code :p"
        MsgBox Final(1) & vbCrLf & Final(2) & vbCrLf & Final(3) & vbCrLf & Final(4) & vbCrLf & Final(5) & vbCrLf & "Andres Zacarias" & vbCrLf & "Zacarias Software" & "www.zacarias.mainpage.net"
      End If
    End If
  End If
  
End Sub

Public Sub Exit1()

  If MapLoaded = "A2" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 10
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "A3" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 29
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "A4" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 35
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "B2" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_B1
        MapLoaded = "B1"
        CharY = 7
        CharX = 7
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "B3" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_B1
        MapLoaded = "B1"
        CharY = 10
        CharX = 8
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "B4" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_B1
        MapLoaded = "B1"
        CharY = 8
        CharX = 36
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "B5" Then
    If CharY = 20 Then
      If CharX = 7 Then
        Call Map_B1
        MapLoaded = "B1"
        CharY = 9
        CharX = 21
        ItemTownFound1 = False
      End If
    End If
  End If

End Sub

Public Sub StarGate1()
  If MapLoaded = "A1" Then
    If CharY = 10 Then
      If CharX = 64 Then
        Call Map_B1
        MapLoaded = "B1"
        SpeechLoaded = "B1"
        CharY = 41
        CharX = 15
        ItemTownFound1 = False
      End If
    End If
  End If
End Sub
