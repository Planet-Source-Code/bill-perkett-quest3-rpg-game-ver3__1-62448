Attribute VB_Name = "modSaveLoad"
  Public SAVE_MapLoaded As String
  Public SAVE_Midi As String
  Public SAVE_OutsideMidi As String
  Public SAVE_SpeechLoaded As String
  Public SAVE_CharX As Integer
  Public SAVE_CharY As Integer
  Public SAVE_CharFacing As Integer
  Public SAVE_Wood As Integer
  Public SAVE_Bomb As Integer
  Public SAVE_Coin As Integer
  Public SAVE_Magic As Integer
  Public SAVE_Ticket As Integer
  Public SAVE_Toast As Integer
  Public SAVE_SpellCut As String
  Public SAVE_SpellDestroy As String
  Public SAVE_SpellWade As String
  Public SAVE_SpellLight As String
  Public SAVE_SpellAxe As String
  
  

Public Sub SaveGame()
  'SAVE_MapLoaded = MapLoaded
  SAVE_CharX = CharX
  SAVE_CharY = CharY
  SAVE_CharFacing = CharFacing
  SAVE_Wood = Wood
  SAVE_Coin = Coin
  SAVE_Magic = Magic
  SAVE_Ticket = Ticket
  SAVE_Toast = Toast
  SAVE_Bomb = Bomb
  SAVE_SpellCut = SpellCut
  SAVE_SpellDestroy = SpellDestroy
  SAVE_SpellWade = SpellWade
  SAVE_SpellLight = SpellLight
  SAVE_SpellAxe = SpellAxe
   Dim i As Integer
   Dim j As Integer
   i = 0
     For j = strSaveMapName(iSaveMapLast, 2) To strSaveMapName(iSaveMapLast, 3)
      strSaveMap(j) = Map(i)
      i = i + 1
     Next
  Open "c:\Kids\Quest\" & strChrName & ".qst" For Output As 1
  Write #1, iBunnyCaught, SAVE_MapLoaded, SAVE_CharX, SAVE_CharY, SAVE_CharFacing, SAVE_Wood, SAVE_Coin, SAVE_Magic, SAVE_Ticket, SAVE_Toast, SAVE_Bomb, SAVE_SpellCut, SAVE_SpellDestroy, SAVE_SpellWade, SAVE_SpellLight, SAVE_SpellAxe
  Write #1, strInventory
  Write #1, iSaveMapcnt, iSaveMapLast, iSaveMapLoc
  For i = 1 To 90
   Write #1, MyQuest(i)
  Next
  Write #1, iSaveMapcnt
  For i = 0 To iSaveMapcnt
   Write #1, strSaveMapName(i, 1), strSaveMapName(i, 2), strSaveMapName(i, 3)
  Next
  Write #1, iSaveMapLoc
  For i = 0 To iSaveMapLoc
   Write #1, strSaveMap(i)
  Next
   Close #1
  MsgBox strChrName & " - Game Saved"
  
End Sub

Public Sub LoadGame()
  Dim i As Integer
  Open "c:\Kids\Quest\" & strChrName & ".qst" For Input As 1
  Input #1, iBunnyCaught, SAVE_MapLoaded, SAVE_CharX, SAVE_CharY, SAVE_CharFacing, SAVE_Wood, SAVE_Coin, SAVE_Magic, SAVE_Ticket, SAVE_Toast, SAVE_Bomb, SAVE_SpellCut, SAVE_SpellDestroy, SAVE_SpellWade, SAVE_SpellLight, SAVE_SpellAxe
  Input #1, strInventory
  Input #1, iSaveMapcnt, iSaveMapLast, iSaveMapLoc
  For i = 1 To 90
    Input #1, MyQuest(i)
  Next
    Input #1, iSaveMapcnt
  For i = 0 To iSaveMapcnt
    Input #1, strSaveMapName(i, 1), strSaveMapName(i, 2), strSaveMapName(i, 3)
  Next
   Input #1, iSaveMapLoc
  For i = 0 To iSaveMapLoc
   Input #1, strSaveMap(i)
  Next
    iSaveMapLast = 0
     MapLoaded = SAVE_MapLoaded
    CharX = SAVE_CharX
    CharY = SAVE_CharY
    CharFacing = SAVE_CharFacing
    Wood = SAVE_Wood
    Coin = SAVE_Coin
    Magic = SAVE_Magic
    Ticket = SAVE_Ticket
    Toast = SAVE_Toast
    Bomb = SAVE_Bomb
    SpellCut = SAVE_SpellCut
    SpellDestroy = SAVE_SpellDestroy
    SpellWade = SAVE_SpellWade
    SpellLight = SAVE_SpellLight
    SpellAxe = SAVE_SpellAxe
   Close #1
    'FrmMake.Show
    
End Sub

