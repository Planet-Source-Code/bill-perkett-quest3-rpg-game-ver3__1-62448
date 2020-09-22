Attribute VB_Name = "modSpeech"
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long


Public Sub InitializeSpeech()
  
  If CharFacing = 1 Then
    If MapLoaded = "A1" Then
      Call Character_A1
      SpeechLoaded = "A1"
    End If
    If MapLoaded = "A3" Then
      Call Character_A3
      SpeechLoaded = "A1"
    End If
    If MapLoaded = "A4" Then
      Call Character_A4
      SpeechLoaded = "A1"
    End If
    If MapLoaded = "B1" Then
      Call Character_B1
      SpeechLoaded = "B1"
    End If
    If MapLoaded = "B2" Then
      Call Character_B2
      SpeechLoaded = "B1"
    End If
    If MapLoaded = "B3" Then
      Call Character_B3
      SpeechLoaded = "B1"
    End If
    If MapLoaded = "B4" Then
      Call Character_B4
      SpeechLoaded = "B1"
    End If
    If MapLoaded = "B5" Then
      Call Character_B5
      SpeechLoaded = "B1"
    End If
  End If
  
End Sub


Public Sub Speech_A1()

  message(0) = "Sheik's House."
  message(1) = "To make spells you will need magic, i think theres some magic in the house."
  message(2) = "Hey brother, the White Wizard whants to talk to you, you can find him in the mini island."
  message(3) = "Yep. I can fix this bridge but i will need some wood. Bring it to me and i will do the rest."
  message(4) = "You whont be able to leave this island without magic. You should find the black wizards hidden in the island's to learn new spells."
  message(5) = "You have learned the Cut spell. Now you can use this spell by pressing the Z key. Now head east to start your Quest."
  message(6) = "Feew!!, this weed is to hard. I wish someone could help me."
  message(7) = "To the mountin of Wizards."
  message(8) = "To mini Island."
  message(9) = "To Centurion Island."
 message(10) = "Lot of weed grows in the bridge. Only those with special powers will leave this island."
 message(11) = "You found some wood! This thing doesnt has any use to you but you could try to give it to someone."
 message(12) = "Good. I see that you got wood, ill fix this bridge in no time."
 message(13) = "Nothing in there."
 message(14) = "There should be some wood in one of those jars, go ahead and take a peek."
 message(15) = "My housband left the island to look for a job, but he hasnt come back. Im really worried about this."
 message(16) = "You found a Coin! With coins you can buy things to get equiped."
 message(17) = "Thanks Sheik, i really needed your help. Please recive 10 coins as my gratitud."
 message(18) = "Thanks Sheik, i really needed your help."
 message(19) = "Star Gate to Centurion Island."
 message(20) = "You found a bag of Magic! With this your magic increments by 10 and you can make spells."
 message(21) = "You got no magic, you need to get some more to make your spells."
 
End Sub

Public Sub Speech_B1()

  message(0) = "Good to see you. In this island you will find the first temple, learn the spell and then you will be able to enter to the temple."
  message(1) = "Entering the forest of kings."
  message(2) = "Lost Village."
  message(3) = "This village is lost. No one has ever found us except for the Black Wizard and you!!"
  message(4) = "It wont be easy to get with the Black Wizard, there are some obstalces you must reach first to get to him."
  message(5) = "How the hell you got here!!. O well you are welcome any ways to the Lost Village."
  message(6) = "Strange. If you look to the right you will see another room, but i dont know how to get there."
  message(7) = "Me and my troops got lost on the forest and by luck we were rescue by the villagers from this town. As our gratitud we offer them our protection."
  message(8) = "I wonder how the king is."
  message(9) = "Were are you heading to??. Ohhh!!. You are the chosen one, you whont be able to pass to the temple without something to prove it."
 message(10) = "Ahhhh!! im so hungry, i whant some food."
 message(11) = "You found some wood! This thing doesnt has any use to you but you could try to give it to someone."
 message(12) = "Give me 50 coins and i will use my sword to cut that tree thats blocking your way."
 message(13) = "Nothing in there."
 message(14) = "Thanks!!"
 message(15) = "Hey. Give me 20 of wood and i will fix the hole of water in the floor."
 message(16) = "You found a Coin! With coins you can buy things to get equiped."
 message(17) = "Well thanks for cutting the weeds. Just kidding, you have now learned the Destroy Spell, now you can destroy hills!!. Press the X for this."
 message(18) = "That rock is not leting my weeds grow. Take it off and i will give you something."
 message(19) = "Thanks for removing it. Take this Bread as my gratitud."
 message(20) = "You found a bag of Magic! With this your magic increments by 10 and you can make spells."
 message(21) = "You got no magic, you need to get some more to make your spells."
 message(22) = "Temple of Glory."
 message(23) = "Only royal family may pass to the Temple of Glory."
 message(24) = "Hey i got my own shop but i cant sell you anything right now. I wonder why nobody comes!!."
 message(25) = "Thanks for removing it."
 message(26) = "Wow, thanks kid for the bread!!. Here you will need this Gold Ticket to pass to the Temple of Glory."
 message(27) = "Ohh!!. This ticket is from the lost troop!!. So they are alive!!, i will let you pass to so you can tell the king. You may stay with the ticket as a prove."
 message(28) = "ZzzZzz... ZzzZZzz..."
 message(29) = "Thats good, the troop is alive!!!."
 message(30) = "How did you find them??."
 
End Sub

Public Sub Character_A1()

  If CharY = 7 Then
    If CharX = 8 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(0)
    End If
  End If
  If CharY = 9 Then
    If CharX = 13 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Mother: " & message(1)
    End If
  End If
  If CharY = 8 Then
    If CharX = 6 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Brother: " & message(2)
    End If
  End If
  If CharY = 18 Then
    If CharX = 9 Then
      frmGame.TxtSpeech.Visible = True
      If Wood > 0 Then
        frmGame.TxtSpeech.Text = "Carpinter: " & message(12)
        Mid(Map(25), 13, 1) = "G"
        Wood = Wood - 1
      Else
        frmGame.TxtSpeech.Text = "Carpinter: " & message(3)
      End If
    End If
  End If
  If CharY = 19 Then
    If CharX = 30 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "White Wizard: " & message(4)
    End If
  End If
  If CharY = 35 Then
    If CharX = 31 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Black Wizard: " & message(5)
      SpellCut = True
    End If
  End If
  If CharY = 10 Then
    If CharX = 32 Then
      If Mid(Map(13), 30, 1) = "G" Then
        If ItemTownFound1 = False Then
          frmGame.TxtSpeech.Visible = True
          frmGame.TxtSpeech.Text = "Lady: " & message(17)
          Coin = Coin + 10
          ItemTownFound1 = True
        Else
          frmGame.TxtSpeech.Visible = True
          frmGame.TxtSpeech.Text = "Lady: " & message(18)
        End If
      Else
        frmGame.TxtSpeech.Visible = True
        frmGame.TxtSpeech.Text = "Lady: " & message(6)
        
      End If
    End If
  End If
  If CharY = 18 Then
    If CharX = 11 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(7)
    End If
  End If
  If CharY = 11 Then
    If CharX = 21 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(8)
    End If
  End If
  If CharY = 8 Then
    If CharX = 38 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(9)
    End If
  End If
  If CharY = 10 Then
    If CharX = 43 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(10)
    End If
  End If
  If CharY = 9 Then
    If CharX = 59 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(19)
    End If
  End If
  
End Sub

Public Sub Character_A3()
  If CharY = 10 Then
    If CharX = 9 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Lady: " & message(14)
    End If
  End If
End Sub

Public Sub Character_A4()
  If CharY = 9 Then
    If CharX = 12 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Lady: " & message(15)
    End If
  End If
End Sub

Public Sub Character_B1()
  If CharY = 39 Then
    If CharX = 18 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "White Wizard: " & message(0)
    End If
  End If
  If CharY = 31 Then
    If CharX = 15 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(1)
    End If
  End If
  If CharY = 11 Then
    If CharX = 19 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(2)
    End If
  End If
  If CharY = 6 Then
    If CharX = 10 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Lady: " & message(3)
    End If
  End If
  If CharY = 7 Then
    If CharX = 5 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Lady: " & message(4)
    End If
  End If
  If CharY = 10 Then
    If CharX = 31 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Boy: " & message(5)
    End If
  End If
  If CharY = 32 Then
    If CharX = 61 Then
      frmGame.TxtSpeech.Visible = True
      If Mid(Map(34), 56, 1) = ">" Then
        frmGame.TxtSpeech.Text = "Boy: " & message(18)
      Else
        If ItemTownFound1 = False Then
          Toast = Toast + 1
          frmGame.TxtSpeech.Text = "Boy: " & message(19)
        Else
          frmGame.TxtSpeech.Text = "Boy: " & message(25)
        End If
      End If
    End If
  End If
  If CharY = 30 Then
    If CharX = 61 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "The sign reads: " & message(22)
    End If
  End If
  If CharY = 29 Then
    If CharX = 59 Then
      frmGame.TxtSpeech.Visible = True
      If Ticket < 1 Then
        frmGame.TxtSpeech.Text = "Soldier: " & message(23)
      Else
        frmGame.TxtSpeech.Text = "Soldier: " & message(27)
        Mid(Map(31), 61, 1) = "G"
      End If
    End If
  End If
  If CharY = 29 Then
    If CharX = 60 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Soldier: " & message(28)
    End If
  End If
  If CharY = 26 Then
    If CharX = 58 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Soldier: " & message(29)
    End If
  End If
  If CharY = 26 Then
    If CharX = 61 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Soldier: " & message(30)
    End If
  End If
  
End Sub

Public Sub Character_B2()
  If CharY = 10 Then
    If CharX = 9 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Lady: " & message(6)
    End If
  End If
End Sub

Public Sub Character_B3()
  If CharY = 10 Then
    If CharX = 6 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Soldier: " & message(7)
    End If
  End If
  If CharY = 6 Then
    If CharX = 17 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Soldier: " & message(8)
    End If
  End If
  If CharY = 6 Then
    If CharX = 19 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Soldier: " & message(9)
    End If
  End If
  If CharY = 6 Then
    If CharX = 21 Then
      frmGame.TxtSpeech.Visible = True
      If Toast < 1 Then
        frmGame.TxtSpeech.Text = "Soldier: " & message(10)
      Else
        Ticket = Ticket + 1
        Toast = Toast - 1
        frmGame.TxtSpeech.Text = "Soldier: " & message(26)
      End If
    End If
  End If

End Sub

Public Sub Character_B4()
  
  If CharY = 9 Then
    If CharX = 11 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "White Wizard: " & message(24)
    End If
  End If

End Sub

Public Sub Character_B5()

  If CharY = 19 Then
    If CharX = 27 Then
      frmGame.TxtSpeech.Visible = True
      If Coin < 50 Then
        frmGame.TxtSpeech.Text = "Soldier: " & message(12)
      Else
        Mid(Map(20), 13, 1) = "G"
        Coin = Coin - 50
        frmGame.TxtSpeech.Text = "Soldier: " & message(14)
      End If
    End If
  End If
  
  If CharY = 14 Then
    If CharX = 6 Then
      frmGame.TxtSpeech.Visible = True
      If Wood < 20 Then
        frmGame.TxtSpeech.Text = "Carpinter: " & message(15)
      Else
        Mid(Map(16), 15, 1) = "G"
        Coin = Coin - 50
        frmGame.TxtSpeech.Text = "Carpinter: " & message(14)
      End If
    End If
  End If
  
  If CharY = 7 Then
    If CharX = 6 Then
      frmGame.TxtSpeech.Visible = True
      frmGame.TxtSpeech.Text = "Black Wizard: " & message(17)
      SpellDestroy = True
    End If
  End If
  
End Sub
