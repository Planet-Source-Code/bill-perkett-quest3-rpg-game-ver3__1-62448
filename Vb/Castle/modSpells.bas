Attribute VB_Name = "modSpells"
Public Sub Transport(iMove As Integer)
  Dim i As Integer
  If iMagicPoints > 34 + iMove * 10 Then
   If Magic > 4 + iMove Then
    If CharFacing = 1 Then 'Up
      PositionMap = Mid(Map(CharY + 3 - iMove), CharX + 3, 1)
      If Mid(Map(CharY + 2), CharX + 3, 1) = "G" Then
          CharY = CharY - iMove
          'Mid(Map(CharY + 2), CharX + 3, 1) = "#"
         ' Mid(Map(CharY + 3 - iMove), CharX + 3, 1) = "~"
        Magic = Magic - 4 - iMove
       End If
    End If
  
    If CharFacing = 2 Then 'Down
      If Mid(Map(CharY + 3 + iMove), CharX + 3, 1) = "G" Then
       ' Mid(Map(CharY + 4), CharX + 3, 1) = "G"
        'Mid(Map(CharY + 2), CharX + 3, 1) = "#"
           CharY = CharY + iMove
        Magic = Magic - 4 - iMove
      End If
    End If
  
    If CharFacing = 3 Then 'Right
      If Mid(Map(CharY + 3), CharX + 3 + iMove, 1) = "G" Then
         CharX = CharX + iMove
        Magic = Magic - 4 - iMove
      End If
    End If
  
    If CharFacing = 4 Then 'Left
      If Mid(Map(CharY + 3), CharX + 3 - iMove, 1) = "G" Then
        CharX = CharX - iMove
        Magic = Magic - 4 - iMove
      End If
    End If
  Else
   If strMyForm = "FrmMake" Then
    FrmMake.txtSpeech.Visible = True
    FrmMake.Frame2.Visible = False
    FrmMake.txtSpeech.Text = "You need more magic to make your spells."
   Else
    frmGame.txtSpeech.Visible = True
    frmGame.txtSpeech.Text = "You have no magic, you need to get some more to make your spells."
  End If
  End If
 Else
     FrmMake.txtSpeech.Visible = True
     FrmMake.Frame2.Visible = False
     FrmMake.txtSpeech.Text = "Your Level of skill is too low."
  End If
End Sub
Public Sub Make_GrassGrow()
  If iMagicPoints > 34 Then
   If Magic > 4 Then
    If CharFacing = 1 Then 'Up
      PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
      If PositionMap = Chr(34) Then
        Mid(Map(CharY + 2), CharX + 3, 1) = "G"
        Magic = Magic - 4
       End If
    End If
  
    If CharFacing = 2 Then 'Down
      PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
      If PositionMap = Chr(34) Then
        Mid(Map(CharY + 4), CharX + 3, 1) = "G"
        Magic = Magic - 4
      End If
    End If
  
    If CharFacing = 3 Then 'Right
      PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
      If PositionMap = Chr(34) Then
        Mid(Map(CharY + 3), CharX + 4, 1) = "G"
        Magic = Magic - 4
      End If
    End If
  
    If CharFacing = 4 Then 'Left
      PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
      If PositionMap = Chr(34) Then
        Mid(Map(CharY + 3), CharX + 2, 1) = "G"
        Magic = Magic - 4
      End If
    End If
  Else
   If strMyForm = "FrmMake" Then
    FrmMake.txtSpeech.Visible = True
    FrmMake.Frame2.Visible = False
    FrmMake.txtSpeech.Text = "You need more magic to make your spells."
   Else
    frmGame.txtSpeech.Visible = True
    frmGame.txtSpeech.Text = "You have no magic, you need to get some more to make your spells."
  End If
  End If
 Else
     FrmMake.txtSpeech.Visible = True
     FrmMake.Frame2.Visible = False
     FrmMake.txtSpeech.Text = "Your Level of skill is too low."
  End If
End Sub
Public Sub Make_Spell_Cut()
  
  If Magic > 0 Then
    If CharFacing = 1 Then 'Up
      PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
      If PositionMap = "<" Then
        Mid(Map(CharY + 2), CharX + 3, 1) = "G"
        Magic = Magic - 1
        Weed = Weed + 4
      End If
       If PositionMap = "p" Then
         Mid(Map(CharY + 2), CharX + 3, 1) = "?"
         Magic = Magic - 1
         Weed = Weed + 3
      End If
    End If
  
    If CharFacing = 2 Then 'Down
      PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
      If PositionMap = "<" Then
        Mid(Map(CharY + 4), CharX + 3, 1) = "G"
        Magic = Magic - 1
         Weed = Weed + 4
      End If
      If PositionMap = "p" Then
         Mid(Map(CharY + 4), CharX + 3, 1) = "?"
         Magic = Magic - 1
         Weed = Weed + 3
      End If
    End If
  
    If CharFacing = 3 Then 'Right
      PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
      If PositionMap = "<" Then
        Mid(Map(CharY + 3), CharX + 4, 1) = "G"
        Magic = Magic - 1
         Weed = Weed + 4
      End If
      If PositionMap = "p" Then
         Mid(Map(CharY + 3), CharX + 4, 1) = "?"
         Magic = Magic - 1
         Weed = Weed + 3
      End If
    End If
  
    If CharFacing = 4 Then 'Left
      PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
      If PositionMap = "<" Then
        Mid(Map(CharY + 3), CharX + 2, 1) = "G"
        Magic = Magic - 1
         Weed = Weed + 4
      End If
      If PositionMap = "p" Then
         Mid(Map(CharY + 3), CharX + 2, 1) = "?"
         Magic = Magic - 1
         Weed = Weed + 3
      End If
    End If
  Else
   If strMyForm = "FrmMake" Then
    FrmMake.txtSpeech.Visible = True
    FrmMake.Frame2.Visible = False
    FrmMake.txtSpeech.Text = "You need more magic to make your spells."
   Else
    frmGame.txtSpeech.Visible = True
    frmGame.txtSpeech.Text = "You have no magic, you need to get some more to make your spells."
  End If
  End If
    
End Sub
Public Function OpenJarDoor() As String
  
  
    If CharFacing = 1 Then 'Up
      PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
      If PositionMap = "'" Then
        Mid(Map(CharY + 2), CharX + 3, 1) = "?"
       End If
      
    End If
  
    If CharFacing = 2 Then 'Down
      PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
      If PositionMap = "'" Then
        Mid(Map(CharY + 4), CharX + 3, 1) = "?"
       
      End If
      
    End If
  
    If CharFacing = 3 Then 'Right
      PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
      If PositionMap = "'" Then
        Mid(Map(CharY + 3), CharX + 4, 1) = "?"
       
      End If
     
    End If
  
    If CharFacing = 4 Then 'Left
      PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
      If PositionMap = "'" Then
        Mid(Map(CharY + 3), CharX + 2, 1) = "?"
       
      End If
     
    End If
  OpenJarDoor = PositionMap
    
End Function

Public Sub Make_Spell_Destroy()

  If Magic > 4 Then
    If CharFacing = 1 Then 'Up
      PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
      If PositionMap = ">" Then
        Mid(Map(CharY + 2), CharX + 3, 1) = "G"
        Magic = Magic - 5
         Rock = Rock + 3
      End If
    End If
  
    If CharFacing = 2 Then 'Down
      PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
      If PositionMap = ">" Then
        Mid(Map(CharY + 4), CharX + 3, 1) = "G"
        Magic = Magic - 5
        Rock = Rock + 3
      End If
    End If
  
    If CharFacing = 3 Then 'Right
      PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
      If PositionMap = ">" Then
        Mid(Map(CharY + 3), CharX + 4, 1) = "G"
        Magic = Magic - 5
        Rock = Rock + 3
      End If
    End If
  
    If CharFacing = 4 Then 'Left
      PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
      If PositionMap = ">" Then
        Mid(Map(CharY + 3), CharX + 2, 1) = "G"
        Magic = Magic - 5
        Rock = Rock + 3
      End If
    End If
  Else
   If strMyForm = "FrmMake" Then
    FrmMake.txtSpeech.Visible = True
    FrmMake.Frame2.Visible = False
    FrmMake.txtSpeech.Text = "You need more magic to make your spells."
   Else
    frmGame.txtSpeech.Visible = True
    frmGame.txtSpeech.Text = "You have no magic, you need to get some more to make your spells."
  End If
  End If

End Sub
Public Sub Make_Spell_Axe()

  If Magic > 4 Then
    If CharFacing = 1 Then 'Up
      PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
      If PositionMap = "B" Then
        Mid(Map(CharY + 2), CharX + 3, 1) = "G"
        Magic = Magic - 4
       End If
    End If
  
    If CharFacing = 2 Then 'Down
      PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
      If PositionMap = "B" Then
        Mid(Map(CharY + 4), CharX + 3, 1) = "G"
        Magic = Magic - 4
      End If
    End If
  
    If CharFacing = 3 Then 'Right
      PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
      If PositionMap = "B" Then
        Mid(Map(CharY + 3), CharX + 4, 1) = "G"
        Magic = Magic - 4
      End If
    End If
  
    If CharFacing = 4 Then 'Left
      PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
      If PositionMap = "B" Then
        Mid(Map(CharY + 3), CharX + 2, 1) = "G"
        Magic = Magic - 4
      End If
    End If
  Else
   If strMyForm = "FrmMake" Then
    FrmMake.txtSpeech.Visible = True
    FrmMake.Frame2.Visible = False
    FrmMake.txtSpeech.Text = "You need more magic to make your spells."
   Else
    frmGame.txtSpeech.Visible = True
    frmGame.txtSpeech.Text = "You have no magic, you need to get some more to make your spells."
  End If
  End If

End Sub
Public Sub Make_Spell_Light()

  If Magic > 1 Then
    If CharFacing = 1 Then 'Up
      PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
      If PositionMap = "_" Then
        Mid(Map(CharY + 2), CharX + 3, 1) = "G"
        Magic = Magic - 2
       End If
    End If
  
    If CharFacing = 2 Then 'Down
      PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
      If PositionMap = "_" Then
        Mid(Map(CharY + 4), CharX + 3, 1) = "G"
        Magic = Magic - 2
      End If
    End If
  
    If CharFacing = 3 Then 'Right
      PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
      If PositionMap = "_" Then
        Mid(Map(CharY + 3), CharX + 4, 1) = "G"
        Magic = Magic - 2
      End If
    End If
  
    If CharFacing = 4 Then 'Left
      PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
      If PositionMap = "_" Then
        Mid(Map(CharY + 3), CharX + 2, 1) = "G"
        Magic = Magic - 2
      End If
    End If
  Else
   If strMyForm = "FrmMake" Then
    FrmMake.txtSpeech.Visible = True
    FrmMake.Frame2.Visible = False
    FrmMake.txtSpeech.Text = "You need more magic to make your spells."
   Else
    frmGame.txtSpeech.Visible = True
    frmGame.txtSpeech.Text = "You have no magic, you need to get some more to make your spells."
  End If
  End If

End Sub

Public Sub Make_Spell_Wade()

  If Magic > 2 Then
    If CharFacing = 1 Then 'Up
      PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
      If PositionMap = "+" Then
        Mid(Map(CharY + 2), CharX + 3, 1) = "G"
        Magic = Magic - 3
      End If
      If PositionMap = "`" Then
         Mid(Map(CharY + 2), CharX + 3, 1) = "?"
         Magic = Magic - 3
      End If
    End If
  
    If CharFacing = 2 Then 'Down
      PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
      If PositionMap = "+" Then
        Mid(Map(CharY + 4), CharX + 3, 1) = "G"
        Magic = Magic - 3
      End If
      If PositionMap = "`" Then
          Mid(Map(CharY + 4), CharX + 3, 1) = "?"
         Magic = Magic - 3
      End If
    End If
  
    If CharFacing = 3 Then 'Right
      PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
      If PositionMap = "+" Then
        Mid(Map(CharY + 3), CharX + 4, 1) = "G"
        Magic = Magic - 3
       End If
       If PositionMap = "`" Then
         Mid(Map(CharY + 3), CharX + 4, 1) = "?"
         Magic = Magic - 3
      End If
    End If
  
    If CharFacing = 4 Then 'Left
      PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
      If PositionMap = "+" Then
        Mid(Map(CharY + 3), CharX + 2, 1) = "G"
        Magic = Magic - 3
      End If
      If PositionMap = "`" Then
         Mid(Map(CharY + 3), CharX + 2, 1) = "?"
         Magic = Magic - 3
      End If
    End If
  Else
   If strMyForm = "FrmMake" Then
    FrmMake.Frame2.Visible = False
    FrmMake.txtSpeech.Visible = True
    FrmMake.txtSpeech.Text = "You need more magic to make your spells."
   Else
    frmGame.txtSpeech.Visible = True
    frmGame.txtSpeech.Text = "You have no magic, you need to get some more to make your spells."
  End If
  End If

End Sub

