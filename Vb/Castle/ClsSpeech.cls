VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private cEventRNumber As String
Private cEventName As String
Private mElements As Integer
Private mStrArray(20) As String
Private cEventTalk As String
Private cEventGive As String
Private cEventNeed As String
Private cEventDoQuest As String
Private cEventEndQuest As String
Private cEventFixX As String
Private cEventFixY As String
Private cEventNeedQty As Integer
Private cEventQuestName As String
Private cEventGiveQty As Integer
Private cEventQuestYes As String
Private cEventQuestNo As String
Private cEventBombQty As Integer
Private cEventSayOnce As String
Private cEventTakeAny As String
Private cEventStatus As String
Public Property Get Status() As String
    Status = cEventStatus
End Property

Public Property Let Status(cStatus As String)
    cEventStatus = cStatus
End Property
Public Property Get BombQty() As Integer
    BombQty = cEventBombQty
End Property

Public Property Let BombQty(cBombQty As Integer)
    cEventBombQty = cBombQty
End Property
Public Property Get SayOnce() As String
    SayOnce = cEventSayOnce
End Property

Public Property Let SayOnce(cSayOnce As String)
    cEventSayOnce = cSayOnce
End Property
Public Property Get TakeAny() As String
    TakeAny = cEventTakeAny
End Property

Public Property Let TakeAny(cTakeAny As String)
    cEventTakeAny = cTakeAny
End Property
Public Property Get QuestNo() As String
    QuestNo = cEventQuestNo
End Property

Public Property Let QuestNo(cQuestNo As String)
    cEventQuestNo = cQuestNo
End Property
Public Property Get QuestYes() As String
    QuestYes = cEventQuestYes
End Property

Public Property Let QuestYes(cQuestYes As String)
    cEventQuestYes = cQuestYes
End Property

Public Property Get GiveQty() As Integer
    GiveQty = cEventGiveQty
End Property

Public Property Let GiveQty(cGiveQty As Integer)
    cEventGiveQty = cGiveQty
End Property


Public Property Get QuestName() As String
    QuestName = cEventQuestName
End Property

Public Property Let QuestName(cQuestName As String)
    cEventQuestName = cQuestName
End Property
Public Property Get NeedQty() As Integer
    NeedQty = cEventNeedQty
End Property

Public Property Let NeedQty(cNeedQty As Integer)
    cEventNeedQty = cNeedQty
End Property
Public Property Get FixY() As String
    FixY = cEventFixY
End Property

Public Property Let FixY(cFixY As String)
    cEventFixY = cFixY
End Property
Public Property Get FixX() As String
    FixX = cEventFixX
End Property

Public Property Let FixX(cFixX As String)
    cEventFixX = cFixX
End Property
Public Property Get EndQuest() As String
    EndQuest = cEventEndQuest
End Property

Public Property Let EndQuest(cEndQuest As String)
    cEventEndQuest = cEndQuest
End Property
Public Property Get DoQuest() As String
    DoQuest = cEventDoQuest
End Property

Public Property Let DoQuest(cDoQuest As String)
    cEventDoQuest = cDoQuest
End Property

Public Property Get Need() As String
    Need = cEventNeed
End Property

Public Property Let Need(cNeed As String)
    cEventNeed = cNeed
End Property

Public Property Get Give() As String
    Give = cEventGive
End Property

Public Property Let Give(cGive As String)
    cEventGive = cGive
End Property

Public Property Get Talk() As String
    Talk = cEventTalk
End Property

Public Property Let Talk(cTalk As String)
    cEventTalk = cTalk
End Property

Public Property Get Name() As String
    Name = cEventName
End Property

Public Property Let Name(cName As String)
    cEventName = cName
End Property
Public Property Get RNumber() As String
    RNumber = cEventRNumber
End Property

Public Property Let RNumber(cRNumber As String)
    cEventRNumber = cRNumber
End Property
Public Property Get Question(Index As Integer) As String
    If Index < 20 Then
        Question = mStrArray(Index)
    End If
End Property
' Assign a new value to the specified item
Public Property Let Question(Index As Integer, newValue As String)
    If Index < 20 Then
        mStrArray(Index) = newValue
        If Index >= mElements Then mElements = Index + 1
    End If
End Property
Public Property Get Count() As Long
    Count = mElements
End Property
' Erase the array
Public Sub Clear()
    mElements = 0
End Sub

