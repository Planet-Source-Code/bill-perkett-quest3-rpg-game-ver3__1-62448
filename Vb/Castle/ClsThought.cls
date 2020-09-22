VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsThought"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cEventThought As String
Private cEventPence As String
Private cEventPerson As String
Private cEventHideThought As Boolean
Private cEventWaitt As Boolean
Private cEventWaitp As Boolean
Private cEventGivet As String
Private cEventGivep As String
Public Property Get Givep() As String
    Givep = cEventGivep
End Property

Public Property Let Givep(cGivep As String)
    cEventGivep = cGivep
End Property
Public Property Get Givet() As String
    Givet = cEventGivet
End Property

Public Property Let Givet(cGivet As String)
    cEventGivet = cGivet
End Property
Public Property Get Waitp() As Boolean
    Waitp = cEventWaitp
End Property

Public Property Let Waitp(cWaitp As Boolean)
    cEventWaitp = cWaitp
End Property

Public Property Get Waitt() As Boolean
    Waitt = cEventWaitt
End Property

Public Property Let Waitt(cWaitt As Boolean)
    cEventWaitt = cWaitt
End Property
Public Property Get HideThought() As Boolean
    HideThought = cEventHideThought
End Property

Public Property Let HideThought(cHideThought As Boolean)
    cEventHideThought = cHideThought
End Property

Public Property Get Person() As String
    Person = cEventPerson
End Property

Public Property Let Person(cPerson As String)
    cEventPerson = cPerson
End Property
Public Property Get Pence() As String
    Pence = cEventPence
End Property

Public Property Let Pence(cPence As String)
    cEventPence = cPence
End Property
Public Property Get Thought() As String
    Thought = cEventThought
End Property

Public Property Let Thought(cThought As String)
    cEventThought = cThought
End Property
