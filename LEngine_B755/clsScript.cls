VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event onCast(sSpell As String, Target)
Public Event onAttack(Target)

Dim iIndex As Integer

Public Property Let Index(newIndex As Integer)
    iIndex = newIndex
End Property

Public Property Get Index() As Integer
    Index = iIndex
End Property

Public Function Cast(sSpell As String, Target)
    RaiseEvent onCast(sSpell, Target)
End Function

Public Function Attack(Target)
    RaiseEvent onAttack(Target)
End Function
