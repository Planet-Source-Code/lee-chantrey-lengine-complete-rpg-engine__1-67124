VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Party(3) As clsBattlePlayer

Public Function PartyUbound() As Integer

    Dim I As Integer, lOut As Integer
    lOut = 0
    
    For I = 0 To 3
        If Party(I).Name <> "" Then
            lOut = lOut + 1
        End If
    Next
    
    '0 Counts
    PartyUbound = (lOut - 1)

End Function

Public Function AttachPlayer(Index As Integer, ByRef Player As clsBattlePlayer)
    Set Party(Index) = Player
End Function

Public Property Get Player(ByVal Index) As clsBattlePlayer
Attribute Player.VB_MemberFlags = "200"
    Set Player = Party(Index)
End Property
