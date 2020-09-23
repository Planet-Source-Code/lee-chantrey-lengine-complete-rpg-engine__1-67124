Attribute VB_Name = "CharBank"
Option Explicit

'This module retains character stats
'after a character leaves the party
'* sigh *

Private bpBank(15) As New clsBattlePlayer
Private iBankCount As Integer

Public Function DumpBank(ByRef bcIni As clsIniObj)

Dim I As Integer

    For I = 0 To 15
        BattleCharINI bpBank(I), I, bcIni, "BattleCharBank_"
    Next

End Function

Public Function RestoreBank(ByRef bcIni As clsIniObj)

Dim I As Integer

    For I = 0 To 15
        BattleINIChar bpBank(I), I, bcIni, "BattleCharBank_"
    Next

End Function

Public Function MoveOut(sId As String, ByRef pbBattleChar As clsBattlePlayer) As Boolean

Dim cIndex As Integer

    cIndex = charIndex(sId)

    If cIndex = -1 Then
        MoveOut = False

        Exit Function
    End If
    
    CopyBattleChar pbBattleChar, bpBank(cIndex)

    'Move to end
    If cIndex < iBankCount Then
        CopyBattleChar bpBank(cIndex), bpBank(iBankCount - 1)
    End If
    'Set pbBattleChar = bpBank(cIndex)
    Set bpBank(iBankCount - 1) = New clsBattlePlayer

    iBankCount = iBankCount - 1

    MoveOut = True

End Function

Public Function MoveIn(ByRef bpBattleChar As clsBattlePlayer)
    
    If iBankCount = 15 Then
        MsgBox "Character bank exceeds 16.", vbCritical
        Exit Function
    End If
    
    CopyBattleChar bpBank(iBankCount), bpBattleChar
    Set bpBattleChar = New clsBattlePlayer
    
    iBankCount = iBankCount + 1
    
End Function

Private Function charIndex(sId As String) As Integer

Dim I As Integer

    sId = LCase(sId)
    
    For I = 0 To UBound(bpBank)
        If bpBank(I).ID = sId Then
            charIndex = I
            Exit Function
        End If
    Next
    
    charIndex = -1

End Function
