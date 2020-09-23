Attribute VB_Name = "modBattle"
Public Function StealTest(iSubject As Integer) As Boolean

Dim iRand As Integer

    StealTest = False

    iRand = RandomNumber(100, 1)

    If iSubject = 0 Then
        Exit Function
    End If
    
    If iRand <= iSubject Then
        StealTest = True
        Exit Function
    End If

End Function

Public Function StatLang(sSrc As String) As String

    sSrc = Replace(sSrc, "def", "Defence")
    sSrc = Replace(sSrc, "str", "Strength")
    sSrc = Replace(sSrc, "spr", "Spirit")
    sSrc = Replace(sSrc, "mag", "Magic")
    sSrc = Replace(sSrc, "atb", "ATB")
    
    StatLang = sSrc

End Function

Public Function PositionChars(sIniPath As String, Optional bDefaults As Boolean = False)
    'Position Characters on Battle Field

Dim I As Integer, iHum As New clsIniObj, iAI As New clsIniObj

    If FileExist(sIniPath) = False Then
        If bDefaults = False Then
            PositionChars sPath_Defaults & StrEnd(sIniPath, "\"), True
        End If
        
        Exit Function
    End If

    On Error Resume Next

    iHum.File = sIniPath
    iAI.File = sIniPath

    For I = 0 To C_MAX_PLAYERS
    
        If iHum.Read("X", pBattleHu(I).ID) <> "" Then
            With pBattleHu(I)
                .X = Evaluate(iHum.Read("X", pBattleHu(I).ID))
                .Y = Evaluate(iHum.Read("Y", pBattleHu(I).ID))
            End With
        Else
            iHum.Section = "Player " & CStr(I)
        
            With pBattleHu(I)
                .X = Evaluate(iHum.Read("X"))
                .Y = Evaluate(iHum.Read("Y"))
            End With
        End If
        
        If iAI.Read("X", pBattleAI(I).ID) <> "" Then
            With pBattleAI(I)
                .X = Evaluate(iAI.Read("X", pBattleAI(I).ID))
                .Y = Evaluate(iAI.Read("Y", pBattleAI(I).ID))
            End With
        Else
            iAI.Section = "Enemy " & CStr(I)
            
            With pBattleAI(I)
                .X = Evaluate(iAI.Read("X"))
                .Y = Evaluate(iAI.Read("Y"))
            End With
        End If
    Next

End Function
