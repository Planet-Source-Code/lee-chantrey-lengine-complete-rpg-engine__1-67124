Attribute VB_Name = "modDumpLoad"
Option Explicit

Function BattleINIChar(ByRef BattleChar As clsBattlePlayer, Index As Integer, ByRef bcIni As clsIniObj, Optional sSection As String = "BattleChar_")

    bcIni.Section = sSection & Index

    With BattleChar
         .Alive = bcIni.Read("Alive")
         .ATB = bcIni.Read("ATB")
        
         .Defence = bcIni.Read("Defence")
         .Strength = bcIni.Read("Strength")
        
         .Element = bcIni.Read("Element")
         .Experience = bcIni.Read("Experience")
         .Hp = bcIni.Read("Hp")
         .MaxHp = bcIni.Read("MaxHP")
         .Mp = bcIni.Read("Mp")
         .Name = bcIni.Read("Name")
    End With

End Function

Function BattleCharINI(ByRef BattleChar As clsBattlePlayer, Index As Integer, ByRef bcIni As clsIniObj, Optional sSection As String = "BattleChar_")

    bcIni.Section = sSection & Index

    With BattleChar
        bcIni.WriteData .Alive, "Alive"
        bcIni.WriteData .ATB, "ATB"
        
        bcIni.WriteData .Defence, "Defence"
        bcIni.WriteData .Strength, "Strength"
        
        bcIni.WriteData .Element, "Element"
        bcIni.WriteData .Experience, "Experience"
        bcIni.WriteData .Hp, "Hp"
        bcIni.WriteData .MaxHp, "MaxHP"
        bcIni.WriteData .Mp, "Mp"
        bcIni.WriteData .Name, "Name"
    End With

End Function

Public Function ReadMemory()

Dim bcIni As New clsIniObj, I As Integer, Scen As usrScen

    If FileExist(sPath_Save & "Dump.ini") = False Then
        WarnUser "System:ReadMemory:INIRead Error(Stream invalid)", True
    
        Exit Function
    End If

    bcIni.File = sPath_Save & "Dump.ini"

'Restore Story stuff
    modStory.RetrieveFunctions bcIni

    bcIni.Section = "Story"

'Restore Story-Characters
    With frmMain.Scen
        .ScenPath = sPath_Resources & bcIni.Read("ScenPath")
        .RestoreChars bcIni
    End With
    
    Exit Function
    
    With modStory.StoryFile
        .OpenStream sPath_Resources & bcIni.Read("Current")
        .SkipTo bcIni.Read("Line")
    End With
    
'Restore Story-Audio
    modMCI.RestoreAliases bcIni
    
'Restore Battle-Characters
    For I = 0 To C_MAX_PLAYERS
        BattleINIChar pBattleHu(I), I, bcIni
    Next
    CharBank.RestoreBank bcIni
    
    frmMain.Battle.RestoreParty bcIni

'Restore Inventory
    bcIni.Section = "Inventory"
    RestoreSuperCollection bcIni, Inventory.Types

End Function

Public Function RestoreSuperCollection(ByRef bcIni As clsIniObj, ByRef SuperCol As SuperCollection)

Dim sKeys() As String, I As Integer

    Set SuperCol = New SuperCollection
    sKeys = Split(bcIni.Read("Keys"), ",")
    
    For I = 0 To UBound(sKeys)
        SuperCol.Add bcIni.Read(sKeys(I)), sKeys(I)
    Next

End Function

Public Function DumpSuperCollection(ByRef bcIni As clsIniObj, ByRef SuperCol As SuperCollection)

Dim I As Integer, sKeys As String

    For I = 1 To SuperCol.Count - 1
        bcIni.WriteData SuperCol.GetItem(I), SuperCol.Key(I)
        sKeys = sKeys & SuperCol.Key(I) & ","
    Next
    
    If SuperCol.Count > 0 Then
        bcIni.WriteData SuperCol.GetItem(SuperCol.Count), SuperCol.Key(SuperCol.Count)
        sKeys = sKeys & SuperCol.Key(SuperCol.Count)
    End If
    
    bcIni.WriteData sKeys, "Keys"

End Function

Public Function DumpMemory()
    
    Dim bcIni As New clsIniObj, I As Integer
    bcIni.File = sPath_Save & "Dump.ini"
    
'Dump Story Stuff
    bcIni.Section = "Story"
    bcIni.WriteData KillHome(modStory.StoryFile.Path), "Current"
    bcIni.WriteData CStr(modStory.StoryFile.Position), "Line"
    bcIni.WriteData frmMain.Scen.ScenPath, "ScenPath"
    
    frmMain.Scen.Dump bcIni
    modStory.DumpFunctions bcIni
    
    modMCI.DumpAliases bcIni
    
'Dump Char Stats
    For I = 0 To C_MAX_PLAYERS
        BattleCharINI pBattleHu(I), I, bcIni
    Next
    CharBank.DumpBank bcIni
    
    frmMain.Battle.DumpParty bcIni

'Dump Inventory
    bcIni.Section = "Inventory"
    DumpSuperCollection bcIni, Inventory.Types
    
End Function
