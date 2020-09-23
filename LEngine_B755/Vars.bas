Attribute VB_Name = "Vars"
Public Const CD_Left As Integer = 1
Public Const CD_Right As Integer = 2
Public Const CD_Down As Integer = 3
Public Const CD_Up As Integer = 4

Public cSubOptions(5) As New SuperCollection
Public Inventory As New clsInventory

Public Aliases As New SuperCollection
'Public AliasesAdded As New Collection

Public cPosBattleChanges As New Collection
Public cNegBattleChanges As New Collection

Public Const C_MAX_PLAYERS As Integer = 3

Public pBattleHu(C_MAX_PLAYERS) As New clsBattlePlayer
Public pBattleAI(C_MAX_PLAYERS) As New clsBattlePlayer
Public iParty_Count As Integer

Public sPath_Resources As String
Public sPath_Save As String

Public sPath_Battle As String
Public sPath_BattleChars As String
Public sPath_Fiends As String
Public sPath_Bosses As String
Public sPath_Equipment As String

Public sPath_Story As String
Public sPath_StoryChars As String
Public sPath_StoryPlaces As String
Public sPath_Items As String
Public sPath_Defaults As String
Public sPath_System As String
Public sPath_Audio As String
Public sPath_Music As String
Public sPath_SFX As String

Public bInBattle As Boolean
Public DebugWin As New frmDebug
Public VarsWin As New frmVariables

Public sMusicCurrent As String

Public Const lTrans As Long = 8388863

Sub SetVars()

    sPath_Resources = App.Path & "\resources\"
    sPath_Save = App.Path & "\Save\"
    
    sPath_Items = sPath_Resources & "items\"
    
    sPath_Battle = sPath_Resources & "Battle\"
    sPath_BattleChars = sPath_Battle & "Characters\"
    
    sPath_Fiends = sPath_Battle & "Fiends\"
    sPath_Bosses = sPath_Battle & "Bosses\"
    
    sPath_Story = sPath_Resources & "Story\"
    sPath_StoryChars = sPath_Story & "Characters\"
    
    sPath_StoryPlaces = sPath_Story & "Places\"
    sPath_Defaults = sPath_Resources & "Defaults\"
    
    sPath_Equipment = sPath_Resources & "Equipment\"
    
    sPath_System = sPath_Resources & "System\"
    
    sPath_Audio = sPath_Resources & "Audio\"
    sPath_Music = sPath_Audio & "Music\"
    
    sPath_SFX = sPath_Battle & "SFX\"
    
    sMusicCurrent = "music"
    
    SetJoypad
    
    GetItemTypes

End Sub
