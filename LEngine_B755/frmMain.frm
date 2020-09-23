VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LEngine GDI+"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4815
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   StartUpPosition =   2  'CenterScreen
   Begin prjLEngine.usrName NameEntry 
      Height          =   3600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin prjLEngine.usrScen Scen 
      Height          =   3600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
      Begin prjLEngine.usrChar usrChar1 
         Left            =   240
         Top             =   600
         _ExtentX        =   423
         _ExtentY        =   423
      End
   End
   Begin prjLEngine.usrBattle Battle 
      Height          =   3600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6218
   End
   Begin prjLEngine.usrEquip Equip 
      Height          =   3600
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin prjLEngine.usrEquip Equip 
      Height          =   3600
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin prjLEngine.usrEquip Equip 
      Height          =   3600
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin prjLEngine.usrEquip Equip 
      Height          =   3600
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin prjLEngine.usrMainMenu MainMenu 
      Height          =   3615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuVideo 
         Caption         =   "&Video"
         Begin VB.Menu mnuNormal 
            Caption         =   "&Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFull 
            Caption         =   "&Full Screen"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuBattle 
      Caption         =   "&Battle"
      Begin VB.Menu mnuSpawn 
         Caption         =   "&Spawn Fiend"
      End
      Begin VB.Menu MnuSep232 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFiends 
         Caption         =   "&Fiends"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuBosses 
         Caption         =   "&Bosses"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuStory 
      Caption         =   "&Story"
      Begin VB.Menu mnuSpeech 
         Caption         =   "&Speech"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVoice 
         Caption         =   "&Voice"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuMemory 
      Caption         =   "&Memory"
      Begin VB.Menu mnuRead 
         Caption         =   "&Read"
      End
      Begin VB.Menu mnuDump 
         Caption         =   "&Write"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cFiends As New Collection
Private COpts As New Collection

Private sCharID As String
Private bResume As Boolean

Private bBattleOptional As Boolean
Private sWinCmd As String

Sub Start(Optional sStory As String = "main.sty")
    StartStory sStory
End Sub

Private Sub SetCorrectSize()
    'Seems to bring back titlebar ?
    '[!] I wrote it to correctly resize the window
    
    'Set to 320x240
    While Me.ScaleHeight < 240
        Me.Height = Me.Height + 15
    Wend
    
    While Me.ScaleWidth < 320
        Me.Width = Me.Width + 15
    Wend

End Sub

Sub BounceChar(CArgs As Collection)

    If CArgs.Count = 2 Then
        '
    ElseIf CArgs.Count = 1 Then
        CArgs.Add 1
    Else
        WarnUser "BounceChar CharID [Speed -ms]", True
        Exit Sub
    End If
    
    Scen.BounceChar CArgs(1), CArgs(2) + 8

End Sub

Sub BounceCharExclem(CArgs As Collection)

    If CArgs.Count = 2 Then
        '
    ElseIf CArgs.Count = 1 Then
        CArgs.Add 1
    Else
        WarnUser "BounceChar CharID [Speed -ms]", True
        Exit Sub
    End If
    
    Scen.BounceCharExclem CArgs(1), CArgs(2) + 8

End Sub

Sub FiendChance(CArgs As Collection)

    If CArgs.Count <> 2 Then
        WarnUser "FiendChance Steps FreeSteps"
        Exit Sub
    End If

    Scen.SetFiendChance CArgs(1), CArgs(2)

End Sub

Sub NamePrompt(CArgs As Collection)

    If CArgs.Count <> 3 Then
        WarnUser "NamePrompt CharID DefaultName Picture"
        Exit Sub
    End If

    sCharID = CArgs(1)
    NameEntry.SetPicture Vars.sPath_Story & CArgs(3)
    
    Scen.keys.Enabled = False

    NameEntry.keys.Enabled = True
    NameEntry.Visible = True
    NameEntry.CharName = CArgs(2)
    
    HaltStory

End Sub

Sub FrameProp(CArgs As Collection)

    If CArgs.Count <> 6 Then
        WarnUser "FrameProp Name Path Interval X Y Loop"
        Exit Sub
    End If

    Scen.FrameProp CArgs(1), CArgs(2), CArgs(3), CArgs(4), CArgs(5), CBol(CArgs(6))

End Sub

Sub Options(CArgs As Collection)

    If CArgs.Count < 1 Then
        WarnUser "Options Option 1 [Option 2].. "
        Exit Sub
    End If
    
    'Scen.Options CArgs
    Set COpts = CArgs

End Sub

Sub Answers(CArgs As Collection)

    If COpts.Count = 0 Then
        WarnUser "Options must be called before Answers"
        Exit Sub
    ElseIf CArgs.Count < 1 Or COpts.Count <> CArgs.Count Then
        WarnUser "There must be the same amount of Questions as there is answers"
        Exit Sub
    End If
    
    Scen.Answers COpts, CArgs
    
    Set COpts = New Collection
    Set CArgs = New Collection
    
    HaltStory

End Sub

Sub AmendInventory(CArgs As Collection)
'Items work differently then Equipment, their name is their ID

    If CArgs.Count <> 1 Then
        WarnUser "AmendInventory Item"
    End If

    Inventory.AddItem CArgs(1)
    Scen.AmendInventory CArgs(1), False

End Sub

Sub AmendEquipment(CArgs As Collection)
'Equipment's Name can be found in its file

Dim sPath As String, sName As String

    If CArgs.Count <> 1 Then
        WarnUser "AmendEquipment EquipmentFile"
        Exit Sub
    End If
    
    sPath = sPath_Equipment & CArgs(1)
    Equip(0).AddEquip sPath
    
    sName = ReadINIValue(sPath, "Visual", "Name", "????")
    
    Scen.AmendInventory sName, False

End Sub

Sub PlayAudio(CArgs As Collection)

    If CArgs.Count < 1 Then
        WarnUser "PlayAudio ID [Loops]"
        Exit Sub
        
    ElseIf CArgs.Count = 1 Then
        CArgs.Add "1"
    End If

    modMCI.PlayAudio CArgs(1), CBol(CArgs(2))

End Sub

Sub LoadMusic(CArgs As Collection)

    If CArgs.Count <> 2 Then
        WarnUser "LoadMusic ID Path"
        Exit Sub
    End If

    modMCI.LoadFile sPath_Music & CArgs(2), CArgs(1)

End Sub

Sub LoadAudio(CArgs As Collection)

    If CArgs.Count <> 2 Then
        WarnUser "LoadAudio ID Path"
        Exit Sub
    End If

    modMCI.LoadFile sPath_Audio & CArgs(2), CArgs(1)

End Sub

Sub SetMusicAlias(CArgs As Collection)
    
    If CArgs.Count <> 1 Then
        WarnUser "SetMusic New_Alias", True
        Exit Sub
    End If
    
    Vars.sMusicCurrent = CArgs(1)

End Sub

Sub StopMusic(CArgs As Collection)

    If CArgs.Count > 1 Then
        WarnUser "StopMusic [Alias]", True
        Exit Sub
        
    ElseIf CArgs.Count = 1 Then
        modMCI.StopAudio CArgs(1)
    Else
        modMCI.StopAudio sMusicCurrent
    End If

End Sub

Sub PauseMusic(CArgs As Collection)

    If CArgs.Count <> 1 Then
        WarnUser "PauseMusic Alias", True
        Exit Sub
    End If

    modMCI.PauseAudio sMusicCurrent

End Sub

Sub PlayMusic(CArgs As Collection)

    If CArgs.Count > 1 Then
        WarnUser "PlayMusic Path (Macro)"
        WarnUser "PlayMusic"
        
        Exit Sub
    End If
    
    If CArgs.Count > 0 Then
        'ID doesnt exist, so Arguement must be Path of Music
        modMCI.LoadFile sPath_Music & CArgs(1), sMusicCurrent
    End If
    
    modMCI.PlayAudio sMusicCurrent, True
    
End Sub

Sub SwitchMusic(CArgs As Collection)

    If CArgs.Count <> 1 Then
        WarnUser "SwitchMusic Alias", True
        Exit Sub
    End If
    
    modMCI.StopAudio sMusicCurrent
    Me.SetMusicAlias CArgs
    
    modMCI.PlayAudio sMusicCurrent, True

End Sub

Sub HideProp(CArgs As Collection)

    If CArgs.Count <> 1 Then
        WarnUser "HideProp Name"
        Exit Sub
    End If

    Scen.HideProp CArgs(1)
    
End Sub

Sub ShowProp(CArgs As Collection)

    If CArgs.Count <> 1 Then
        WarnUser "ShowProp Name"
        Exit Sub
    End If

    Scen.ShowProp CArgs(1)

End Sub

Sub SmallDoorY(CArgs As Collection)

    If CArgs.Count <> 4 Then
        WarnUser "SmallDoorY Name Path X Y"
        Exit Sub
    End If

    Scen.SmallDoorY CArgs(1), sPath_Story & CArgs(2), CArgs(3), CArgs(4)

End Sub

Sub StaticProp(CArgs As Collection)

    If CArgs.Count <> 4 Then
        WarnUser "StaticProp Name Path X Y"
        Exit Sub
    End If

    Scen.StaticProp CArgs(1), sPath_Story & CArgs(2), CArgs(3), CArgs(4)

End Sub

Sub ReplaceImg(CArgs As Collection)

    If CArgs.Count <> 2 Then
        WarnUser "ReplaceImg CharName Path"
        Exit Sub
    End If
    
    Scen.ReplaceImg CArgs(1), CArgs(2)

End Sub

Sub MaxStat(CArgs As Collection)

Dim charIndex As Integer, sStat As String

    If CArgs.Count <> 1 Then
        WarnUser "MaxStat HP/MP/Both", True
        Exit Sub
    End If
    
    sStat = LCase(CArgs(1))
    
    If sStat = "hp" Then
        
        For charIndex = 0 To C_MAX_PLAYERS
            pBattleHu(charIndex).TempStats = False
            pBattleHu(charIndex).Hp = pBattleHu(charIndex).MaxHp
        Next
        
        Scen.NotifyText "Party's HP was fully restored!"
        
    ElseIf sStat = "mp" Then
    
        For charIndex = 0 To C_MAX_PLAYERS
            pBattleHu(charIndex).TempStats = False
            pBattleHu(charIndex).Mp = pBattleHu(charIndex).MaxMp
        Next
        
        Scen.NotifyText "Party's MP was fully restored!"
    
    ElseIf sStat = "both" Then
    
        For charIndex = 0 To C_MAX_PLAYERS
            pBattleHu(charIndex).TempStats = False
            pBattleHu(charIndex).Mp = pBattleHu(charIndex).MaxMp
            pBattleHu(charIndex).Hp = pBattleHu(charIndex).MaxHp
        Next
        
        Scen.NotifyText "Party's HP and MP was fully restored!"
    
    Else
        WarnUser "MaxStat: Unrecognised parameters.", True
    End If

End Sub

Sub BgColour(CArgs As Collection)

    If CArgs.Count <> 1 Then
        WarnUser "BgColour #Colour"
        Exit Sub
    End If
    
    Scen.SetBGColour CArgs(1)

End Sub

Sub LeaveParty(CArgs As Collection)

Dim charIndex As Integer

    If CArgs.Count <> 1 Then
        WarnUser "LeaveParty Id"
        Exit Sub
    End If
    
    charIndex = FindBattleCharIndex(CArgs(1))

    If Battle.RemovePlayer(charIndex) = True Then
        iParty_Count = iParty_Count - 1
    
        'Move to End of party (if not end)
        If charIndex < iParty_Count Then
            Equip(charIndex).Char = Equip(iParty_Count).Char
        End If
            
        Equip(iParty_Count).Char = ""
    End If

End Sub

Sub AmendParty(CArgs As Collection)
        
Dim sId As String
        
    If CArgs.Count <> 1 Then
        WarnUser "AmendParty %CharPath%"
        Exit Sub
    End If
    
    sId = LCase(CArgs(1))

    If Battle.CreatePlayer(sId) = True Then
        Equip(iParty_Count).Char = sId
        
        iParty_Count = iParty_Count + 1
    Else
        WarnUser "AmendParty Failed:: Unkown", True
    End If
    
End Sub

Sub ReleaseFiend(CArgs As Collection)

    If CArgs.Count > 4 And CArgs.Count < 1 Then
        WarnUser "ReleaseFiend Path, Path2, etc"
        Exit Sub
    End If

    While CArgs.Count > 0
        cFiends.Add CArgs(1)
        CArgs.Remove 1
    Wend

End Sub

Sub IntAdd(CArgs As Collection)

    'Story Math
    If CArgs.Count <> 2 Then
        WarnUser "IntAdd %VariableName% AmountToAdd"
        Exit Sub
    End If

    Op_IntAdd CArgs(1), CArgs(2)

End Sub

Sub Variable(CArgs As Collection)

    'Story Math
    If CArgs.Count <> 2 Then
        WarnUser "Variable %VariableName% NewVariableValue"
        Exit Sub
    End If
    
    Op_Variable CArgs(1), CArgs(2)

End Sub

Sub StrCmp(CArgs As Collection)

Dim sFalse As String

    'Story Math
    If CArgs.Count < 3 Then
        WarnUser "StrCmp Value1 Value2 CommandIfEqual Optional(CommandIfNotEqual)"
        Exit Sub
    End If

    If CArgs.Count = 4 Then
        sFalse = CArgs(4)
    End If

    Op_StrCmp CArgs(1), CArgs(2), CArgs(3), sFalse

End Sub

'Sub AsyncDialog(CArgs As Collection)

    'Scenhandle
    'If CArgs.Count <> 2 Then
        'WarnUser "AsyncDialog Name Text"
        'Exit Sub
    'End If

    'Scen.AsyncDialog CArgs(1), CArgs(2)

'End Sub

Sub AsyncAction(CArgs As Collection)
    '(Replaces AsyncDialog)

    'Scenhandle
    If CArgs.Count <> 2 Then
        WarnUser "AsyncAction CharName Action"
        Exit Sub
    End If

    Scen.AsyncAction CArgs(1), CArgs(2)

End Sub

Sub AsyncCollision(CArgs As Collection)

    'Scenhandle
    If CArgs.Count <> 2 Then
        WarnUser "AsyncCollision CharName Command"
        Exit Sub
    End If

    Scen.AsyncCollision CArgs(1), CArgs(2)

End Sub

Sub Say(CArgs As Collection)

    'Scenhandle
    If CArgs.Count < 2 Then
        WarnUser "Say Name Text [Halt]"
        Exit Sub
    ElseIf CArgs.Count = 2 Then
        CArgs.Add "1"
    End If

    If mnuSpeech.Checked = True Then
        Scen.SayText CArgs(1), CArgs(2), CBol(CArgs(3))
    End If

End Sub

Sub StoryVoice(CArgs As Collection)

    If CArgs.Count < 1 Then
        WarnUser "StoryVoice TextToVoice [Halt 0/1]"
        Exit Sub
        
    ElseIf CArgs.Count = 1 Then
        CArgs.Add "1"
    End If

    If mnuVoice.Checked = True Then
        Scen.StoryVoice CArgs(1), CBol(CArgs(2))
    End If

End Sub

Sub Sleep(CArgs As Collection)

    'Scenhandle
    If CArgs.Count <> 1 Then
        WarnUser "Sleep Miliseconds"
        Exit Sub
    End If

    Scen.Sleep CArgs(1)

End Sub

Sub SleepUntilNext(CArgs As Collection)

    HaltStory

End Sub

Sub SetScen(CArgs As Collection)

    If CArgs.Count <> 1 Then
        WarnUser "SetScen ScenFile"
        Exit Sub
    End If

    'Clear Fiends
    Set cFiends = New Collection
    
    Scen.StartTransistion
    Scen.ScenPath = sPath_StoryPlaces & CArgs(1)
   
End Sub

Sub GotoStory(CArgs As Collection)
    If CArgs.Count <> 1 Then
        WarnUser "GotoStory StoryFile.sty"
        Exit Sub
    End If
    
    'StartNewStory CArgs(1)
    StartStory CArgs(1)
End Sub

Sub SetMainCharPos(CArgs As Collection)
    If CArgs.Count <> 3 Then
        WarnUser "SetMainCharPos X Y Direction"
        Exit Sub
    End If
        
    Scen.Start CArgs(1), CArgs(2), CArgs(3)
End Sub

Sub ReSkinChar(CArgs As Collection)

    If CArgs.Count <> 3 Then
        WarnUser "ReSkinChar ID Name Path"
        Exit Sub
    End If
    
    Scen.SkinChar Scen.getCharByName(CArgs(1)), CArgs(2), sPath_StoryChars & CArgs(3)

End Sub

Sub LoadNPC(CArgs As Collection)

    If CArgs.Count <> 3 Then
        WarnUser "LoadNPC ID Name Path"
        Exit Sub
    End If
    
    Scen.LoadNPC CArgs(1), CArgs(2), CArgs(3)
    
End Sub

Sub WalkCharX(CArgs As Collection)

    If CArgs.Count < 2 Then
        WarnUser "WalkCharX Name NewX [Speed] [Halt]"
        Exit Sub
    ElseIf CArgs.Count = 2 Then
        CArgs.Add "8"
        CArgs.Add "1"
    ElseIf CArgs.Count = 3 Then
    
        CArgs.Add "1"
    ElseIf CArgs.Count > 4 Then
        WarnUser "WalkCharX Name NewX [Speed] [Halt]"
    End If
        
    If 8 Mod CInt(CArgs(3)) <> 0 Then
        WarnUser "WalkCharX:: Speed must be divisable by 8"
        Exit Sub
    End If
        
    Scen.WalkCharX CArgs(1), CArgs(2), CArgs(3), CArgs(4)

End Sub

Sub WalkCharY(CArgs As Collection)

    If CArgs.Count < 2 Then
        WarnUser "WalkCharY Name NewY [Speed] [Halt]"
        Exit Sub
    ElseIf CArgs.Count = 2 Then
        CArgs.Add "8"
        CArgs.Add "1"
    ElseIf CArgs.Count = 3 Then
    
        CArgs.Add "1"
    ElseIf CArgs.Count > 4 Then
        WarnUser "WalkCharY Name NewY [Speed] [Halt]"
    End If
    
    If 8 Mod CArgs(3) <> 0 Then
        WarnUser "WalkCharY Speed must be divisable by 8"
        Exit Sub
    End If
        
    Scen.WalkCharY CArgs(1), CArgs(2), CArgs(3), CBol(CArgs(4))

End Sub

Sub EventSquare(CArgs As Collection)

    If CArgs.Count < 3 Then
        WarnUser "EventSquare X, Y, Command, [Expire 1/0] [Halt 1/0]"
        Exit Sub
    ElseIf CArgs.Count = 3 Then
        CArgs.Add "1"
        CArgs.Add "0"
    ElseIf CArgs.Count = 4 Then
        CArgs.Add "0"
    End If
    
    Scen.SetEventSquare CArgs(1), CArgs(2), CArgs(3), CBol(CArgs(4)), CBol(CArgs(5))

End Sub

Sub ActionSquare(CArgs As Collection)

    If CArgs.Count = 4 Then
        '
    ElseIf CArgs.Count = 3 Then
        CArgs.Add 1
    Else
        WarnUser "ActionSquare X Y Command [Visible]"
        Exit Sub
    End If
    
    Scen.SetActionSquare CArgs(1), CArgs(2), CArgs(3), CArgs(4)

End Sub

Sub SetCharDirection(CArgs As Collection)

    If CArgs.Count <> 2 Then
        WarnUser "SetCharDirection Name, Direction"
        Exit Sub
    End If
    
    Scen.LookDir CArgs(1), CArgs(2)

End Sub

Sub Control(CArgs As Collection)

    If CArgs.Count <> 1 Then
        WarnUser "Control 1/0"
        Exit Sub
    End If

    If CArgs(1) = "1" Then
        Player_Control = True
    ElseIf CArgs(1) = "0" Then
        Player_Control = False
    End If
    
End Sub

Sub HideChar(CArgs As Collection)

    If CArgs.Count <> 1 Then
        WarnUser "HideChar Name"
        Exit Sub
    End If

    Scen.HideChar CArgs(1)

End Sub

Sub ShowChar(CArgs As Collection)

    If CArgs.Count = 1 Then
        '
    ElseIf CArgs.Count = 4 Then
        Scen.PositionChar CArgs(1), CArgs(2), CArgs(3), CArgs(4)
        
    ElseIf CArgs.Count = 6 Then
        'Classic LoadChar

        Scen.LoadNPC CArgs(1), CArgs(2), CArgs(3)
        Scen.PositionChar CArgs(1), CArgs(4), CArgs(5), CArgs(6)
    Else
        WarnUser "ShowChar ID Name Path X Y Direction (Macro)"
        WarnUser "ShowChar ID X Y Direction (Macro)"
        WarnUser "ShowChar ID"
        
        Exit Sub
    End If
    
    Scen.ShowChar CArgs(1)

End Sub

Sub PositionChar(CArgs As Collection)

    If CArgs.Count <> 4 Then
        WarnUser "PositionChar Name X Y Direction"
        Exit Sub
    End If

    Scen.PositionChar CArgs(1), CArgs(2), CArgs(3), CArgs(4)

End Sub

Sub BossBattleOptional(CArgs As Collection)

    If CArgs.Count < 2 Then
        WarnUser "BossBattleOptional CommandIfWin, Path1, Path2, etc"
        Exit Sub
    End If
    
    sWinCmd = CArgs(1)
    CArgs.Remove 1

    BossBattle CArgs, True
    
End Sub

Sub BossBattle(CArgs As Collection, Optional bOptional As Boolean = False)

    If mnuBosses.Checked = False Then
        Exit Sub
    End If

    modMCI.StopAudio sMusicCurrent

    If Battle.PartyCount = 0 Then
        WarnUser "Boss Event Aborted:: No Party Members", False
        Exit Sub
    End If

    If CArgs.Count > 4 Or CArgs.Count < 1 Then
        WarnUser "BossBattle Path, Path2, etc"
        Exit Sub
    End If

    If Script_Halt = False Then
        bResume = True
        modStory.HaltStory
    End If

    'Initialise Screen
    Scen.Visible = False
    Scen.keys.Enabled = False
    
    bBattleOptional = bOptional
    
    Battle.Boss = True
    Battle.Visible = True

    Battle.LoadBattFile Scen.ScenPath
    Battle.PrepareField

    While CArgs.Count > 0
        Battle.CreateEnemy CArgs(1), True
        CArgs.Remove 1
    Wend

    Battle.StartATMs

End Sub

Sub Prologue(CArgs As Collection)

    If CArgs.Count < 1 Then
        WarnUser "Prologue TextFile [Picture] [Interval]"
        Exit Sub
        
    ElseIf CArgs.Count = 1 Then
        CArgs.Add ""
        CArgs.Add 70
    ElseIf CArgs.Count = 2 Then
        UpdateCol CArgs, 2, sPath_Story & CArgs(2)
        CArgs.Add 70
    Else
        UpdateCol CArgs, 2, sPath_Story & CArgs(2)
    End If

    Scen.SetPrologue sPath_Story & CArgs(1), CArgs(2), CArgs(3)

End Sub

Private Sub Battle_BattleFinished(iRes As Integer)

    If bBattleOptional = True And iRes = 0 Then
        ExeStoryCmd sWinCmd
        sWinCmd = ""
    End If

    If iRes = 0 Or bBattleOptional = True Then
        'Player Win / Lost in optional
    
        Scen.Visible = True
        Scen.PlayMusic
        Scen.SetFocus
        
        Scen.keys.Enabled = True
        
        If Battle.Boss = False Then
            Player_Control = True
        End If
            
        bInBattle = False
        
        If bResume = True Then
            bResume = False
            modStory.ResumeStory
        End If
        
    Else
        Unload Me
        
    End If

End Sub

Private Sub Equip_OnFinished(Index As Integer)
    
    With Equip(Index)
        .Visible = False
        .Enabled = False
        
        Player_Control = True
    End With
    
End Sub

Private Sub mnuAbout_Click()

    MsgBox "This is a Demonstration build!" & _
            vbCrLf & _
            "Build: " & CInt((App.Major * 100) + (App.Minor * 10) + (App.Revision)) & _
        vbCrLf & _
        vbCrLf & _
            "Programmed by Lee Matthew Chantrey", vbInformation
            
                                                                    
End Sub

Private Sub mnuBosses_Click()

    If mnuBosses.Checked = False Then
        mnuBosses.Checked = True
    Else
        mnuBosses.Checked = False
    End If

End Sub

Private Sub mnuDump_Click()
    'Read all memory values, story positions
    DumpMemory
    
    Scen.NotifyText "Memory Dumped!"

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFiends_Click()
    
    If mnuFiends.Checked = False Then
        mnuFiends.Checked = True
    Else
        mnuFiends.Checked = False
    End If
    
End Sub

Private Sub mnuFull_Click()
    If mnuFull.Checked = True Then
        Exit Sub
    End If
    
    mnuFull.Checked = True
    mnuNormal.Checked = False
    
    HideFileMenus
    Hide_Cursor
    
    FormOnTop hWnd

    TitleBar False
    ChangeToGameRes
    
    Me.Move 0, 0
End Sub

Private Sub mnuRead_Click()

    HaltStory
    
    Battle.RemoveAllPlayers
    Scen.ClearScen

    ReadMemory
    
    ResumeStory

End Sub

Private Sub mnuSpawn_Click()
    Scen_OnFiend
End Sub

Private Sub mnuSpeech_Click()
    If mnuSpeech.Checked = True Then
        mnuSpeech.Checked = False
    Else
        mnuSpeech.Checked = True
    End If
End Sub

Private Sub mnuVoice_Click()
    If mnuVoice.Checked = True Then
        mnuVoice.Checked = False
    Else
        mnuVoice.Checked = True
    End If
End Sub

Private Sub ShowFileMenus()

On Error Resume Next
    
    Dim Control
    For Each Control In Controls
        If Left(Control.Name, 3) = "mnu" Then
            Control.Visible = True
        End If
    Next

End Sub

Private Sub HideFileMenus()

On Error Resume Next
    
    Dim Control
    For Each Control In Controls
        If Left(Control.Name, 3) = "mnu" Then
            Control.Visible = False
        End If
    Next

End Sub

Private Sub mnuNormal_Click()

    If mnuNormal.Checked = True Then
        Exit Sub
    End If
    
    mnuNormal.Checked = True
    mnuFull.Checked = False

    TitleBar True
    SetOrigRes
    
    SetCorrectSize
    
End Sub

Public Sub ToggleFileMenu()

    If mnuFull.Checked = False Then
        Exit Sub
    End If

    If mnuFile.Visible = False Then
        ShowFileMenus
        Show_Cursor
    Else
        HideFileMenus
        Hide_Cursor
    End If

End Sub

Private Sub NameEntry_OnConfirmed(ByVal sName As String)
        
    'Prevent Deletion
    Aliases.Locked(sCharID) = True
    
    Aliases.Add sName, sCharID
    Op_Variable sCharID, sName
    
    NameEntry.Visible = False
    ResumeStory
    
    Scen.keys.Enabled = True
    NameEntry.keys.Enabled = False
End Sub

Private Sub Scen_OnEquip(charIndex As Integer)

    If pBattleHu(charIndex).Name = "" Then
        NotifyUser "F" & (charIndex + 1) & "- Disused/Expired Battle Character Index"
        Player_Control = True
        
        Exit Sub
    End If

    With Equip(charIndex)
        .ZOrder 0
        .Visible = True
        
        .RefreshStats
        .Enabled = True
    End With
        
End Sub

Private Sub Scen_OnFiend()

    If Battle.PartyCount = 0 Then
        WarnUser "Fiend Event Aborted:: No Party Members", False
        Exit Sub
    End If

    If cFiends.Count = 0 Or mnuFiends.Checked = False Then
        Exit Sub
    End If
    
    modMCI.StopAudio sMusicCurrent
    
    Scen.keys.Enabled = False
    
    Battle.LoadBattFile Scen.ScenPath
    
    Battle.PrepareField
    Battle.Boss = False

    Dim I As Integer, iFiendNo As Integer, iFiendType(3) As Integer
    'At least 2 Fiends
    iFiendNo = RandomNumber(3, 0)
    
    For I = 0 To iFiendNo
        iFiendType(I) = RandomNumber(cFiends.Count, 1)
        Battle.CreateEnemy cFiends(iFiendType(I))
    Next
    
    Scen.Visible = False
    bInBattle = True
    
    Battle.StartATMs
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If NotInDbg = True And _
        mnuNormal.Checked = False Then
        
        SetOrigRes
    End If
    
    bUnload = True
    
    Unload frmLib
    'Unload Known Forms
    Unload DebugWin
    Unload VarsWin

    End
End Sub
