VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.UserControl usrBattle 
   BackColor       =   &H00808080&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin prjLEngine.usrSprite SFX 
      Height          =   3600
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin VB.Timer timDied 
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   4200
      Top             =   960
   End
   Begin VB.Timer timTurnFinished 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer timNotify 
      Enabled         =   0   'False
      Interval        =   2200
      Left            =   0
      Top             =   120
   End
   Begin prjLEngine.usrTransPic imgSlash 
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      MaskColor       =   -2147483633
   End
   Begin prjLEngine.usrTransPic imgEnemy 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      MaskColor       =   -2147483633
   End
   Begin prjLEngine.usrTransPic imgNumbers 
      Height          =   120
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   120
      _ExtentX        =   212
      _ExtentY        =   212
      MaskColor       =   16777215
   End
   Begin prjLEngine.usrTransPic Char 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      MaskColor       =   -2147483633
   End
   Begin VB.Timer timWin 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1560
      Top             =   2280
   End
   Begin VB.Timer timFloat 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Tag             =   "0"
      Top             =   2280
   End
   Begin VB.Timer timATM 
      Enabled         =   0   'False
      Index           =   3
      Left            =   480
      Top             =   2280
   End
   Begin VB.Timer timATM 
      Enabled         =   0   'False
      Index           =   2
      Left            =   360
      Top             =   2280
   End
   Begin VB.Timer timATM 
      Enabled         =   0   'False
      Index           =   1
      Left            =   240
      Top             =   2280
   End
   Begin VB.Timer timATM 
      Enabled         =   0   'False
      Index           =   0
      Left            =   120
      Top             =   2280
   End
   Begin prjLEngine.usrMenu MnuEnemies 
      Height          =   720
      Left            =   75
      TabIndex        =   1
      Top             =   2805
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   1270
      Begin prjLEngine.usrMenu MnuActions 
         Height          =   735
         Left            =   1080
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1270
      End
      Begin prjLEngine.usrMenu MnuCustom 
         Height          =   735
         Left            =   600
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1270
      End
      Begin prjLEngine.usrCharMenu MnuCharacters 
         Height          =   705
         Left            =   2295
         TabIndex        =   2
         Top             =   0
         Width           =   2430
         _ExtentX        =   3863
         _ExtentY        =   1244
      End
      Begin VB.Image imgCur 
         Height          =   195
         Index           =   0
         Left            =   1800
         Picture         =   "usrBattle.ctx":0000
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin prjLEngine.keyReciever keys 
      Left            =   3000
      Top             =   1680
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin VB.Timer timShowAttack 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   2520
      Tag             =   "0"
      Top             =   1080
   End
   Begin VB.Timer timShowDamage 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   1080
   End
   Begin VB.Timer timAnim 
      Enabled         =   0   'False
      Index           =   0
      Left            =   2520
      Top             =   2280
   End
   Begin prjLEngine.usrMenu MnuItems 
      Height          =   720
      Left            =   75
      TabIndex        =   4
      Top             =   2805
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1270
   End
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   4200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin prjLEngine.usrNotify Notify 
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   423
   End
   Begin VB.Shape shpShadow 
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   600
      Shape           =   2  'Oval
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Numbers 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "0"
      Top             =   3360
      Width           =   150
   End
   Begin VB.Image imgCur 
      Height          =   195
      Index           =   1
      Left            =   0
      Picture         =   "usrBattle.ctx":03A0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBack 
      Height          =   3600
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "usrBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Mnu_Custom As Integer = 4
Private Const Mnu_Items As Integer = 3
Private Const Mnu_Characters As Integer = 2
Private Const Mnu_Enemies As Integer = 1
Private Const Mnu_Action As Integer = 0

Private Const Act_Attack As String = "A"
Private Const Act_Item As String = "I"
Private Const Act_Steal As String = "S"
Private Const Act_Custom As String = "C"

Private Const Sta_Attack As Integer = 10

Private pSlashSize(3) As picSize

Private iAniStart(40) As Integer
Private iAniEnd(40) As Integer

Private imenuState As Integer
Private sMnuAnswers(4) As String
Private cMnuHistory As New Collection

Private bTakingTurn As Boolean
Private bTurnFinished As Boolean
Private bAnimationBusy As Boolean

Private bFirstMember As Boolean 'Special Conditions for firsts
Private bFirstEnemy As Boolean '0 is already loaded
Private bMagic As Boolean

Private cNeedsActivation As New Collection

Private iPartyIndex As Integer
Private iEnemyIndex As Integer

Private iFloatState As Integer

Private bAttackAi As Boolean
Private bTargetAi As Boolean

Private iWinState As Integer
Private bBattleEnd As Boolean

Private CurrentMnu

Private imgKill(7)
Private b_imgKillUsed(7) As Boolean
Private b_imgKillAI(7) As Boolean
Private i_HuIndex(7) As Integer

Private bKill As Boolean

Private iSubjectIndex As Integer
Private iOffenderIndex As Integer

Private LastTarget As New clsBattlePlayer
Private LastOffender As New clsBattlePlayer

Private cBattleActions As New Collection
Private cCharsReady As New Collection
Private cNotifications As New Collection
Private bAttackFin As Boolean
Private sLastSpell As String

Private pParty(1) As New clsParty
Private iAiATM(3) As Integer

Private sSFXFinished As Boolean
Private bWaitNotify As Boolean

Public Event BattleFinished(iRes As Integer)
Private WithEvents ScriptDef As clsScript
Attribute ScriptDef.VB_VarHelpID = -1

Private bBoss As Boolean
Private sBattlePath As String

Public Property Get PartyCount() As Integer
    PartyCount = MnuCharacters.ListCount
End Property

Public Property Let Boss(bNewBoss As Boolean)
    bBoss = bNewBoss
End Property

Public Property Get Boss() As Boolean
    Boss = bBoss
End Property

Sub DigGrave(ByRef imgSrc, bAI As Boolean, Optional charIndex As Integer = -1)

Dim I As Integer

    For I = 0 To UBound(imgKill)
        If b_imgKillUsed(I) = False Then
            b_imgKillUsed(I) = True
            b_imgKillAI(I) = bAI
            i_HuIndex(I) = charIndex
            
            Set imgKill(I) = imgSrc
            Exit Sub
        End If
    Next
            
End Sub

Sub BerryDead()

Dim I As Integer

    For I = 0 To UBound(imgKill)
        If b_imgKillUsed(I) = True Then
        
            b_imgKillUsed(I) = False
        
            If b_imgKillAI(I) = True Then
                DealWithDeadAI i_HuIndex(I)
            Else
                DealWithDeadHuman i_HuIndex(I)
            End If

            b_imgKillAI(I) = False
        End If
    Next

End Sub

Sub DumpParty(ByRef bcIni As clsIniObj)

Dim I As Integer, sIds As String

    bcIni.Section = "Battle_Party"
    
    For I = 0 To MnuCharacters.ListUbound
        sIds = sIds & pBattleHu(I).ID & ","
    Next
    
    If Len(sIds) > 0 Then
        sIds = Mid(sIds, 1, Len(sIds) - 1)
    End If
    
    bcIni.WriteData sIds, "Members"
    
    bcIni.Section = "Battle_Aliases"
    DumpSuperCollection bcIni, Aliases

End Sub

Sub RestoreParty(ByRef bcIni As clsIniObj)

Dim I As Integer, sIds() As String

    bcIni.Section = "Battle_Aliases"
    RestoreSuperCollection bcIni, Aliases

    bcIni.Section = "Battle_Party"
    sIds = Split(bcIni.Read("Members"), ",")
    
    For I = 0 To UBound(sIds)
        Me.CreatePlayer sIds(I)
    Next

End Sub

Sub RemoveAllPlayers()

Dim I As Integer

    For I = 0 To MnuCharacters.ListUbound
        RemovePlayer I
    Next

End Sub

Sub UpdateCharDisplay(charIndex As Integer)

    If pBattleHu(charIndex).Hp > pBattleHu(charIndex).MaxHp Then
        pBattleHu(charIndex).Hp = pBattleHu(charIndex).MaxHp
    End If

    MnuCharacters.UpdateATB charIndex, pBattleHu(charIndex).ATB
    MnuCharacters.UpdateHP charIndex, pBattleHu(charIndex).Hp

End Sub

Sub NextNotify()

    '! I should only be called, when the previous notify is
    'gone [implying the notify object is invisisble]

    If cNotifications.Count > 0 Then
        NotifyAction cNotifications(1)
        cNotifications.Remove 1
    ElseIf bWaitNotify = True Then
        bWaitNotify = False
        timShowAttack.Enabled = True
    End If

End Sub

Sub NotifyAction(sCaption As String)
    
    If Notify.Visible = False Then
        Notify.Caption = sCaption
        Notify.Visible = True
        
        'Max visiable 2secs
        timNotify.Enabled = True
    Else
        'Que
        cNotifications.Add sCaption
    End If
    
End Sub

Sub InitialiseActions(BattleChar As clsBattlePlayer, Optional bCompare As Boolean)

    On Error GoTo Catch_E

Dim sPath As String, cActions As New Collection, sRes As String, I As Integer, actionIndex As Integer
Dim CActionNode As New Collection, sIniPath As String, iniMp As New clsIniObj, iniLvl As New clsIniObj

    sPath = Vars.sPath_BattleChars & BattleChar.ID & "\Actions\"
    
    With frmLib.Dir1
        .Path = sPath
    End With
    
    If frmLib.Dir1.ListCount > 0 Then
        For I = 0 To frmLib.Dir1.ListCount - 1
            sRes = StrEnd(frmLib.Dir1.List(I), "\")
            If sRes <> "False" Then
                cActions.Add sRes
            End If
        Next
        
        If bCompare = False Then
            'start an array
            BattleChar.Actions = cActions
        End If

        If cActions.Count > 0 Then
            actionIndex = 1
            
            iniMp.Key = "MP Cost"
            iniMp.Default = 0
            
            iniLvl.Key = "Level"
            iniLvl.Default = 0
            
            While actionIndex <= cActions.Count
                'sIniPath = sPath & cActions(actionIndex) & ".ini"
                Set CActionNode = BattleChar.ActionNode(actionIndex)
                
                iniMp.File = sPath & cActions(actionIndex) & ".ini"
                iniLvl.File = iniMp.File
        
                With frmLib.File1
                    .Path = sPath & cActions(actionIndex)
                    .Pattern = "*.def"
                End With
                
                For I = 0 To frmLib.File1.ListCount - 1
                    sRes = StrFront(frmLib.File1.List(I), ".")

                    If BattleChar.Level >= iniLvl.Read(, sRes) Then
                        If ColExists(CActionNode, iniMp.Key & "_" & sRes) = False Then
                            CActionNode.Add sRes & ":" & iniMp.Read(, sRes), iniMp.Key & "_" & sRes
                            
                            If bCompare = True Then
                                NotifyAction BattleChar.Name & " has learnt Spell '" & sRes & "'"
                            End If
                        End If
                    End If
                Next
                
                BattleChar.ActionNode(actionIndex) = CActionNode
                
                actionIndex = actionIndex + 1
            Wend
        End If
    End If
    
    Exit Sub
    
Catch_E:
    'MsgBox ":("
    
    'MsgBox pBattleHu(Index).Actions.Count

End Sub

Sub InitialiseItems()
'Lists all available items in item menu

Dim cTypes As SuperCollection

    MnuItems.ClearList

    Set cTypes = Inventory.Types
    
    Dim sItem As String, I As Integer
    
    For I = 1 To cTypes.Count
        If Inventory.TypeCount(cTypes.Key(I)) > 0 Then
            sItem = CStr(cTypes.Key(I))
            
            MnuItems.AddItem sItem & " x " & Inventory.TypeCount(sItem), sItem
        End If
    Next

End Sub

Sub InitialiseSteals(BattleChar As clsBattlePlayer)
'all available steals

Dim cSteals As New Collection, cIni As New clsIniObj, _
    iSlotIndex As Integer

    cIni.File = BattleChar.Path & "inventory.ini"

    For iSlotIndex = 1 To 16
        cIni.Section = "Slot " & CStr(iSlotIndex)
        
        If cIni.Read("Name") <> "" Then
            cSteals.Add cIni.Read("Type") & ":" & cIni.Read("Name") & ":" & Replace(cIni.Read("Steal Chance"), "%", "") & ":" & Replace(cIni.Read("Drop Chance"), "%", "")
        End If
    Next
    
    BattleChar.Steals = cSteals

End Sub

Sub InitialiseKeys()

    'Notified on these key changes
    keys.AddKey Control_Cancel
    keys.AddKey Control_Select
    
    keys.AddKey Control_Down
    keys.AddKey Control_Up
    keys.AddKey Control_Right
    keys.AddKey Control_Left

End Sub

Private Function RunScript(sStatement As String)

    On Error Resume Next

    Script.ExecuteStatement sStatement
    If Err Then
        With Script.Error
            MsgBox "Syntax Error : " & .Number & ": " & .Description & " at line " & .Line & " column " & .Column & ": " & vbCrLf, vbCritical, "Script Error"
        End With
    End If

End Function

Private Function SetScript(ByRef pTarget As clsBattlePlayer, ByRef pOfender As clsBattlePlayer, Optional bSpell As Boolean = False)
        
    Set LastTarget = pTarget
    Set LastOffender = pOfender
    
    'Setup variables in script
    If pOfender.Enemy = False Then
        Script.ExecuteStatement "Set Offender = HumanBattle_" & pOfender.Index
        Script.ExecuteStatement "Set OffensiveParty = HumanParty"
        
        SFX.ExecuteStatement "Set OffensiveParty = HumanParty"
        SFX.ExecuteStatement "Set Offender = HumanBattle_" & pOfender.Index
    Else
        Script.ExecuteStatement "Set Offender = AIBattle_" & pOfender.Index
        Script.ExecuteStatement "Set OffensiveParty = AiParty"
        
        SFX.ExecuteStatement "Set OffensiveParty = AiParty"
        SFX.ExecuteStatement "Set Offender = AIBattle_" & pOfender.Index
    End If
    
    If pTarget.Enemy = False Then
        Script.ExecuteStatement "Set Target = HumanBattle_" & pTarget.Index
        Script.ExecuteStatement "Set TargetParty = HumanParty"
        
        SFX.ExecuteStatement "Set TargetParty = HumanParty"
        SFX.ExecuteStatement "Set Target = HumanBattle_" & pTarget.Index
    Else
        Script.ExecuteStatement "Set Target = AIBattle_" & pTarget.Index
        Script.ExecuteStatement "Set TargetParty = AiParty"
        
        SFX.ExecuteStatement "Set TargetParty = AiParty"
        SFX.ExecuteStatement "Set Target = AIBattle_" & pTarget.Index
    End If

End Function

Private Function SetScriptDef(pOfender As clsBattlePlayer)

Dim iAttackIndex As Integer

    'Get Next Avail
    iAttackIndex = RandomNumber(Char.UBound, 0)
    While pBattleHu(iAttackIndex).Alive = False
        iAttackIndex = RandomNumber(Char.UBound, 0)
    Wend

    ScriptDef.Index = pOfender.Index
    
    'Setup variables in script
    If pOfender.Enemy = False Then
        Script.ExecuteStatement "Set Offender = HumanBattle_" & pOfender.Index
        Script.ExecuteStatement "Set OffensiveParty = HumanParty"
    Else
        Script.ExecuteStatement "Set Offender = AIBattle_" & pOfender.Index
        Script.ExecuteStatement "Set OffensiveParty = AiParty"
    End If
    
    Script.ExecuteStatement "Set Self = AIBattle_" & pOfender.Index
    Script.ExecuteStatement "Set RandomPlayer = HumanBattle_" & iAttackIndex
    
End Function

Private Function DoQuedActions() As Boolean

Dim bTargetIsAI As Boolean, bOfenderIsAI As Boolean, pTarget As clsBattlePlayer, pOfender As clsBattlePlayer, _
sP() As String

    If cBattleActions.Count = 0 Then
        'Allow normal gameplay to continue
        
        SelectNextReadyChar
        bAnimationBusy = False
    
        Exit Function
    End If
    
    'Determine the qued actions
    sP = Split(cBattleActions(1), ":")
    cBattleActions.Remove 1
        
    If Mid(sP(2), 1, 1) = "H" Then
        bTargetIsAI = False
            
        Set pTarget = pBattleHu(sP(0))
    Else
        bTargetIsAI = True
            
        Set pTarget = pBattleAI(sP(0))
    End If
            
    If Mid(sP(2), 2, 1) = "A" Then
        bOfenderIsAI = True
            
        Set pOfender = pBattleAI(sP(1))
    Else
        bOfenderIsAI = False
            
        Set pOfender = pBattleHu(sP(1))
        iPartyIndex = CInt(sP(1))
    End If
        
    If pOfender.Alive = False Then
        'Ofender cant attack if they are dead
        DoQuedActions
        Exit Function
    End If

    'Execute the determined actions
    Select Case sP(3)
        
    Case Act_Attack
        'Attack
        AttackTarget pTarget, pOfender
        
    Case Act_Item
        'Give Item
        GiveTarget pTarget, pOfender
        
    Case Act_Steal
        'Steal Item
        StealTarget pTarget, pOfender
        
    Case Act_Custom
        'Cast Magic
        If pOfender.Enemy = True Then
            'MsgBox ":D"
        End If
        
        CastTarget pTarget, pOfender, CStr(sP(4)), CStr(sP(5)), CStr(sP(6))
        
    Case Else
        MsgBox "Ooops " & sP(3), vbCritical
        
    End Select
    
End Function

Private Function BeforeTurnFinished()

Dim I As Integer

    If LastOffender.Enemy = True Then
    'Offender was AI
        
        'Restore enemy normal stance
        imgEnemy(LastOffender.Index).ShowStance (0)
        
        If LastTarget.Enemy = False Then
            'Restore normal character stance
            Char(iSubjectIndex).ShowStance (0)
        End If
    Else
        'Offender was human (did they give an object maybe)
        Char(iOffenderIndex).ShowStance (0)
    End If
        
    If bKill = True Then
        'Something has died
        bKill = False
        
        BerryDead
    End If

    For I = 0 To MnuEnemies.ListCount
        'Hide the dead
        If pBattleAI(I).Alive = False Then
            'The dead dont have shadows
            DealWithDeadAI I
        End If
    Next

    If timShowAttack.Enabled = True Then
        MsgBox "An expected error occured, please qoute #1 Expected error", vbCritical
        Exit Function
    End If
    
    TurnFinished

End Function

Private Function DealWithDeadAI(Index As Integer)

    imgEnemy(Index).Visible = False

    shpShadow(Index).Visible = False
    
    ApplyDrops pBattleAI(Index)
    
    MnuEnemies.RemoveItem pBattleAI(Index).Index
            
    'imgEnemy(pBattleAI(Index).Index).Visible = False
    timATM(Index).Enabled = False
    iAiATM(Index) = 0

End Function

Private Function DealWithDeadHuman(Index As Integer)
    
    Debug.Print "DealWithDeadHuman"

    MnuCharacters.CharDied Index
    RemoveSpare Index
    
    If MnuCharacters.HighlightedIndex = Index And bTakingTurn = True Then
        If cCharsReady.Count > 0 Then
            'Select next ready char
            SwitchMenu MnuActions, Mnu_Action
            
            bTakingTurn = True
            
            MnuCharacters.Highlight cCharsReady(1)
            cCharsReady.Remove 1
        Else
            bTakingTurn = False
            CurrentMnu.HideCursor
        End If
    End If

    Char(Index).ShowStance (8)

End Function

Private Function PlayersToCommand(pTarget As clsBattlePlayer, pOfender As clsBattlePlayer)

Dim sTa As String, sOfe As String

    If pTarget.Enemy = True Then
        sTa = "A"
    Else
        sTa = "H"
    End If
        
    If pOfender.Enemy = True Then
        sOfe = "A"
    Else
        sOfe = "H"
    End If
    
    PlayersToCommand = sTa & sOfe

End Function

Private Function CastWhenReady(pTarget As clsBattlePlayer, pOfender As clsBattlePlayer, sPath As String, iMpCost As Integer, Optional sSpell As String)

    Dim sSpellPath As String

    CastWhenReady = True

    If sSpell = "" Then
        sSpell = MnuCustom.GetSelectedText
    Else
        'AI
        sSpellPath = pOfender.Path & "Actions\" & sSpell & ".def"
    End If

    If pOfender.Mp >= iMpCost Then
        'We can cast
        
        pOfender.Mp = pOfender.Mp - iMpCost
        
        If pOfender.Enemy = False Then
            'Show MP Deduction in Menu
            
            MnuCharacters.UpdateMP pOfender.Index, pOfender.Mp
        End If
    Else
        'Insuffi
        If pOfender.Enemy = False Then
            NotifyAction "Insufficient MP to cast!"
            
            SwitchMenu MnuActions, Mnu_Action
        End If
    
        CastWhenReady = False
        Exit Function
    End If

    If bAnimationBusy = True Then
        cBattleActions.Add pTarget.Index & ":" & pOfender.Index & ":" & PlayersToCommand(pTarget, pOfender) & ":" & _
                           Act_Custom & ":" & sPath & ":" & CStr(iMpCost) & ":" & sSpell
        
        Exit Function
    End If
    
    If pOfender.Enemy = False Then
        iPartyIndex = pOfender.Index
    End If
    
    NotifyUser "CastWhenReady: " & pOfender.Name & " <> " & pTarget.Name
    
    CastTarget pTarget, pOfender, sPath, iMpCost, sSpell, sSpellPath
    
End Function

Private Function StealWhenReady(pTarget As clsBattlePlayer, pOfender As clsBattlePlayer)

    If bAnimationBusy = True Then
        cBattleActions.Add pTarget.Index & ":" & pOfender.Index & ":" & PlayersToCommand(pTarget, pOfender) & ":" & _
                           Act_Steal
        
        Exit Function
    End If
    
    StealTarget pTarget, pOfender
    
End Function

Private Function GiveWhenReady(pTarget As clsBattlePlayer, pOfender As clsBattlePlayer)

    If bAnimationBusy = True Then
        cBattleActions.Add pTarget.Index & ":" & pOfender.Index & ":" & PlayersToCommand(pTarget, pOfender) & ":" & _
                           Act_Item
        
        Exit Function
    End If
    
    If pOfender.Enemy = False Then
        iPartyIndex = pOfender.Index
    End If
    
    GiveTarget pTarget, pOfender
    
End Function

Private Function AttackWhenReady(pTarget As clsBattlePlayer, pOfender As clsBattlePlayer)

    If bAnimationBusy = True Then
        cBattleActions.Add pTarget.Index & ":" & pOfender.Index & ":" & PlayersToCommand(pTarget, pOfender) & ":" & _
                           Act_Attack
        
        Exit Function
    End If
    
    If pOfender.Enemy = False Then
        iPartyIndex = pOfender.Index
    End If
    
    AttackTarget pTarget, pOfender

End Function

Public Function StartATMs()

Dim sMusic As String, I As Integer

    modBattle.PositionChars sBattlePath & "\battle-pos.ini"
    AlignShadows
    
    InitialiseItems

    If bBoss = True Then
        sMusic = "Boss"
    Else
        sMusic = "Fiend"
    End If
    
    PlayAudio sMusic, True
    
    MnuCharacters.StartATM
    
    For I = 0 To MnuEnemies.ListCount
        NotifyUser "Enabling Enemy ATB: " & I
        timATM(I).Enabled = True
    Next
    
    For I = 0 To MnuCharacters.ListUbound
    
        pBattleHu(I).TempStats = True
    Next
    
    'Engage shadow
    timFloat.Enabled = True
    
    'Hide Cursors
    imgCur(0).Visible = False
    imgCur(1).Visible = False
    
    imgSlash.ZOrder 0
    imgSlash.MaskColor = lTrans

    'Enable Keys
    keys.Enabled = True

End Function

Private Function AlignShadows()

Dim I As Integer

    For I = 0 To imgEnemy.UBound
        With shpShadow(I)
        
            .Width = imgEnemy(I).Width
            .Height = imgEnemy(I).Height / 5
        
            .Left = imgEnemy(I).Left
            .Top = imgEnemy(I).Top + imgEnemy(I).Height
            
            .ZOrder 0
        End With
    Next

End Function

Private Function GetPlayer(Mnu, Index As Integer) As clsBattlePlayer

    If Mnu.Tag = "Player" Then
        'Player
        GetPlayer = pBattleHu(Index)
    
    ElseIf Mnu.Tag = "Enemies" Then
        'Enemies
        GetPlayer = pBattleAI(Index)
        
    End If

End Function

Public Sub PrepareField()

Dim I As Integer

    Set cCharsReady = New Collection
    Set cMnuHistory = New Collection

    timShowAttack.Enabled = False
    timShowAttack.Tag = 0
    
    timShowDamage.Enabled = False
    timShowDamage.Tag = 0
    
    imenuState = Mnu_Action

    For I = 1 To imgEnemy.UBound
        Unload imgEnemy(I)
        Unload shpShadow(I)
    Next
    
    For I = 0 To MnuEnemies.ListCount
        timATM(I).Enabled = False
        timATM(I).Interval = 0
    Next
    
    imgEnemy(0).Tag = ""

    If IsSomething(CurrentMnu) = True Then
        CurrentMnu.HideCursor
    End If

    For I = 0 To MnuCharacters.ListUbound
    
        If pBattleHu(I).Alive = True Then
            MnuCharacters.ResetATM I
        End If
    Next
    
    MnuEnemies.ClearList
    MnuActions.ClearList

    iFloatState = 0
    
    bFirstEnemy = False
    bTakingTurn = False

    'Refresh Script Classes, etc
    FormatField

End Sub

Private Sub FormatField()
    
    MnuActions.AddItem "Attack", "A"
    MnuActions.AddItem "Items", "I"
    
    imgCur(1).Visible = False
    imgCur(0).Visible = False
    
    imgSlash.MaskColor = &H8000000F
    
    RefreshScriptObjects
    
End Sub

Private Sub RefreshScriptObjects()

Dim I As Integer, Support As New clsScriptSupport

    Script.Reset
    SFX.Reset True
    
    Script.Reset
    Script.AddObject "Actions", ScriptDef
    Script.AddObject "Debug", DebugWin

    'Add Partys
    For I = 0 To C_MAX_PLAYERS
        pParty(0).AttachPlayer I, pBattleHu(I)
        pParty(1).AttachPlayer I, pBattleAI(I)
    
        Script.AddObject "HumanBattle_" & I, pBattleHu(I)
        Script.AddObject "AiBattle_" & I, pBattleAI(I)
        SFX.AddObject "HumanBattle_" & I, pBattleHu(I)
        SFX.AddObject "AiBattle_" & I, pBattleAI(I)
    Next
    
    Script.AddObject "HumanParty", pParty(0)
    Script.AddObject "AiParty", pParty(1)
    Script.AddObject "Support", Support, True
    
    SFX.AddObject "HumanParty", pParty(0)
    SFX.AddObject "AiParty", pParty(1)

End Sub

Function RemovePlayer(I As Integer) As Boolean

Dim lLen As Integer

    RemovePlayer = True

    lLen = MnuCharacters.ListUbound

    If I = -1 Then
        RemovePlayer = False
        Exit Function
    End If
    
    MnuCharacters.RemoveCharacter I
    CharBank.MoveIn pBattleHu(I)
    
    'Swap with end (if not already)
    If I < lLen Then
        'Set pBattleHu(I) = pBattleHu(lLen)
        CopyBattleChar pBattleHu(I), pBattleHu(lLen)
        
        pSlashSize(I) = GetImageSize(Char(lLen).RetrieveImage(Sta_Attack))
        
        Char(I).LoadStore Char(lLen).Path, 8
        Char(I).LoadAdditional Char(lLen).Path & "attack.bmp"
        
        Char(I).ShowStance 0
    End If
    
    'Remove end
    Set pBattleHu(lLen) = New clsBattlePlayer
    
    If lLen = 0 Then
        'We cant unload stuff at runtime
        'all members are unloaded
        
        bFirstMember = False
    Else
        Unload Char(lLen)
    End If
    
End Function

Function CreatePlayer(ByVal sPath As String, Optional bAction As Boolean) As Boolean

Dim sIniPath As String, sName As String, sId As String, dIni As New clsIniObj

    On Error GoTo Catch_E
    
    CreatePlayer = True

    If bBattleEnd = False Then
        Debug.Print "Warning: Battle not yet ended."
    End If

    sId = sPath

    sPath = sPath_BattleChars & sPath
    sIniPath = sPath & "\data.ini"
    
    dIni.File = sIniPath
    
    sName = ReadINIValue(sIniPath, "Visual", "Name", "????")
    sName = Aliases.GetItem(sName, "????")

    If bFirstMember = False Then
        bFirstMember = True
        
        Char(0).Visible = True
        'Do stuff for 1st
    Else
        'Load new char
        Load Char(Char.Count)
    End If

    dIni.Section = "Stats"
    
    With pBattleHu(Char.UBound)
    
        If CharBank.MoveOut(sId, pBattleHu(Char.UBound)) = False Then
            .Image = Char(Char.UBound)
            .Alive = True
            .ID = sId
                
            .Strength = dIni.Read("Strength", , , 1)
            .Defence = dIni.Read("Defence", , , 1)
            .Magic = dIni.Read("Magic", , , 1)
            .Spirit = dIni.Read("Spirit", , , 1)
                
            .MaxHp = dIni.Read("MaxHP", , , 1)
            .MaxMp = dIni.Read("MaxMp", , , 1)
            .Hp = dIni.Read("HP", , , 1)
            .Mp = dIni.Read("MP", , , 1)
                
            .Experience = dIni.Read("Experience", "Level", , 1)
            .Level = dIni.Read("Level", , , 1)
        
            .Name = sName
                
            .Enemy = False
            .ATB = ReadINIValue(sIniPath, "Stats", "ATB", 255)
            
            'Add in predefined abilities
            .Steal = CBol(dIni.Read("Steal", "Abilities", , "false"))
            
            .Image = Char(Char.UBound)
        End If
        
        'Index
        .Index = Char.UBound
    
    End With

    With Char(Char.UBound)
        .LoadStore sPath & "\battle\", 9
        .LoadAdditional sPath & "\battle\attack.bmp"
        
        .ShowStance 0
        .MaskColor = lTrans
        
        .Visible = True
        
        'Get all slash sizes
        pSlashSize(.Index) = GetImageSize(.RetrieveImage(Sta_Attack))
    End With
    
    'MsgBox Char.UBound
    MnuCharacters.AddCharacter sName, Char.UBound, pBattleHu(Char.UBound).Hp, pBattleHu(Char.UBound).Mp, CInt(ReadINIValue(sIniPath, "Stats", "ATB", 255))

    InitialiseActions pBattleHu(Char.UBound)

    'Ignore changes weve made
    Set cPosBattleChanges = New Collection
    Set cNegBattleChanges = New Collection
    
    Exit Function
    
Catch_E:
    CreatePlayer = False
    MsgBox "CreatePlayer Failed: " & Err.Description, vbCritical

End Function

Sub Over()

Dim I As Integer

    bBattleEnd = False
    bAnimationBusy = False
    
    timWin.Enabled = False
    timFloat.Enabled = False
    timTurnFinished.Enabled = False

    For I = 0 To Char.UBound
        If pBattleHu(I).Alive = True Then
            Char(I).ShowStance 0
        End If
    Next

    'Kill all enemies
    For I = 0 To MnuEnemies.ListCount
        DealWithDeadAI I
    Next

    MnuEnemies.Visible = True
    MnuEnemies.ClearList
    
    'keys.ClearKeys
    keys.Enabled = False

End Sub

Sub ApplyExp()

    'Add up all AI Exp
Dim I As Integer, iExp As Long, indExp As Long, iLevel As Integer

    For I = 0 To 3
        iExp = iExp + pBattleAI(I).Experience
    Next

    indExp = iExp / MnuCharacters.AliveCount

    For I = 0 To MnuCharacters.ListUbound
    
        If pBattleHu(I).Alive = True Then
            pBattleHu(I).CleanStats

            pBattleHu(I).Experience = pBattleHu(I).Experience + indExp
            
            iLevel = pBattleHu(I).Level
            
            Script.ExecuteStatement "Set Character = HumanBattle_" & I
            RunScript Fload(sPath_BattleChars & pBattleHu(I).ID & "\exp.def")
            
            If iLevel < pBattleHu(I).Level Then
                NotifyAction pBattleHu(I).Name & " has grown a level !"
                InitialiseActions pBattleHu(I), True
            End If
            
            UpdateCharDisplay I
        End If
    Next

End Sub

Sub ApplyDrops(pTarget As clsBattlePlayer)

Dim iDropIndex As Integer, sP() As String, sEquip As String

    iDropIndex = RandomNumber(pTarget.Steals.Count, 1)

    If pTarget.Steals.Count = 0 Then
        Exit Sub
    Else
        'Attempt Steal
        sP = Split(pTarget.Steals(iDropIndex), ":")
            
        If StealTest(CInt(sP(3))) = True Then
            
            If LCase(sP(0)) = "item" Then
                NotifyAction "Recieved " & sP(1) & "!"
                Inventory.AddItem sP(1)
                    
                pTarget.Steals.Remove iDropIndex
                    
                InitialiseItems
            ElseIf LCase(sP(0)) = "equip" Then
                
                sEquip = sPath_Equipment & sP(1)
                frmMain.Equip(0).AddEquip sEquip
                
                NotifyAction "Recieved " & ReadINIValue(sEquip, "Visual", "Name", "????") & "!"
                    
                pTarget.Steals.Remove iDropIndex
            Else
                MsgBox "Drop type: " & sP(0) & " should be item or equip.", vbCritical
                End
            End If
        End If
    End If

End Sub

Sub TurnFinished()
    'Everything thats not sprite related

    'On Error GoTo why
    
Dim I As Integer, sMusic As String

    For I = 0 To MnuCharacters.ListUbound
    
        MnuCharacters.UpdateHP I, pBattleHu(I).Hp
        MnuCharacters.UpdateMP I, pBattleHu(I).Mp
    Next

    'QuickCheck
    If AllEnemiesDead = True And AllAllysDead = False Then
        'Repeat
        
        'Reset battle actions (stop other chars attacks)
        Set cBattleActions = New Collection
        
        StopAudio "boss"
        StopAudio "fiend"
        
        PlayAudio "victory", True
        
        ApplyExp
        
        For I = 0 To MnuEnemies.ListCount
            Set pBattleAI(I) = New clsBattlePlayer
            iAiATM(I) = 0
        Next
        
        'Hide all cursors
        imgCur(0).Visible = False
        imgCur(1).Visible = False
        
        bAnimationBusy = False
        bTakingTurn = False
        
        timWin.Enabled = True
        
    ElseIf AllAllysDead = True Then
    
        Over
        timDied.Enabled = True
        
        Exit Sub
        
    End If

    MnuCharacters.ResumeATM
    Me.ResumeAIATB

    If bAttackAi = False Then
        'debug.print "Reseting ATM: " & iPartyIndex
        MnuCharacters.ResetATM iPartyIndex
    Else
        timATM(iEnemyIndex).Enabled = True
    End If

    DoQuedActions
    
    Exit Sub
why:
    MsgBox Err.Description
        
End Sub

Sub SwitchMenu(Mnu, MnuLocation As Integer)

    'debug.print "Switching Menu"
Dim iScale As Integer, I As Integer, CurrentChar As clsBattlePlayer

    iScale = 15

    If IsObject(CurrentMnu) = True Then
        CurrentMnu.HideCursor
    End If
    
    'This is always the latest mnu
    Set CurrentMnu = Mnu
    Set CurrentChar = pBattleHu(MnuCharacters.HighlightedIndex)
    
    Mnu.Visible = True
    'Bring to front
    Mnu.ZOrder 0
    
    imenuState = MnuLocation
    
    '2 Different cursors
    Dim iCur As Integer
    
    If MnuLocation = Mnu_Enemies Then
        MnuCustom.Visible = False
        MnuItems.Visible = False
    
        iCur = 1
        iScale = 1
    End If
    
    'MnuItems is behind MnuEnemies..
    If MnuLocation = Mnu_Action Then
        'Hide Items / Show Enemies
        MnuEnemies.Visible = True
        MnuCustom.Visible = False
        MnuItems.Visible = False
        
        'Get Available Actions
        MnuActions.ClearList
        
        MnuActions.AddItem "Attack", Act_Attack
        
        If CurrentChar.Steal = True Then
            MnuActions.AddItem "Steal", Act_Steal
        End If
        
        If CurrentChar.Actions.Count > 0 Then
            For I = 1 To CurrentChar.Actions.Count
                MnuActions.AddItem CurrentChar.Actions(I), CStr(I)
            Next
            
            '(So we can find path later)
            MnuCustom.Tag = pBattleHu(MnuCharacters.HighlightedIndex).Actions(1)
        End If
        
        MnuActions.AddItem "Items", Act_Item
        
    ElseIf MnuLocation = Mnu_Items Then
        'Hide Items / Show Enemies
        MnuEnemies.Visible = False
        imgCur(1).Visible = True

        iScale = 1
        iCur = 1
        
    ElseIf MnuLocation = Mnu_Characters Then
        MnuCustom.Visible = False
        MnuItems.Visible = False
    End If
    
    Mnu.AttachCur imgCur(iCur), Mnu.Left, Mnu.Top, iScale
End Sub

Private Sub Keys_OnKeyPressed(ByVal KeyAscii As Integer)

    If IsSomething(CurrentMnu) = False Then
        Debug.Print "Error: CurrentMnu is Null"
        Exit Sub
    End If
    
    If imgCur(0).Visible = False And imgCur(1).Visible = False Then
        Debug.Print "Error: No Cursor to move"
        Exit Sub
    End If

    Select Case GetDirection(KeyAscii)
    
    Case 4
        'Up Key
        CurrentMnu.GoUp
    
    Case 3
        'Down Key
        CurrentMnu.GoDown
    Case 2
        'Switch menus ?
        If imenuState = Mnu_Enemies Then
            CurrentMnu.HideCursor
            MnuActions.Visible = False

            SwitchMenu MnuCharacters, Mnu_Characters
        End If
        
    Case 1
        'Switch menus ?
        If imenuState = Mnu_Characters Then
            MnuCharacters.HideCursor
            MnuActions.Visible = True

            SwitchMenu MnuEnemies, Mnu_Enemies
        End If
        
    Case Else
        If KeyAscii = Control_Select Then
            Call Selected
        ElseIf KeyAscii = Control_Cancel Then
            Call Cancel
        End If
    
    End Select

End Sub

Private Function CastTarget(ByVal pTarget As clsBattlePlayer, ByVal pOfender As clsBattlePlayer, sPath As String, iMpCost As Integer, sSpell As String, Optional sSpellPath As String) As Boolean
                            '[!] ByVal, Pass COPY of pTarget!

Dim attackImage, I As Integer

    NotifyUser "CastTarget"

    On Error GoTo why

    If sSpellPath = "" Then
        sSpellPath = sPath_BattleChars & sPath
    End If

    'For SFX
    sLastSpell = sSpell
    
    NotifyAction sSpell

    'Change if dead etc
    ValidateTarget pTarget

    bAttackAi = pOfender.Enemy
    bTargetAi = pTarget.Enemy
    
    iSubjectIndex = pTarget.Index
    iOffenderIndex = pOfender.Index

   If pTarget.ReflectStatus = True Then
        If pOfender.Enemy = False And pTarget.Enemy = False Then
            'IF Caster was Human, and
            'target has reflectstatus AND
            'the target was human
            
            'change target to AI
            bTargetAi = True
            
            iSubjectIndex = FindRandomAI
            Set pTarget = pBattleAI(iSubjectIndex)
            
        ElseIf pOfender.Enemy = True And pTarget.Enemy = False Then
            'If Caster was Enemy, and
            'Target has reflectstatus AND
            'target was human
        
            'change target to AI
            bTargetAi = True
            
            'Find Random AI (thats alive)
            For I = 0 To UBound(pBattleAI)
                If pBattleAI(I).Alive = True Then
                
                    iSubjectIndex = I
                    Set pTarget = pBattleAI(I)

                    Exit For
                End If
            Next
            
        ElseIf pOfender.Enemy = False And pTarget.Enemy = True Then
            'If Caster was Human, and
            'Target has reflectstatus AND
            'Target was AI
            
            'change target to Human
            bTargetAi = False
            
            'Change subject to Caster
            iSubjectIndex = pOfender.Index
            Set pTarget = pBattleHu(pOfender.Index)
            
        ElseIf pOfender.Enemy = True And pTarget.Enemy = True Then
            'Ai casting a spell on an AI that ahs reflect
            
            'change target to Human
            bTargetAi = False

            'Change subject random alive human
            iSubjectIndex = FindRandomHuman
            Set pTarget = pBattleHu(iSubjectIndex)
            
        Else
            MsgBox "Unexpected Condition!", vbCritical
        End If
    End If

    bAnimationBusy = True
    bMagic = True

    'Pause ATB
    MnuCharacters.PauseATM
    Me.PauseAIATB

    SetScript pTarget, pOfender, True
    RunScript Fload(sSpellPath)

    ShowResults False

    If pOfender.Enemy = False Then
        'Do the walking thing
        timShowAttack.Tag = 0
        timShowAttack.Enabled = True
    Else
        imgEnemy(pOfender.Index).ShowStance (1)
    
        ShowDelayedDamage
        timTurnFinished.Enabled = True
    End If

    If pTarget.Enemy = True Then
        Set attackImage = imgEnemy(pTarget.Index)
    Else
        Set attackImage = Char(pTarget.Index)
    End If

    Exit Function
why:
    MsgBox Err.Description

End Function

Private Function FindRandomAI() As Integer

Dim I As Integer

    FindRandomAI = -1

    'Find Random AI (thats alive)
    For I = 0 To UBound(pBattleAI)
        If pBattleAI(I).Alive = True Then
            FindRandomAI = I
            
            Exit For
        End If
    Next

End Function

Private Function FindRandomHuman() As Integer

Dim I As Integer

    FindRandomHuman = -1

    'Find Random AI (thats alive)
    For I = 0 To UBound(pBattleHu)
        If pBattleHu(I).Alive = True Then
            FindRandomHuman = I
            
            Exit For
        End If
    Next

End Function

Private Function ValidateTarget(ByRef pTarget As clsBattlePlayer)

Dim attackIndex As Integer

    If pTarget.Alive = False Then
        If pTarget.Enemy = False Then
            attackIndex = RandomNumber(Char.UBound, 0)
            While pBattleHu(attackIndex).Alive = False
                attackIndex = RandomNumber(Char.UBound, 0)
            Wend
            
            Set pTarget = pBattleHu(attackIndex)
        Else
            attackIndex = RandomNumber(imgEnemy.UBound, 0)
            While pBattleAI(attackIndex).Alive = False
                attackIndex = RandomNumber(imgEnemy.UBound, 0)
            Wend
            
            Set pTarget = pBattleAI(attackIndex)
        End If
        
    End If

End Function

Private Function StealTarget(ByRef pTarget As clsBattlePlayer, ByRef pOfender As clsBattlePlayer) As Boolean

Dim attackImage, sType As String, iStealIndex As Integer, sP() As String, sEquip As String

    On Error GoTo why
    
    StealTarget = True
    bAnimationBusy = True
    
    bAttackAi = pOfender.Enemy
    bTargetAi = pTarget.Enemy
    
    iSubjectIndex = pTarget.Index
    iOffenderIndex = pOfender.Index

    SetScript pTarget, pOfender
    
    'Pause ATB
    MnuCharacters.PauseATM
    Me.PauseAIATB
    
    If pTarget.Steals.Count = 0 Then
        NotifyAction "Nothing left to steal!"
    Else
        'Attempt Steal
        iStealIndex = RandomNumber(pTarget.Steals.Count, 1)
        sP = Split(pTarget.Steals(iStealIndex), ":")
        
        If StealTest(CInt(sP(2))) = True Then
        
            If LCase(sP(0)) = "item" Then
                NotifyAction "Stole " & sP(1) & "!"
                Inventory.AddItem sP(1)
                
                pTarget.Steals.Remove iStealIndex
                
                InitialiseItems
            ElseIf LCase(sP(0)) = "equip" Then
                
                sEquip = sPath_Equipment & sP(1)
                frmMain.Equip(0).AddEquip sEquip
                
                NotifyAction "Stole " & ReadINIValue(sEquip, "Visual", "Name", "????") & "!"
                
                pTarget.Steals.Remove iStealIndex
            Else
                MsgBox "Steal type: " & sP(0) & " should be item or equip.", vbCritical
                End
            End If
        Else
            NotifyAction "Steal failed!"
        End If
    End If
    
    'Set to giving stance
    If pOfender.Enemy = False Then
        Char(pOfender.Index).ShowStance (9)
    End If
    
    ShowResults True
    timTurnFinished.Enabled = True

    Exit Function
why:
    MsgBox Err.Description

End Function

Private Function GiveTarget(ByRef pTarget As clsBattlePlayer, ByRef pOfender As clsBattlePlayer) As Boolean

Dim attackImage, sType As String

    On Error GoTo why
    
    GiveTarget = True
    bAnimationBusy = True
    
    bAttackAi = pOfender.Enemy
    bTargetAi = pTarget.Enemy
    
    iSubjectIndex = pTarget.Index
    iOffenderIndex = pOfender.Index

    SetScript pTarget, pOfender
    
    'Pause ATB
    MnuCharacters.PauseATM
    Me.PauseAIATB
    
    'Subtract item
    sType = MnuItems.GetSelectedTag
    
    Inventory.TakeItem sType
    
    If Inventory.TypeCount(sType) < 1 Then
        MnuItems.RemoveItem MnuItems.ListIndex
    Else
        MnuItems.ChangeCaption MnuItems.ListIndex, sType & " x " & Inventory.TypeCount(sType)
    End If
    
    'Load what item does
    RunScript Fload(sPath_Resources & "\items\" & sType & ".def")

    If pTarget.Enemy = True Then
        Set attackImage = imgEnemy(pTarget.Index)
    Else
        Set attackImage = Char(pTarget.Index)
    End If
    
    'Set to giving stance
    If pOfender.Enemy = False Then
        Char(pOfender.Index).ShowStance (6)
    End If
    
    ShowResults True
    timTurnFinished.Enabled = True

    Exit Function
why:
    MsgBox Err.Description

End Function

Sub ActivateDamage(ByVal charIndex As Integer, Img, lAmount As Integer, Optional bStart As Boolean = True, Optional bGood As Boolean, Optional bEnemy As Boolean = False)

Dim iFreeAni As Integer

    If bEnemy = True Then
        'Same ANI's used, unless Enemy ANI's use +4
        charIndex = charIndex + 4
    End If

    ShowNumbers charIndex, Img, CLng(lAmount), 20, bGood
    iFreeAni = charIndex * 5

    If bStart = True Then
        imgNumbers(iFreeAni).Visible = True
        timAnim(iFreeAni).Enabled = True
    Else
        cNeedsActivation.Add iFreeAni
    End If
    
End Sub

Sub ActivateMP(ByVal charIndex As Integer, Img, lAmount As Integer, Optional bStart As Boolean = True, Optional bGood As Boolean)

    'Dont interfer with HP (Dont raise TurnFinished event
    charIndex = 4 + charIndex

    ShowNumbers charIndex, Img, CLng(lAmount), 20, bGood, True
    charIndex = (charIndex * 5)

    If bStart = True Then
        imgNumbers(charIndex).Visible = True
        timAnim(charIndex).Enabled = True
    Else
        cNeedsActivation.Add charIndex
    End If
    
End Sub

Private Function AttackTarget(ByVal pTarget As clsBattlePlayer, ByVal pOfender As clsBattlePlayer)
                              '[!] ByVal, Pass COPY of pTarget!
    NotifyUser "AttackTarget"
    
Dim lHp As Integer, allyIndex As Integer, attackImage, attackIndex As Integer
    
    'if target is dead then find a random on, on the same team, but first check there not all dead
    If AllAllysDead = True Or AllEnemiesDead = True Then
        AttackTarget = False
        Exit Function
    End If
    
    'Change if dead etc
    ValidateTarget pTarget
    AttackTarget = True
    
    bAnimationBusy = True
    bMagic = False
    
    bAttackAi = pOfender.Enemy
    bTargetAi = pTarget.Enemy

    iSubjectIndex = pTarget.Index
    iOffenderIndex = pOfender.Index
    
    'Pause ATB
    MnuCharacters.PauseATM
    Me.PauseAIATB
    
    If pTarget.Enemy = True Then
        Set attackImage = imgEnemy(pTarget.Index)

    Else
        Set attackImage = Char(pTarget.Index)

    End If
    
    If pTarget.Alive = False Then
        lHp = 0
    End If
 
    SetScript pTarget, pOfender
    RunScript Fload(sPath_Battle & "attack.def")
    
    'Show Results, but not yet
    ShowResults False

    If pOfender.Enemy = False Then
        ''debug.print "AttackTarget False"
        
        'Each char has different slash type
        Set imgSlash.Picture = Char(iOffenderIndex).RetrieveImage(Sta_Attack)

        With imgSlash
            .Height = pSlashSize(iOffenderIndex).Height / 15
            .Width = pSlashSize(iOffenderIndex).Width / 15
        End With
        
        imgSlash.Top = attackImage.Top + (attackImage.Height / 2) - (imgSlash.Height / 2)
        imgSlash.Left = attackImage.Left + (attackImage.Width / 2) - (imgSlash.Width / 2)
        
        timShowAttack.Tag = 0
        timShowAttack.Enabled = True
    Else
        'debug.print "AttackTarget True"
        
        imgEnemy(pOfender.Index).ShowStance (1)

        With Char(pTarget.Index)
            .ShowStance 7
        End With
        
        iEnemyIndex = pOfender.Index
        
        'Number Allocation
        timAnim(pTarget.Index * 5).Enabled = True
        timTurnFinished.Enabled = True
    End If

    Exit Function
    
why:
    MsgBox Err.Description

End Function

Function ShowResults(Optional bStart As Boolean = True)

    bAttackFin = False

Dim Item, sP() As String, charIndex As Integer, Img, I As Integer

bKill = False
  
    Set cNeedsActivation = New Collection
  
    While cPosBattleChanges.Count > 0
        sP = Split(CStr(cPosBattleChanges(1)), ":")
        charIndex = sP(1)
        
        If sP(3) = True Then
            Set Img = imgEnemy(charIndex)
        Else
            Set Img = Char(charIndex)
        End If
        
        Select Case sP(0)
            
        Case "atb"
            'No Activation of anything
            If sP(3) = False Then
                NotifyAction pBattleHu(charIndex).Name & "'s ATB has decreased by " & CLng(sP(2))
                MnuCharacters.UpdateATB charIndex, pBattleHu(charIndex).ATB
            Else
                NotifyAction pBattleAI(charIndex).Name & "'s ATB has decreased by " & CLng(sP(2))
                timATM(charIndex).Interval = pBattleAI(charIndex).ATB
            End If
        
        Case "mp"
            ActivateMP charIndex, Img, CLng(sP(2)), bStart, True
        
        Case "hp"
            'Because we need to show damage (and prevent minus) from Negs
            If sP(3) = False And pBattleHu(charIndex).Alive = True Then
                ActivateDamage charIndex, Img, CLng(sP(2)), bStart, True, False
            ElseIf sP(3) = True And pBattleAI(charIndex).Alive = True Then
                ActivateDamage charIndex, Img, CLng(sP(2)), bStart, True, True
            End If
        
        Case "life"
            'Enemies cannot be revived
            MnuCharacters.CharRevived charIndex
            
        Case Else
            'No Activation of anything
            If sP(3) = False Then
                NotifyAction pBattleHu(charIndex).Name & "'s " & StatLang(sP(0)) & " has increased by " & CLng(sP(2))
            Else
                NotifyAction pBattleAI(charIndex).Name & "'s " & StatLang(sP(0)) & " has increased by " & CLng(sP(2))
            End If
        
        End Select
        
        cPosBattleChanges.Remove 1
    Wend

    While cNegBattleChanges.Count > 0
        sP = Split(CStr(cNegBattleChanges(1)), ":")
        charIndex = sP(1)
        
        If sP(3) = True Then
            Set Img = imgEnemy(charIndex)
        Else
            Set Img = Char(charIndex)
        End If
        
        Select Case sP(0)
        
        Case "atb"
            'No Activation of anything
            If sP(3) = False Then
                NotifyAction pBattleHu(charIndex).Name & "'s ATB has increased by " & CLng(sP(2))
                MnuCharacters.UpdateATB charIndex, pBattleHu(charIndex).ATB
            Else
                NotifyAction pBattleAI(charIndex).Name & "'s ATB has increased by " & CLng(sP(2))
                timATM(charIndex).Interval = pBattleAI(charIndex).ATB
            End If
        
        Case "mp"
            'Never show MP going down on battlefield
            ActivateMP charIndex, Img, 1111, bStart, False
            
            imgNumbers(20 + (charIndex * 5)).Top = -8
            imgNumbers(20 + (charIndex * 5) + 1).Top = -8
            imgNumbers(20 + (charIndex * 5) + 2).Top = -8
            imgNumbers(20 + (charIndex * 5) + 3).Top = -8
        
        Case "hp"
            If sP(3) = True Then
                ActivateDamage charIndex, Img, CLng(sP(2)), bStart, False, True
            Else
                ActivateDamage charIndex, Img, CLng(sP(2)), bStart, False, False
            End If
            
        Case "life"
            'Kill it
            'Set imgKill(charIndex) = Img
            bKill = True

            DigGrave Img, CBol(sP(3)), charIndex

            If sP(3) = True Then

                'Disable auto bots
                timATM(charIndex).Enabled = False
                timATM(charIndex).Interval = 0
            End If
        
        Case Else
            'No Activation of anything
            If sP(3) = False Then
                NotifyAction pBattleHu(charIndex).Name & "'s " & StatLang(sP(0)) & " has decreased by " & CLng(sP(2))
            Else
                NotifyAction pBattleAI(charIndex).Name & "'s " & StatLang(sP(0)) & " has decreased by " & CLng(sP(2))
            End If
        
        End Select
        
        cNegBattleChanges.Remove 1
    Wend

End Function

Function AllEnemiesDead() As Boolean

    Dim I As Integer
    
    For I = 0 To imgEnemy.UBound
        If pBattleAI(I).Alive = True Then
            AllEnemiesDead = False
            Exit Function
        End If
    Next
    
    bBattleEnd = True
    AllEnemiesDead = True

End Function

Function AllAllysDead() As Boolean

    Dim I As Integer
    
    For I = 0 To Char.UBound
        If pBattleHu(I).Alive = True Then
            AllAllysDead = False
            Exit Function
        End If
    Next
    
    bBattleEnd = True
    AllAllysDead = True

End Function

Sub CreateEnemy(sPath As String, Optional bBoss As Boolean = False)

    'An enemy is a player
    'path must have .ply
    On Error Resume Next

    Dim Img, sName As String, sIniPath As String, pSize As picSize, sId As String
    sId = sPath

    If bBoss = False Then
        sPath = sPath_Fiends & sPath
    Else
        sPath = sPath_Bosses & sPath
    End If
    
    sIniPath = sPath & "\data.ini"
    sName = ReadINIValue(sIniPath, "Visual", "Name", "????")

    If bFirstEnemy = False Then
        bFirstEnemy = True
        
        Set Img = imgEnemy(0)
        Img.Visible = True
        'Do stuff for 1st
        With imgEnemy(imgEnemy.UBound)
        
            .LoadStore sPath, 1
            .ShowStance 0
            
            pSize = GetImageSize(imgEnemy(imgEnemy.UBound).Picture)
        
            .Height = pSize.Height / 15
            .Width = pSize.Width / 15
            
            '.Left = 37.5 - (.Width / 2)
            '.Top = 46.5 - (.Height / 2)
        End With
    Else
        Load shpShadow(shpShadow.Count)
        'Load after (to default zorder ontop)
        Load imgEnemy(imgEnemy.Count)
        
        Set Img = imgEnemy(imgEnemy.UBound)
        
        pSize = GetImageSize(imgEnemy(imgEnemy.UBound).Picture)
        With imgEnemy(imgEnemy.UBound)
            .Height = pSize.Height / 15
            .Width = pSize.Width / 15

            .LoadStore sPath, 1
            .ShowStance 0
            'If imgEnemy.UBound = 1 Then
                '.Left = 112.5 - (.Width / 2)
                '.Top = 46.5 - (.Height / 2)
                
            'ElseIf imgEnemy.UBound = 2 Then
                '.Top = 139.5 - (.Height / 2)
                '.Left = 37.5 - (.Width / 2)
                
            'ElseIf imgEnemy.UBound = 3 Then
                '.Left = 112.5 - (.Width / 2)
                '.Top = 139.5 - (.Height / 2)
                
            'End If
            
            .Visible = True
        End With
    End If
    
    If CBol(ReadINIValue(sIniPath, "Visual", "Floats", False)) = True Then
        shpShadow(imgEnemy.UBound).Visible = True
    End If
    
    With pBattleAI(imgEnemy.UBound)
        .Image = imgEnemy(imgEnemy.UBound)
        .Boss = bBoss
    
        .Alive = True
    
        'Set .iImg = imgEnemy(imgEnemy.UBound)
        .ID = sId
        
        .MaxHp = ReadINIValue(sIniPath, "Stats", "HP", 0)
        .Hp = .MaxHp
        
        .Strength = ReadINIValue(sIniPath, "Stats", "Strength", 0)
        .Defence = ReadINIValue(sIniPath, "Stats", "Defence", 0)
        
        .Experience = ReadINIValue(sIniPath, "Stats", "Experience", 0)
        .ATB = ReadINIValue(sIniPath, "Stats", "ATB", 3000)
        .Element = ReadINIValue(sIniPath, "Stats", "Element", 0)
        .Level = ReadINIValue(sIniPath, "Stats", "Level", 0)
        .Magic = ReadINIValue(sIniPath, "Stats", "Magic", 0)
        .Spirit = ReadINIValue(sIniPath, "Stats", "Spirit", 0)

        .Name = sName
        
        .Enemy = True

        .Index = imgEnemy.UBound

        .MaxMp = ReadINIValue(sIniPath, "Stats", "MP", 3)
        .Mp = .MaxMp
        
        .Path = sPath & "\"
        .HasDef = FileExist(sPath & "\actions.def")
    End With
    
    With timATM(imgEnemy.UBound)
        .Interval = ReadINIValue(sIniPath, "Stats", "ATB", 3000)
        .Enabled = True
    End With
    
    InitialiseSteals pBattleAI(imgEnemy.UBound)
    
    imgEnemy(imgEnemy.UBound).MaskColor = lTrans
    MnuEnemies.AddItem sName, ""

    If Err Then
        MsgBox "ReleaseFiend/CreateEnemy failed: " & Err.Description, vbCritical
    End If
    
    'Ignore changes
    Set cPosBattleChanges = New Collection
    Set cNegBattleChanges = New Collection

End Sub

Private Sub Cancel()
    'debug.print "Selected()"
    
    'Reasons to ignore user input
    If imgCur(1).Visible = False And imgCur(0).Visible = False Then
        'debug.print "Exiting Selected()"
        Exit Sub
    End If

    If cMnuHistory.Count = 0 Then
        Exit Sub
    End If

    Select Case cMnuHistory(cMnuHistory.Count)
    
    Case Mnu_Enemies
        SwitchMenu MnuEnemies, Mnu_Enemies
        
    Case Mnu_Items
        SwitchMenu MnuItems, Mnu_Items
        
    Case Mnu_Action
        SwitchMenu MnuActions, Mnu_Action
        
    Case Mnu_Custom
        SwitchMenu MnuCustom, Mnu_Custom
        
    Case Else
        MsgBox "External History Menu Error (Exception): " & cMnuHistory(1), vbCritical
    
    End Select
    
    cMnuHistory.Remove cMnuHistory.Count

End Sub

Private Sub Selected()

    Debug.Print "Selected()"

    If bBattleEnd = True Then
        If cNotifications.Count <> 0 Or _
            timNotify.Enabled = True Then
            
            Exit Sub
        End If
 
        StopAudio "victory"
 
        Over
        RaiseEvent BattleFinished(0)
        
        Exit Sub
    End If

    'Reasons to ignore user input
    If (imgCur(1).Visible = False And imgCur(0).Visible = False) Then
        
        Debug.Print "Exiting Selected()"
        Exit Sub
    End If

    cMnuHistory.Add imenuState

    If imenuState = Mnu_Action Then
        
        Handle_ActionOperation CurrentMnu.GetSelectedTag
    ElseIf imenuState = Mnu_Enemies Then
        
        Handle_EnemyOperation CurrentMnu.GetSelectedTag
    ElseIf imenuState = Mnu_Characters Then
        
        Handle_CharacterOperation
    ElseIf imenuState = Mnu_Items Then
    
        If Handle_ItemOperation(CurrentMnu.GetSelectedText) = False Then
            'Couldnt handle it! Remove history of last
            cMnuHistory.Remove cMnuHistory.Count
        End If
        
    ElseIf imenuState = Mnu_Custom Then
        
        If Handle_CustomOperation(CurrentMnu.GetSelectedText) = False Then
            'Couldnt handle it! Remove history of last
            cMnuHistory.Remove cMnuHistory.Count
        End If
    Else
        MsgBox "Unkown Mnu Allocation: " & imenuState, vbInformation
        Exit Sub
    End If

End Sub

Private Function Handle_CharacterOperation()

    Select Case sMnuAnswers(Mnu_Action)
    
    Case Act_Attack
        AttackWhenReady pBattleHu(MnuCharacters.ListIndex), pBattleHu(MnuCharacters.HighlightedIndex)
    
    Case Act_Item
        GiveWhenReady pBattleHu(MnuCharacters.ListIndex), pBattleHu(MnuCharacters.HighlightedIndex)
    
    Case Act_Steal
        Exit Function
    
    Case Else
        'Custom / Other
        If CastWhenReady _
        (pBattleHu(MnuCharacters.ListIndex), pBattleHu(MnuCharacters.HighlightedIndex), _
            pBattleHu(MnuCharacters.HighlightedIndex).ID & "\Actions\" & MnuActions.GetSelectedText & "\" & MnuCustom.GetSelectedText & ".def", _
            MnuCustom.GetSelectedNumber) = False Then
            
            'Casting was unsuccesfull, player is STILL taking turn
            Exit Function
        End If
    
    End Select
    
    'Character menu is last menu, clear history
    Set cMnuHistory = New Collection
    
    bTakingTurn = False
    If SelectNextReadyChar = False Then
        'No chars available, Wait
        MnuCharacters.HideCursor
    End If
    
End Function

Private Function Handle_ItemOperation(s As String) As Boolean

    Handle_ItemOperation = True

    If Len(s) = 0 Then
        'Nothing selected
        
        Handle_ItemOperation = False
        Exit Function
    End If
    
    sMnuAnswers(Mnu_Items) = s

    'Always goto characters menu
    MnuEnemies.Visible = True
    
    SwitchMenu MnuCharacters, Mnu_Characters
    bTakingTurn = True
End Function

Private Function Handle_EnemyOperation(s As String)
    'Enemy menu is last menu, clear history
    Set cMnuHistory = New Collection

    sMnuAnswers(Mnu_Enemies) = s
    
    Select Case sMnuAnswers(Mnu_Action)
    
    Case Act_Attack
        AttackWhenReady pBattleAI(MnuEnemies.ListIndex), pBattleHu(MnuCharacters.HighlightedIndex)

    Case Act_Item
        GiveWhenReady pBattleAI(MnuEnemies.ListIndex), pBattleHu(MnuCharacters.HighlightedIndex)

    Case Act_Steal
        StealWhenReady pBattleAI(MnuEnemies.ListIndex), pBattleHu(MnuCharacters.HighlightedIndex)

    Case Else
        'Custom / Other
        If CastWhenReady _
        (pBattleAI(MnuEnemies.ListIndex), pBattleHu(MnuCharacters.HighlightedIndex), _
            pBattleHu(MnuCharacters.HighlightedIndex).ID & "\Actions\" & MnuActions.GetSelectedText & "\" & MnuCustom.GetSelectedText & ".def", _
            MnuCustom.GetSelectedNumber) = False Then
            
            'Casting was unsuccesfull, player is STILL taking turn
            Exit Function
        End If

    End Select

    bTakingTurn = False
    If SelectNextReadyChar = False Then
        'No chars available, Wait
        MnuEnemies.HideCursor

    End If

End Function

Private Function Handle_CustomOperation(s As String) As Boolean
    Handle_CustomOperation = True

    If Len(s) = 0 Then
        Handle_CustomOperation = False
        Exit Function
    End If
    
    sMnuAnswers(Mnu_Custom) = s

    'Always goto characters menu
    MnuEnemies.Visible = True
    MnuActions.Visible = False
    
    SwitchMenu MnuEnemies, Mnu_Enemies
    bTakingTurn = True
    
End Function

Private Function Handle_ActionOperation(s As String)
    sMnuAnswers(Mnu_Action) = s
    
    Select Case s
    
    Case Act_Attack, Act_Steal
        MnuActions.Visible = False
        
        SwitchMenu MnuEnemies, Mnu_Enemies
        
    Case Act_Item
        MnuEnemies.Visible = False
        MnuActions.Visible = False
        
        SwitchMenu MnuItems, Mnu_Items
        
    Case Else
        'Get List of custom actions
        ListActions pBattleHu(MnuCharacters.HighlightedIndex).ActionNode(CInt(s))
        
        SwitchMenu MnuCustom, Mnu_Custom
    
    End Select

End Function

Private Function ListActions(cActions As Collection)

    MnuCustom.ClearList
    
Dim I As Integer, sP() As String
    For I = 1 To cActions.Count
        sP = Split(cActions(I), ":")
        MnuCustom.AddItem sP(0), "", sP(1)
    Next

End Function

Private Function SelectNextReadyChar() As Boolean

    'debug.print "SelectNextReadyChar() "
    SelectNextReadyChar = True

    If bTakingTurn = False And cCharsReady.Count > 0 Then
        'Were alloud to goto the menu
        bTakingTurn = True
        
        'MsgBox "#3"
         MnuCharacters.Highlight cCharsReady(1)
         
         SwitchMenu MnuActions, Mnu_Action
         
         cCharsReady.Remove 1
    Else
        'debug.print "SelectNextReadyChar(No Chars Aval) " & bTakingTurn
        SelectNextReadyChar = False
    End If

End Function

Sub LoadBattFile(sPath As String)
    'PlayNewMusic sPath_Resources & "battle\music.mp3"
    
    sBattlePath = sPath
    
    On Error GoTo Catch_E
    imgBack.Picture = LoadPicture(sBattlePath & "\battle.bmp")
    
    Exit Sub
    
Catch_E:
    If Err.Number = 53 Then
        WarnUser "Battle: Battle.bmp not found (" & sPath & ")", False
    End If
    
End Sub

Private Sub MnuCharacters_OnReady(iIndex As Integer)

    If bTakingTurn = False Then

        If cCharsReady.Count = 0 Then
        
            'MsgBox "#2 " & iIndex
            MnuCharacters.Highlight iIndex
        Else
            'MsgBox iIndex
            cCharsReady.Add iIndex
            
            'MsgBox "#1 " & iIndex
            MnuCharacters.Highlight cCharsReady(1)
            
            cCharsReady.Remove 1
        End If
        
        SwitchMenu MnuActions, Mnu_Action
        
        bTakingTurn = True
    Else
    
        'MsgBox "##" & iIndex
        cCharsReady.Add iIndex
    
    End If

End Sub

Sub RemoveSpare(iIndex As Integer)

    Dim I As Integer
    For I = 1 To cCharsReady.Count
        If cCharsReady(I) = iIndex Then
            cCharsReady.Remove I
            Exit Sub
        End If
    Next

End Sub

Private Sub Script_Error()

    MsgBox "Scripting Error: " & Script.Error.Description & vbCrLf & _
           "Document: " & KillHome(sFloadLast) & vbCrLf & vbCrLf & _
           "Line: " & Script.Error.Text, vbCritical
           
    End
           
End Sub

Private Sub ScriptDef_onAttack(Target As Variant)

Dim pTarget As clsBattlePlayer

    If IsSomething(Target) = False Then
        WarnUser "ScriptDef_onAttack:: Target = Null"
        
        Debug.Print "onAttack !!!"
        BeforeTurnFinished
        Exit Sub
    End If
    
    If Target.Enemy = False Then
        Set pTarget = pBattleHu(CInt(Target.Index))
    Else
        Set pTarget = pBattleAI(CInt(Target.Index))
    End If
    
    AttackWhenReady pTarget, pBattleAI(ScriptDef.Index)

End Sub

Private Sub ScriptDef_onCast(sSpell As String, Target As Variant)

Dim pTarget As clsBattlePlayer, pOfender As clsBattlePlayer, iMpCost As Integer, iAttackIndex As Integer

    If IsSomething(Target) = False Then
        WarnUser "ScriptDef_onAttack:: Target = Null"

        BeforeTurnFinished
        Exit Sub
    End If

    Set pOfender = pBattleAI(ScriptDef.Index)

    If Target.Enemy = False Then
        Set pTarget = pBattleHu(CInt(Target.Index))
    Else
        Set pTarget = pBattleAI(CInt(Target.Index))
    End If

    'Fetch Cost of MP
    iMpCost = ReadINIValue(pOfender.Path & "\Actions\Actions.ini", "MP Cost", sSpell, 0)
    
    If iMpCost = 0 Then
        DebugWin.AddItem "WARNING: MP Cost = 0"
    End If

    'What if we cant cast
    If CastWhenReady _
        (pTarget, pOfender, pOfender.ID & "\Actions\" & sSpell & ".def", iMpCost, sSpell) = True Then

        Exit Sub
    End If
    
    DebugWin.AddItem "AI #" & pOfender.Name & ":: Not enough MP! .. Attacking"

    iAttackIndex = RandomNumber(Char.UBound, 0)
    While pBattleHu(iAttackIndex).Alive = False
        iAttackIndex = RandomNumber(Char.UBound, 0)
    Wend
    
    ScriptDef_onAttack pBattleHu(iAttackIndex)

End Sub

Private Sub SFX_OnFinished()
    SFX.Visible = False
    sSFXFinished = True
    'SFX.Visible = True
    
    If cNotifications.Count = 0 Then
        timShowAttack.Enabled = True
    Else
        bWaitNotify = True
    End If
End Sub

Private Sub timATM_Timer(Index As Integer)

    Dim iAttackIndex As Integer
    iAiATM(Index) = iAiATM(Index) + 2
    
    If iAiATM(Index) > 32 Then
        'AI ATM Guage is ready
        
        If pBattleAI(Index).HasDef = False Then
            'Attack randomly
            iAttackIndex = RandomNumber(Char.UBound, 0)
            While pBattleHu(iAttackIndex).Alive = False
                iAttackIndex = RandomNumber(Char.UBound, 0)
            Wend
            
            AttackWhenReady pBattleHu(iAttackIndex), pBattleAI(Index)
        Else
            'Get Actions from Script
            SetScriptDef pBattleAI(Index)

            RunScript Fload(pBattleAI(Index).Path & "actions.def")
            'Events will Follow _ScriptDef
        End If
        
        iAiATM(Index) = 0
        timATM(Index).Enabled = False
    End If
        
End Sub

Private Sub timDied_Timer()

Dim I As Integer
    For I = 0 To MnuCharacters.ListUbound
    
        pBattleHu(I).Alive = True
        pBattleHu(I).Hp = 1
        
        Char(I).ShowStance 0
        MnuCharacters.CharRevived I
        
        MnuCharacters.UpdateHP I, pBattleHu(I).Hp
    Next
    
    StopAudio "boss"
    StopAudio "fiend"

    timDied.Enabled = False
    RaiseEvent BattleFinished(1)

End Sub

Private Sub timFloat_Timer()
    
Dim I As Integer

    For I = 0 To imgEnemy.UBound
    
        'Does the enemy float ?
        If shpShadow(I).Visible = True Then
            
            Select Case iFloatState
            
            Case 0, 1, 2, 3, 4
                imgEnemy(I).Top = imgEnemy(I).Top - 1
            
                shpShadow(I).Width = shpShadow(I).Width - 2
                shpShadow(I).Left = shpShadow(I).Left + 1
                
            Case 5, 6, 7, 8, 9
                imgEnemy(I).Top = imgEnemy(I).Top + 1
            
                shpShadow(I).Width = shpShadow(I).Width + 2
                shpShadow(I).Left = shpShadow(I).Left - 1
            
            End Select
        
        End If
        
    Next
    
    If iFloatState = 9 Then
        iFloatState = 0
    Else
        iFloatState = iFloatState + 1
    End If
    
End Sub

Private Sub timNotify_Timer()

    If sSFXFinished = False Then
        Exit Sub
    End If

    timNotify.Enabled = False

    If Notify.Visible = True Then
        Notify.Visible = False
        NextNotify
    End If

End Sub

Private Sub timShowAttack_Timer()

    Dim imgChar, CharSprite
    
    Set imgChar = Char(iPartyIndex)
    Set CharSprite = Char(iPartyIndex)

    Select Case timShowAttack.Tag
    
    Case 0, 2, 4
        imgChar.Left = imgChar.Left - 4
        CharSprite.ShowStance (1)
        
    Case 1, 3, 5
        imgChar.Left = imgChar.Left - 4
        CharSprite.ShowStance (0)
        
    Case 6, 8
        If bMagic = False Then
            imgSlash.Visible = False
            CharSprite.ShowStance (2)
        Else
            If timShowAttack.Tag = 6 Then
            
                'Pause on stance 9
                CharSprite.ShowStance (9)

                If SFX.LoadScript(sLastSpell & ".fxs") = True Then
                    sSFXFinished = False
                    SFX.Visible = True
                    
                    timShowAttack.Enabled = False
                End If
                
            Else
                timShowAttack.Interval = 80
            End If
        End If
        
    Case 7, 9
        If bMagic = False Then
            imgSlash.Visible = True
            CharSprite.ShowStance (3)
        End If
        
    Case 10, 12, 14
        If timShowAttack.Tag = 10 Then
        
            ShowDelayedDamage
            timTurnFinished.Enabled = True
            
        End If
        
        imgSlash.Visible = False
        CharSprite.ShowStance (4)
        imgChar.Left = imgChar.Left + 4
        
    Case 11, 13, 15
        CharSprite.ShowStance (5)
        imgChar.Left = imgChar.Left + 4
        
    Case 16
        CharSprite.ShowStance (0)
        timShowAttack.Enabled = False
    
    End Select
    
    timShowAttack.Tag = timShowAttack.Tag + 1

End Sub

Sub ShowDelayedDamage()
    imgSlash.Visible = False

    'Show Numbers (Delayed)
    While cNeedsActivation.Count > 0
        imgNumbers(cNeedsActivation(1)).Visible = True
        timAnim(cNeedsActivation(1)).Enabled = True
        Debug.Print "Activating: " & cNeedsActivation(1)
        
        cNeedsActivation.Remove 1
    Wend
End Sub

Sub ShowNumbers(charIndex As Integer, iTarget, lAmount As Long, Optional lSpeed As Integer = 20, Optional bGood As Boolean = False, Optional bMp As Boolean = False)

    'Get Index
    Dim imgStart As Integer
    imgStart = charIndex * 5
    
    '(Timer to enable when ready)
    'timShowDamage.Tag = charIndex * 5

    Dim lLen As Long, I As Integer
    lLen = Len(CStr(lAmount)) - 1

    If bMp = False Then
        'Try to create in center of target
        imgNumbers(imgStart).Top = (iTarget.Top + iTarget.Height) - 8
    Else
        'Try to create center of target, plus a bit more
        imgNumbers(imgStart).Top = (iTarget.Top + iTarget.Height) - 16
    End If
    
    'MsgBox iTarget.Width & ":" & (lLen * 8)
    If iTarget.Width > lLen * 8 Then
        'Center of itarget, Center of Damage
        imgNumbers(imgStart).Left = iTarget.Left + (iTarget.Width / 2) - (((lLen + 1) * 8) / 2)
    Else
        imgNumbers(imgStart).Left = iTarget.Left
    End If
    
    With imgNumbers(imgStart)
        Set .Picture = GetNumero(CLng(Mid(CStr(lAmount), 1, 1)), bGood)
        
        .Tag = "0"
        .ZOrder 0
    End With
    
    With timAnim(imgStart)
        .Interval = lSpeed
    End With
    
    iAniEnd(imgStart) = lLen + 1
    For I = 1 To lLen
        'Positions
        
        'Ubound (on this occusion)
        iAniEnd(imgStart + I) = lLen + 1
        
        With imgNumbers(imgStart + I)
            '.Caption = Mid(CStr(lAmount), I + 1, 1)
            Set .Picture = GetNumero(Mid(CStr(lAmount), I + 1, 1), bGood)
            
            .Left = imgNumbers(.Index - 1).Left + imgNumbers(.Index - 1).Width
            .Top = imgNumbers(.Index - 1).Top
            
            .Tag = "0"
            .ZOrder 0
            'debug.print .Left & " & " & .Top
        End With
        
        With timAnim(imgStart + I)
            .Interval = lSpeed
        End With
    Next
    
    'timAnim(imgStart).Enabled = True
    'bTurnFinished = False

End Sub

Private Sub timAnim_Timer(Index As Integer)
    
Dim I As Integer, iuBound As Integer, iFirst As Integer
Dim uFive As Integer

    uFive = iAniStart(Index)
    iuBound = uFive + iAniEnd(Index)

    imgNumbers(Index).Visible = True

    If timAnim(Index).Tag = "wait" Then
    
        For I = uFive To iuBound
            imgNumbers(I).Visible = False
            
            timAnim(I).Tag = ""
            timAnim(I).Enabled = False
            timAnim(I).Interval = 0
        Next
        
        Exit Sub
    End If

    If imgNumbers(Index).Tag = 7 Then
        If iuBound = Index + 1 Then
            timAnim(Index).Tag = "wait"
            timAnim(Index).Interval = 1000
        Else
            timAnim(Index).Enabled = False
            timAnim(Index).Interval = 0
            
        End If
    End If

    If imgNumbers(Index).Tag < 4 Then
        'Move imgnumbers up
        imgNumbers(Index).Tag = imgNumbers(Index).Tag + 1
        imgNumbers(Index).Top = imgNumbers(Index).Top - 2
        
    ElseIf imgNumbers(Index).Tag > 3 Then

        If (iuBound + 1) > (Index + 1) Then
            timAnim(Index + 1).Enabled = True
        End If
    
        'Move imgnumbers down
        imgNumbers(Index).Tag = imgNumbers(Index).Tag + 1
        imgNumbers(Index).Top = imgNumbers(Index).Top + 2
        
    End If

End Sub

Private Sub timTurnFinished_Timer()

    '2 Seconds from point of Attack

    timTurnFinished.Enabled = False

    NotifyUser "Activating: PreTurnCommands: " & Now

    BeforeTurnFinished

End Sub

Private Sub timWin_Timer()

    If iWinState = 0 Then
        Dim I As Integer
        For I = 0 To Char.UBound
            If pBattleHu(I).Alive = True Then
                Char(I).ShowStance (6)
            End If
        Next
        
        iWinState = 1
    Else
        For I = 0 To Char.UBound
            If pBattleHu(I).Alive = True Then
                Char(I).ShowStance (0)
            End If
        Next
    
        iWinState = 0
    End If

End Sub

Sub PauseAIATB()

Dim I As Integer

    For I = 0 To MnuEnemies.ListCount
        If timATM(I).Enabled = False Then
            timATM(I).Tag = "R"
        Else
            timATM(I).Enabled = False
        End If
    Next

End Sub

Sub ResumeAIATB()

Dim I As Integer

    For I = 0 To MnuEnemies.ListCount
        'Renable AI ATB's
        If timATM(I).Tag <> "R" Then
            timATM(I).Enabled = True
        End If
    Next

End Sub

Private Sub UserControl_Initialize()

    InitialiseKeys
    keys.Enabled = False

    sSFXFinished = True
    Set ScriptDef = New clsScript

    If NotInDbg = False Then
        Exit Sub
    End If
    
    imgNumbers(0).MaskColor = lTrans
    Set imgNumbers(0).Picture = GetNumero(0)
    
    Dim I As Integer
        For I = 1 To 39
                'HP
        
            If I < 5 Then
                iAniStart(I) = 0
            ElseIf I < 10 Then
                iAniStart(I) = 5
            ElseIf I < 15 Then
                iAniStart(I) = 10
            ElseIf I < 20 Then
                iAniStart(I) = 15
                
                'MP
            ElseIf I < 25 Then
                iAniStart(I) = 20
            ElseIf I < 30 Then
                iAniStart(I) = 25
            ElseIf I < 35 Then
                iAniStart(I) = 30
            ElseIf I < 40 Then
                iAniStart(I) = 35
            End If
            
            Load timAnim(I)
            Load imgNumbers(I)
            imgNumbers(I).MaskColor = lTrans
        Next

    FormatField

End Sub
