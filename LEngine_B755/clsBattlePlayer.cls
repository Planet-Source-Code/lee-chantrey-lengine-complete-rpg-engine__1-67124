VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBattlePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim bSteal As Boolean

Dim iStr As Integer
Dim iStr_Temp As Integer

Dim iDef As Integer
Dim iDef_Temp As Integer

Dim iMag As Integer
Dim iMag_Temp As Integer

Dim iSpr As Integer
Dim iSpr_Temp As Integer

Dim iHp As Integer

Dim iMpoint As Integer
Dim iMaxMp As Integer

Dim iIndex As Integer

Dim iMaxHp As Integer
Dim iMaxHp_Temp As Integer

Dim iAtb As Integer
Dim iAtb_Temp As Integer

Dim sName As String
Dim bEnemy As Boolean
Dim bAlive As Boolean
Dim iExp As Single
Dim iLvl As Integer
Dim sId As String
Dim iElement As Integer
Dim bReflect As Boolean
Dim sPath As String

Dim bBoss As Boolean

Dim cActions As New Collection
Dim cSteals As New Collection

Dim cNodes() As Collection

Dim cBattleStat As New Collection
Dim bTempStats As Boolean
Dim Img

Dim bDef As Boolean

Public Property Let Frame(iNewFrame As Integer)
    Img.ShowStance iNewFrame
End Property

Public Property Let Visible(bNewVisible As Boolean)
    Img.Visible = bNewVisible
End Property

Public Property Get Visible() As Boolean
    Visible = Img.Visible
End Property

Public Property Let X(iNewX As Single)
    If IsObject(Img) = False Then
        Exit Property
    End If
    
    Img.Left = iNewX
End Property

Public Property Get X() As Single
    If IsObject(Img) = False Then
        Exit Property
    End If

    X = Img.Left
End Property

Public Property Let Y(iNewY As Single)
    If IsObject(Img) = False Then
        Exit Property
    End If

    Img.Top = (240 - iNewY) - Img.Height
End Property

Public Property Get Y() As Single
    If IsObject(Img) = False Then
        Exit Property
    End If

    Y = (240 - Img.Top) - Img.Height
End Property

Public Property Get Height()
    If IsObject(Img) = False Then
        Exit Property
    End If

    Height = Img.Height
End Property

Public Property Get Width()
    If IsObject(Img) = False Then
        Exit Property
    End If

    Width = Img.Width
End Property

Public Property Let Image(imgNew)
    Set Img = imgNew
End Property

Public Property Let Boss(bNewBoss As Boolean)
    bBoss = bNewBoss
End Property

Public Property Get Boss() As Boolean
    Boss = bBoss
End Property

Public Property Let Element(newElement As Integer)
    iElement = newElement
End Property

Public Property Get Element() As Integer
    Element = iElement
End Property

Public Property Get ATB() As Integer
    
    If bTempStats = True Then
        ATB = iAtb_Temp
    Else
        ATB = iAtb
    End If
        
End Property

Public Property Let ATB(iNewATB As Integer)
    
    If bTempStats = True Then
        If iAtb_Temp < iNewATB Then
            'Has Decreased (Good)
            cPosBattleChanges.Add "atb:" & iIndex & ":" & iNewATB - iAtb_Temp & ":" & bEnemy
        Else
            'Has Increased (Bad)
            cNegBattleChanges.Add "atb:" & iIndex & ":" & iAtb_Temp - iNewATB & ":" & bEnemy
        End If
    
        iAtb_Temp = iNewATB
    Else
        iAtb = iNewATB
    End If
    
    'Add to Pos Changes anyway (No Numbers occurs
    cPosBattleChanges.Add "atb:" & iIndex & ":" & iNewATB & ":" & bEnemy
        
End Property

Public Property Let Path(sNewPath As String)
    sPath = sNewPath
End Property

Public Property Get Path() As String
    Path = sPath
End Property

Public Property Let HasDef(bNewDef As Boolean)
    bDef = bNewDef
End Property

Public Property Get HasDef() As Boolean
    HasDef = bDef
End Property

Public Function CleanStats()
    bTempStats = False
End Function

Public Property Let ReflectStatus(bNewReflect As Boolean)
    bReflect = bNewReflect
End Property

Public Property Get ReflectStatus() As Boolean
    ReflectStatus = bReflect
End Property

Public Property Let TempStats(newTempStat As Boolean)

    bTempStats = newTempStat

    If bTempStats = True Then
        'Copy Stats to Temp
        iDef_Temp = iDef
        iMaxHp_Temp = iMaxHp
        iAtb_Temp = iAtb

        iStr_Temp = iStr
        iMag_Temp = iMag
        iSpr_Temp = iSpr
    End If

End Property

Public Property Let MaxMp(newMax As Integer)
    iMaxMp = newMax
End Property

Public Property Get MaxMp() As Integer
    MaxMp = iMaxMp
End Property

Public Property Get Mp() As Integer
    Mp = iMpoint
End Property

Public Property Let Mp(iNewMp As Integer)

    If iMaxMp < iNewMp Then
        iNewMp = iMaxMp
    End If

    If iMpoint < iNewMp Then
        cPosBattleChanges.Add "mp:" & iIndex & ":" & iNewMp - iMpoint & ":" & bEnemy
    Else
        cNegBattleChanges.Add "mp:" & CInt(iIndex) & ":" & iMpoint - iNewMp & ":" & bEnemy
    End If
    
    iMpoint = iNewMp
    
End Property

Public Property Let ID(newID As String)
    sId = newID
End Property

Public Property Get ID() As String
    ID = sId
End Property

Public Property Get ActionNode(Index As Integer) As Collection
    Set ActionNode = cNodes(Index)
End Property

Public Property Let ActionNode(Index As Integer, cNewNode As Collection)
    Set cNodes(Index) = cNewNode
End Property

Public Property Get Steal() As Boolean
    Steal = bSteal
End Property

Public Property Let Steal(bNewSteal As Boolean)
    bSteal = bNewSteal
End Property

Public Property Get Actions() As Collection
    Set Actions = cActions
End Property

Public Property Let Actions(cNewActions As Collection)
    Set cActions = cNewActions
    
    If cActions.Count > 0 Then
        'Nodes of action redefine
        ReDim cNodes(cActions.Count)
    End If
End Property

Public Property Get Steals() As Collection
    Set Steals = cSteals
End Property

Public Property Let Steals(cNewSteals As Collection)
    Set cSteals = cNewSteals
End Property

Public Property Let Level(iNewLvl As Integer)
    iLvl = iNewLvl
End Property

Public Property Get Level() As Integer
    Level = iLvl
End Property

Public Property Let Experience(iNewExp As Single)
    iExp = iNewExp
End Property

Public Property Get Experience() As Single
    Experience = iExp
End Property

Public Property Let Strength(iNewStr As Integer)
    
    If bTempStats = True Then
    
        If iStr_Temp < iNewStr Then
            'Has Decreased (Good)
            cPosBattleChanges.Add "str:" & iIndex & ":" & iNewStr - iStr_Temp & ":" & bEnemy
        Else
            'Has Increased (Bad)
            cNegBattleChanges.Add "str:" & iIndex & ":" & iStr_Temp - iNewStr & ":" & bEnemy
        End If
    
        'If in battle, DONT actually change stats
        iStr_Temp = iNewStr
    Else
        iStr = iNewStr
    End If
        
End Property

Public Property Get Strength() As Integer

    If bTempStats = True Then
        'If in battle, DONT actually get stats
        Strength = iStr_Temp
    Else
        'Get Perm Stat Value
        Strength = iStr
    End If
    
End Property

Public Property Let Magic(iNewMag As Integer)

    If bTempStats = True Then
    
        If iMag_Temp < iNewMag Then
            'Has Decreased (Good)
            cPosBattleChanges.Add "mag:" & iIndex & ":" & iNewMag - iMag_Temp & ":" & bEnemy
        Else
            'Has Increased (Bad)
            cNegBattleChanges.Add "mag:" & iIndex & ":" & iMag_Temp - iNewMag & ":" & bEnemy
        End If
    
        'If in battle, DONT actually change stats
        iMag_Temp = iNewMag
    Else
        iMag = iNewMag
    End If

End Property

Public Property Get Magic() As Integer

    If bTempStats = True Then
        'If in battle, DONT actually change stats
        Magic = iMag_Temp
    Else
        'Get Perm Stat Value
        Magic = iMag
    End If
    
End Property

Public Property Let Spirit(iNewSpr As Integer)

    If bTempStats = True Then
    
        If iSpr_Temp < iNewSpr Then
            'Has Decreased (Good)
            cPosBattleChanges.Add "spr:" & iIndex & ":" & iNewSpr - iSpr_Temp & ":" & bEnemy
        Else
            'Has Increased (Bad)
            cNegBattleChanges.Add "spr:" & iIndex & ":" & iSpr_Temp - iNewSpr & ":" & bEnemy
        End If
    
        'If in battle, DONT actually change stats
        iSpr_Temp = iNewSpr
    Else
        iSpr = iNewSpr
    End If
    
End Property

Public Property Get Spirit() As Integer

    If bTempStats = True Then
        'If in battle, DONT actually get stats
        Spirit = iSpr_Temp
    Else
        'Get Perm Stat Value
        Spirit = iSpr
    End If
    
End Property

Public Property Let Defence(iNewDef As Integer)

    If bTempStats = True Then
    
        If iDef_Temp < iNewDef Then
            'Has Decreased (Good)
            cPosBattleChanges.Add "def:" & iIndex & ":" & iNewDef - iDef_Temp & ":" & bEnemy
        Else
            'Has Increased (Bad)
            cNegBattleChanges.Add "def:" & iIndex & ":" & iDef_Temp - iNewDef & ":" & bEnemy
        End If
    
        'If in battle, DONT actually change stats
        iDef_Temp = iNewDef
    Else
        'Change stats PERM
        iDef = iNewDef
    End If
    
End Property

Public Property Get Defence() As Integer

    If bTempStats = True Then
        'If in battle, DONT actually change stats
        Defence = iDef_Temp
    Else
        'Get Perm Stat Value
        Defence = iDef
    End If

End Property

Public Property Let MaxHp(iNewHP As Integer)

    If bTempStats = True Then
        'Dont change stats
        iMaxHp_Temp = iNewHP
    Else
        'Change stats
        iMaxHp = iNewHP
    End If
    
End Property

Public Property Get MaxHp() As Integer
    
    If bTempStats = True Then
        'Get fake stat
        MaxHp = iMaxHp_Temp
    Else
        'Get real stat
        MaxHp = iMaxHp
    End If
        
End Property

Public Property Let Hp(iNewHP As Long)

    'Cant change HP if dead
    If Me.Alive = False And iNewHP <> 0 Then
        Exit Property
    End If

    If Me.MaxHp < iNewHP Then
        iNewHP = Me.MaxHp
    End If
    
    If iNewHP <= 0 Then
        Me.Alive = False
    End If

    If iHp < iNewHP Then
        cPosBattleChanges.Add "hp:" & iIndex & ":" & iNewHP - iHp & ":" & bEnemy
    Else
        cNegBattleChanges.Add "hp:" & CInt(iIndex) & ":" & iHp - iNewHP & ":" & bEnemy
    End If
    
    'Set to 0HP, instead of -
    If iNewHP < 0 Then
        iNewHP = 0
    End If
        
    iHp = iNewHP
    
End Property

Public Property Get Hp() As Long
    Hp = iHp
End Property

Public Property Let Index(iNewIndex As Integer)
    iIndex = iNewIndex
End Property

Public Property Get Index() As Integer
    Index = iIndex
End Property

Public Property Let Name(sNewName As String)
    sName = sNewName
End Property

Public Property Get Name() As String
    Name = sName
End Property

Public Property Let Enemy(B As Boolean)
    bEnemy = B
End Property

Public Property Get Enemy() As Boolean
    Enemy = bEnemy
End Property

Public Property Let Alive(B As Boolean)

    If bAlive = False Then
        If B = True Then
            cPosBattleChanges.Add "life:" & iIndex & ":" & B & ":" & bEnemy
        End If
    Else
        If B = False Then
            cNegBattleChanges.Add "life:" & iIndex & ":" & B & ":" & bEnemy
        End If
    End If

    bAlive = B
End Property

Public Property Get Alive() As Boolean
    Alive = bAlive
End Property

Private Sub Class_Initialize()
    'Set Defaults
    Me.Steal = False
End Sub
