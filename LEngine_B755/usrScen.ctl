VERSION 5.00
Begin VB.UserControl usrScen 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   ControlContainer=   -1  'True
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   Begin prjLEngine.usrTransition Transition 
      Height          =   3615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
   End
   Begin VB.Timer timBounce2 
      Left            =   9240
      Top             =   600
   End
   Begin VB.Timer timBounce 
      Enabled         =   0   'False
      Interval        =   9
      Left            =   9240
      Top             =   120
   End
   Begin prjLEngine.usrPrologue pPrologue 
      Height          =   3600
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin VB.PictureBox picMapC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   180
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   2940
      Width           =   480
      Begin VB.PictureBox picMap 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   0
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   10
         Top             =   0
         Width           =   495
         Begin VB.PictureBox picChars 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   30
            Index           =   0
            Left            =   0
            ScaleHeight     =   30
            ScaleWidth      =   30
            TabIndex        =   11
            Top             =   0
            Width           =   30
         End
      End
   End
   Begin VB.Timer timFrameProp 
      Enabled         =   0   'False
      Index           =   0
      Left            =   9720
      Top             =   120
   End
   Begin VB.TextBox txtOut2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3600
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "usrScen.ctx":0000
      Top             =   2640
      Width           =   4800
   End
   Begin prjLEngine.usrNotify Notify 
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   423
   End
   Begin prjLEngine.usrMenu Options 
      Height          =   720
      Left            =   1800
      TabIndex        =   5
      Top             =   1365
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1270
   End
   Begin prjLEngine.usrSpeech Speech 
      Height          =   780
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   1376
   End
   Begin prjLEngine.usrTransPic imgS_Top 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1508
      MaskColor       =   16777215
   End
   Begin VB.Timer timTerm 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9720
      Top             =   600
   End
   Begin VB.Timer timSleep 
      Enabled         =   0   'False
      Left            =   9720
      Top             =   1080
   End
   Begin prjLEngine.usrTransPic imgChars 
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      MaskColor       =   -2147483633
   End
   Begin prjLEngine.usrChar Chars 
      Index           =   0
      Left            =   0
      Tag             =   "Self"
      Top             =   0
      _ExtentX        =   1296
      _ExtentY        =   661
   End
   Begin prjLEngine.keyReciever keyReciever1 
      Left            =   4200
      Top             =   3960
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin prjLEngine.usrTransPic imgCur 
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      MaskColor       =   16777215
   End
   Begin VB.PictureBox imgS 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   0
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   7
      Top             =   0
      Width           =   4800
      Begin prjLEngine.usrTransPic staticProps 
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   450
         MaskColor       =   -2147483633
      End
   End
   Begin VB.Image BattleBack 
      Height          =   675
      Left            =   9000
      Top             =   4320
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "usrScen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
    ByVal Y As Long) As Long
    
Private Const State_NoWait As Integer = 0
Private Const State_Speech As Integer = 1
Private Const State_AsyncSpeech As Integer = 2
Private Const State_Notify As Integer = 3
Private Const State_Option As Integer = 4
Private Const iProp_Ubound As Integer = 2

Private iSystem_Exlem As Integer

Private Const charSpeed As Integer = 3
Private Char As usrChar
Private iPic As IPictureDisp

Private sMusic As String
Private sScenPath As String

Private cWalkListX As New SuperCollection 'Char Walk Ques, X
Private cWalkListY As New SuperCollection 'Char Walk Ques, Y

Private cEvents As New SuperCollection  'EventSquares
Private cEventsA As New SuperCollection 'ActionSquares '[Triggerd by Action Key]

'[Triggerd by Action Key]
Private cCollisionEvents As New SuperCollection 'Prop Collision Events
Private cAsyncActions As New SuperCollection 'NPC Actions

Private cDoors As New Collection
Private iWalkChance As Integer
Private iFreeSteps As Integer

Private iFiendChance As Integer
Private iBounceCharID As Integer

Private bIgnoreEvent As Boolean

Public Event OnFiend()
Public Event OnEquip(charIndex As Integer)

Public Sub SetFiendChance(iChance As Integer, iFree As Integer)

    iFiendChance = iChance + iFree
    iFreeSteps = iFree

End Sub

Property Get keys() As keyReciever
    Set keys = keyReciever1
End Property

Property Get ScenPath() As String
    ScenPath = sScenPath
End Property

Property Let ScenPath(sNewScen As String)
    sScenPath = sNewScen
    
    LoadScenFile
End Property

Public Function ActivateTerm()

    timTerm_Timer

End Function

Property Let MusicPath(sNewMusic As String)
    sMusic = sNewMusic
End Property

Sub Answers(COpts As Collection, CAns As Collection)

    Options.ClearList

    While COpts.Count > 0
        Options.AddItem COpts(1), CAns(1)
        
        COpts.Remove 1
        CAns.Remove 1
    Wend

    Options.Visible = True
    
    imgCur.Visible = True
    imgCur.ZOrder 0
    
    Options.AttachCur imgCur, Options.Left, Options.Top, 1
    
    User_State = State_Option

End Sub

Sub AmendInventory(sItem As String, bKeyItem As Boolean)

Dim sDesc As String
    
    If bKeyItem = False Then
        sDesc = "Recieved " & """" & sItem & """"
    Else
        sDesc = "Recieved Key Item " & """" & sItem & """"
    End If

    NotifyText sDesc

End Sub

Public Sub BounceChar(sChar As String, iSpeed As Integer)

    HaltStory

    iBounceCharID = getCharByName(sChar)
    
    If iBounceCharID = -1 Then
        MsgBox "BounceChar::Invalid CharID (" & sChar & ")"
        Exit Sub
    End If
        
    timBounce.Interval = iSpeed
    timBounce.Enabled = True

End Sub

Public Sub BounceCharExclem(sChar As String, iSpeed As Integer)

    HaltStory

    iBounceCharID = getCharByName(sChar)
    
    With staticProps(iSystem_Exlem)
        .Top = (((picChars(iBounceCharID).Top) * 8) - 16) - 1
        .Left = (picChars(iBounceCharID).Left) * 8

        .Visible = True
    End With
    
    timBounce2.Interval = iSpeed
    timBounce2.Enabled = True

End Sub

Public Sub NotifyText(sCaption As String)

    HaltStory

    Notify.Caption = sCaption
    Notify.Visible = True
    
    User_State = State_Notify
    
End Sub

Sub StoryVoice(sCaption As String, Optional bHalt As Boolean = True)

    If bHalt = True Then
        HaltStory
    End If

    User_State = State_Speech
    
    Speech.Visible = True
    Speech.Say "", sCaption

End Sub

Sub SmallDoorY(sName As String, sPath As String, X As Integer, Y As Integer)

    Dim PropIndex As Integer
    PropIndex = StaticProp(sName, sPath, X, Y, False)
    
    '+ 2 For Height
    Y = ((picMap.ScaleHeight / 2) - (Y + 2))
    
    RemoveCol cDoors, X & ":" & Y
    cDoors.Add X & ":" & Y & ":" & PropIndex, X & ":" & Y

End Sub

Sub HideProp(sName As String)

Dim PropIndex As Integer
    PropIndex = GetPropByName(sName)
    
    If PropIndex = -1 Then
        WarnUser "HideProp Failed: " & sName & " Does not exist."
        Exit Sub
    End If
    
    staticProps(PropIndex).Visible = False

End Sub

Sub ShowProp(sName As String)

Dim PropIndex As Integer
    PropIndex = GetPropByName(sName)
    
    If PropIndex = -1 Then
        WarnUser "ShowProp Failed: " & sName & " Does not exist."
        Exit Sub
    End If
    
    staticProps(PropIndex).Visible = True

End Sub

Sub StartProp(sName As String)

Dim timIndex As Integer
    timIndex = getFramePropTimerByName(sName)
    
    If timIndex = -1 Then
        WarnUser "StartProp Failed: " & sName & " Does not exist."
        Exit Sub
    End If
    
    timFrameProp(timIndex).Enabled = True

End Sub

Sub StopProp(sName As String)

Dim timIndex As Integer
    timIndex = getFramePropTimerByName(sName)
    
    If timIndex = -1 Then
        WarnUser "StopProp Failed: " & sName & " Does not exist."
        Exit Sub
    End If
    
    timFrameProp(timIndex).Enabled = False

End Sub

Function FrameProp(ByVal sName As String, ByVal sPath As String, ByVal iInterval As Integer, X As String, Y As Integer, bLoop As Boolean)

Dim PropIndex As Integer, timIndex As Integer

    On Error Resume Next

    PropIndex = GetPropByName(sName)
    If PropIndex = -1 Then
        PropIndex = staticProps.Count
        
        Load staticProps(PropIndex)
    End If
    
    timIndex = getFramePropTimerByName(CStr(PropIndex))
    If timIndex = -1 Then
        timIndex = timFrameProp.Count
        
        Load timFrameProp(timIndex)
    End If

    sPath = sPath_Story & sPath
    
    X = X * 2
    Y = Y * 2

    With staticProps(PropIndex)
        .Tag = sName
        .LoadStore sPath, FileCount(sPath, "a*.bmp")
        
        .Visible = True
        .AutoSize = True
        
        .Left = X * 8
        ' - ScaleHeight
        .Top = (imgS.ScaleHeight - (Y * 8)) - (.Height)
        
        .MaskColor = lTrans
        
        .FrameLoop = bLoop
    End With

    With timFrameProp(timIndex)
        .Interval = iInterval
        .Tag = PropIndex
        
        .Enabled = True
    End With
    
    If Err Then
        WarnUser "FrameProp Failed: " & Err.Description
    Else
        FrameProp = PropIndex
    End If

End Function

Function StaticProp(ByVal sName As String, ByVal sPath As String, ByVal X As Integer, ByVal Y As Integer, Optional bVisible As Boolean = True)
    
Dim PropIndex As Integer
    PropIndex = GetPropByName(sName)
    
    If PropIndex = -1 Then
        PropIndex = staticProps.Count
        Load staticProps(PropIndex)
    End If
    
    On Error Resume Next
    'sPath = sPath_Story & sPath
    
    X = X * 2
    Y = Y * 2

    With staticProps(PropIndex)
        .Tag = sName
        Set .Picture = LoadPicture(sPath)
        
        .Left = X * 8
        
        
        .Visible = bVisible
        .AutoSize = True
        
        ' - ScaleHeight
        .Top = (imgS.ScaleHeight - (Y * 8)) - (.Height)
        
        .MaskColor = lTrans
    End With
    
    If Err Then
        WarnUser "StaticProp Failed: " & Err.Description
    Else
        StaticProp = PropIndex
    End If
        
End Function

Private Function MoveProp(sName As String, X As Integer, Y As Integer)

Dim PropIndex As Integer
    PropIndex = GetPropByName(sName)
    
    If PropIndex = -1 Then
        WarnUser "MoveProp: Cannot find prop"
    End If
    
    With staticProps(PropIndex)
        .Left = X * 8
        .Top = (imgS.ScaleHeight - (Y * 8)) - (.Height)
    End With

End Function

Function AsyncAction(sName As String, sAction As String)
    cAsyncActions(Num2Col(getCharByName(sName))) = sAction
End Function

Function AsyncCollision(sName As String, sCmd As String)
    cCollisionEvents.Add sCmd, sName
End Function

Private Function DoAsyncEvents(sKey As String) As Boolean

    ExeStoryCmd cAsyncActions(sKey)
    cAsyncActions(sKey) = ""

End Function

Public Function ReplaceImg(sName As String, sPath As String)

    On Error Resume Next

Dim charIndex As Integer
    charIndex = getCharByName(sName)
    
    If charIndex > -1 Then
        Set imgChars(charIndex).Picture = LoadPicture(sPath_Resources & sPath)
    Else
        Set staticProps(GetPropByName(sName)).Picture = LoadPicture(sPath_Resources & sPath)
    End If

    If Err Then
        If Err.Number = 53 Then
            WarnUser "ReplaceImg(): Picture could not be found. Check Path and try again."
        End If
    End If
    
    If Err Then
        WarnUser "ReplaceImg failed: " & Err.Description
    End If

End Function

Public Function SetBGColour(WebColour As String)

    UserControl.BackColor = Hex2VB(WebColour)

End Function

Function SayText(ByVal sName As String, sText As String, bHalt As Boolean)

Dim charIndex As Integer

    If bHalt = True Then
        HaltStory
    End If

    charIndex = getCharByName(sName)
    
    If charIndex = -1 Then
        WarnUser "Say Failed: " & sName & " could not be found."
        Exit Function
    End If
    
    User_State = State_Speech
    
    Speech.Visible = True
    
    'sName = ColValue(Aliases, sName, "????")
    sName = Aliases.GetItem(sName, "????")
    
    Speech.Say Chars(charIndex).srPath, sName & ": " & sText

End Function

Function SetEventSquare(X As Double, Y As Double, sStory As String, bExpire As Boolean, bHalt As Boolean)
    'When character walks on this square, execute that command
    
    If bHalt = True Then
        HaltStory
    End If
    
    Y = ((picMap.ScaleHeight / 2) - (Y + 1))
    
    If cEvents.Exists(CStr(X & ":" & Y)) Then
        cEvents.Remove CStr(X & ":" & Y)
    End If
    
    cEvents.Add X & ":" & Y & ":" & sStory & ":" & CStr(BolAsNO(bExpire)), X & ":" & Y

End Function

Function SetActionSquare(X As Integer, Y As Integer, sCommand As String, iVisible As Integer)
    'When character presses Activate on this square, execute that command
    
    Y = ((picMap.ScaleHeight / 2) - (Y + 1))
    
    'Scale to usable
    Y = Y * 2
    X = X * 2
    
    If cEventsA.Exists(CStr(X & ":" & Y)) Then
        cEventsA.Remove CStr(X & ":" & Y)
    End If

    cEventsA.Add X & ":" & Y & ":" & sCommand & ":" & iVisible, X & ":" & Y

End Function

Function LoadNPC(sId As String, sName As String, sPath As String) As Integer
    
Dim charIndex As Integer
    
    On Error GoTo Catch_E
    
    sPath = sPath_StoryChars & sPath
    
    charIndex = getCharByName(sId)
    
    If charIndex = 0 Then
        Exit Function
    End If

    If Aliases.Exists(sId) = True Then
        Aliases(sId) = sName
        
        NotifyUser sId & ": alias was updated."
    End If
    
    Aliases.Add sName, sId
    Op_Variable sId, sName
    
    If charIndex = -1 Then
        'Load New
        charIndex = Chars.Count
    
        Load Chars(charIndex)
        Load imgChars(charIndex)
        Load picChars(charIndex)
        
        cAsyncActions.Add "", Num2Col(charIndex)
    End If

    With picChars(charIndex)
        .Top = 0
        .Left = 0
        
        .Visible = False
    End With
    
    With imgChars(charIndex)
        .Top = 0
        .Left = 0
    
        .Visible = False
    End With

    SkinChar charIndex, sId, sPath
    
    Exit Function
    
Catch_E:
    WarnUser "LoadNPC Failed: " & Err.Description

End Function

Function SkinChar(charIndex As Integer, sId As String, sPath As String)

    If charIndex = -1 Then
        WarnUser "SkinChar Failed: Character non existant", True
        Exit Function
    End If
    
    On Error GoTo Catch_E

    With picChars(charIndex)
        .Picture = LoadPicture(sPath & "\map.bmp")
    End With

    With Chars(charIndex)
        .Tag = sId
        .srPath = sPath
        .LoadFrames sPath
    End With
    
    With imgChars(charIndex)
        .MaskColor = lTrans
        .Tag = ""
        Set .Container = imgS
    End With
    
    Exit Function
    
Catch_E:
    WarnUser "SkinChar Failed: " & Err.Description

End Function

Function WalkCharX(sName As String, X As Integer, Optional iSpeed As Integer = 8, Optional bHalt As Boolean = True)

    If bHalt = True Then
        HaltStory
    End If
    
    If cWalkListX.Exists(sName) = True Then
        cWalkListX.Remove sName
    End If
    
    cWalkListX.Add sName & ":" & X & ":" & iSpeed & ":" & CStr(bHalt), sName
    
End Function

Function WalkCharY(sName As String, Y As Integer, Optional iSpeed As Integer = 8, Optional bHalt As Boolean = True)

    If bHalt = True Then
        HaltStory
    End If

    'Get the REAL Y (Not the Human one)
    Y = ((picMap.ScaleHeight / 2) - (Y + 1))
    
    If cWalkListY.Exists(sName) = True Then
        cWalkListY.Remove sName
    End If
    
    cWalkListY.Add sName & ":" & Y & ":" & iSpeed & ":" & CStr(bHalt), sName
    
End Function

Function ClearScen()

    Set imgS.Picture = Nothing
    Set picMap.Picture = Nothing

    timTerm_Timer

End Function

Public Function Dump(ByRef iniChars As clsIniObj)

    'Dump everything on the Scen, into INI format

Dim I As Integer, sKeys As String

    iniChars.Section = "Story_WalksX"
    DumpSuperCollection iniChars, cWalkListX
    
    iniChars.Section = "Story_WalksY"
    DumpSuperCollection iniChars, cWalkListY
    
    iniChars.Section = "EventSquares"
    DumpSuperCollection iniChars, cEvents
    
    iniChars.Section = "ActionSquares"
    DumpSuperCollection iniChars, cEventsA
    
    iniChars.Section = "CollisionEvents"
    DumpSuperCollection iniChars, cCollisionEvents
    
    iniChars.Section = "AsyncActions"
    DumpSuperCollection iniChars, cAsyncActions

    For I = 0 To imgChars.UBound
        iniChars.Section = "ScenChar_" & Chars(I).Tag
    
        With imgChars(I)
            iniChars.WriteData CStr(CharY(picChars(I))), "Y"
            iniChars.WriteData CStr(CharX(picChars(I))), "X"
        End With
        
        With Chars(I)
            iniChars.WriteData StrEnd(.srPath, "\"), "srPath"
            iniChars.WriteData Aliases.GetItem(Chars(I).Tag, "??"), "Name"
            iniChars.WriteData Chars(I).LastDirection, "LastDir"
            iniChars.WriteData imgChars(I).Visible, "Visible"
        End With
        
        sKeys = sKeys & Chars(I).Tag & ","
    Next
    sKeys = Mid(sKeys, 1, Len(sKeys) - 1)
    
    iniChars.WriteData sKeys, "Chars", "Story"

End Function

Public Function RestoreChars(ByRef bcIni As clsIniObj)

Dim sChars() As String, I As Integer

    bcIni.Section = "Story"
    sChars = Split(bcIni.Read("Chars"), ",")
    
    For I = 0 To UBound(sChars)
        bcIni.Section = "ScenChar_" & sChars(I)

        LoadNPC sChars(I), bcIni.Read("Name"), bcIni.Read("srPath")
        PositionChar sChars(I), bcIni.Read("X"), bcIni.Read("Y"), bcIni.Read("LastDir")
        
        If bcIni.Read("Visible") Then
            ExeStoryCmd "ShowChar '" & sChars(I) & "'"
        End If
    Next

    bcIni.Section = "Story_WalksY"
    RestoreSuperCollection bcIni, cWalkListY

    bcIni.Section = "Story_WalksX"
    RestoreSuperCollection bcIni, cWalkListX
    
    bcIni.Section = "EventSquares"
    RestoreSuperCollection bcIni, cEvents

    bcIni.Section = "ActionSquares"
    RestoreSuperCollection bcIni, cEventsA
    
    bcIni.Section = "CollisionEvents"
    RestoreSuperCollection bcIni, cCollisionEvents
    
    bcIni.Section = "AsyncActions"
    RestoreSuperCollection bcIni, cAsyncActions

End Function

Function StartTransistion()

    HaltStory
    
    keyReciever1.Enabled = False
    Transition.Grow

End Function

Private Function LoadScenFile(Optional bTrans As Boolean = True) As Integer

    On Error GoTo CatchErr

    Set cStory_ToDo = New Collection
    
    ClearReport
    
    Echo "Loading...."
    
    'this stops Game Timer
    keyReciever1.Enabled = False
    
    timSleep.Enabled = False

    imgS.Picture = LoadPicture(sScenPath & "\scen.bmp")
    
    If FileExist(sScenPath & "\topscen.bmp") Then
        Set imgS_Top.Picture = LoadPicture(sScenPath & "\topscen.bmp")
        
        imgS_Top.AutoSize = True
        imgS_Top.Visible = True
    Else
        imgS_Top.Visible = False
    End If
    
    picMap.Picture = LoadPicture(sScenPath & "\map.bmp")
    
    If picMap.ScaleHeight Mod 2 Or picMap.ScaleWidth Mod 2 Then
        WarnUser "Map.bmp is not 2BIT."
        End
    End If
    
    If FindExistingMusic(sMusic, "music") = True Then
        PlayNewMusic sMusic
    End If
    
    Exit Function
    
CatchErr:
    WarnUser "LoadScenFile(): " & Err.Description

    LoadScenFile = 1

End Function

Sub MoveTopImg()

    imgS_Top.Move imgS.Left, imgS.Top, imgS.Width, imgS.Height

End Sub

Private Sub ReposistionMainChar(X As Integer, Y As Integer, lDir As Integer)

    'HaltStory

    'Only Utililise big square
    X = X * 2
    Y = Y * 2
    
    '-2 (Because of the height of the dot represnting the character)
    picChars(0).Top = (picMap.ScaleHeight - (Y)) - 2
    picChars(0).Left = (X)
    'picChars(0).ZOrder 0
    
    picMap.Move 0, 0
    picMap.Top = picMap.Top - (picChars(0).Top) + 15
    picMap.Left = picMap.Left - (X) + 15

    imgS.Top = -(imgS.ScaleHeight) + 120 + _
                (Y * 8) + 8
    imgS.Left = 160 - _
                (X * 8) - 8
    MoveTopImg

End Sub

Sub Start(ByVal X As Integer, ByVal Y As Integer, ByVal lDir As Integer)

    ReposistionMainChar X, Y, lDir
    
    imgChars(0).Top = (120 - 8)
    imgChars(0).Left = (160 - 8)
    imgChars(0).Visible = True
    imgChars(0).Tag = "Self"

    picChars(0).BackColor = vbRed
    
    Chars(0).srPath = sPath_Story & "characters\main.chr"
    Chars(0).LoadFrames sPath_Story & "characters\main.chr"
    
    Set imgChars(0).Picture = Chars(0).ShowWalk(lDir)
    picChars(0).Visible = True
    
    If Transition.Shrink = True Then
        HaltStory
    Else
        NotifyUser "Respositioning the main character in this context is 'dirty'"
    End If
    
End Sub

Sub Walk(lDir As Long)


    Select Case lDir
    
    Case 1
        imgS.Left = imgS.Left + charSpeed
    
    End Select


End Sub

Function MoveChar(lDir As Integer, Optional bOverideControl As Boolean = False, Optional iSpeed As Integer = 8) As Integer

Static llastDir As Long
Dim X As Integer, Y As Integer, bMove As Boolean, lResX As Long, lResY As Long

    MoveChar = 0

    If bOverideControl = False And Player_Control = False Then
        Exit Function
    End If

    X = picChars(0).Left: Y = picChars(0).Top

    If llastDir <> lDir Then
        'Show Anim
        Set imgChars(0).Picture = Chars(0).ShowWalk(lDir)
    End If
    
    llastDir = lDir

    Select Case lDir
    
    Case 2
        'Move Right
        If (GetPixel(picMap.hdc, X + 2, Y) = 16777215 And _
            GetPixel(picMap.hdc, X + 2, Y + 1) = 16777215) Or _
                picChars(0).Visible = False Then
            
            bMove = True
            
            imgS.Left = imgS.Left - iSpeed
            MoveTopImg
            
            If imgS.Left Mod 8 = 0 Then
                picChars(0).Left = picChars(0).Left + 1
                picMap.Left = picMap.Left - 1
            End If
        Else
            
        End If
        
    Case 1
        'Move Left
        If (GetPixel(picMap.hdc, X - 1, Y) = 16777215 And _
            GetPixel(picMap.hdc, X - 1, Y + 1) = 16777215) Or _
                picChars(0).Visible = False Then
           
            bMove = True

            imgS.Left = imgS.Left + iSpeed
            MoveTopImg
            
            If imgS.Left Mod 8 = 0 Then
                picChars(0).Left = picChars(0).Left - 1
                picMap.Left = picMap.Left + 1
            End If
        Else: End If
        
    Case 3
        'Move Down
        If (GetPixel(picMap.hdc, X, Y + 2) = 16777215 And _
            GetPixel(picMap.hdc, X + 1, Y + 2) = 16777215) Or _
                picChars(0).Visible = False Then
        
            bMove = True
        
            imgS.Top = imgS.Top - iSpeed
            MoveTopImg
            
            If (imgS.Top) Mod 8 = 0 Then
                picChars(0).Top = picChars(0).Top + 1
                picMap.Top = picMap.Top - 1
            End If
        Else: End If
        
    Case 4
        'Move Up
        If (GetPixel(picMap.hdc, X, Y - 1) = 16777215 And _
            GetPixel(picMap.hdc, X + 1, Y - 1) = 16777215) Or _
                picChars(0).Visible = False Then
           
            bMove = True
           
            imgS.Top = imgS.Top + iSpeed
            MoveTopImg

            If (imgS.Top) Mod 8 = 0 Then
                picChars(0).Top = picChars(0).Top - 1
                picMap.Top = picMap.Top + 1
            End If
        Else: End If
    
    End Select

    'Get New Variables
    X = picChars(0).Left: Y = picChars(0).Top
    
    PlaceExclemES

    If bMove = True Then
        'Show Anim
        Set imgChars(0).Picture = Chars(0).ShowWalk(lDir)
        
        'Check Squares
        If CheckEventSquares(X, Y) Then
            Exit Function
        End If
        
        'No EventSquares, then Check Fiends
        If bOverideControl = False And iFiendChance > 0 Then 'Check user actually has control
            If iWalkChance > iFreeSteps Then
                If RandomNumber(iFiendChance, iWalkChance) = iFiendChance Then
                    iWalkChance = 1
                    MoveChar = 1
                End If
            End If
            
            'Increase Fiend Chance
            iWalkChance = iWalkChance + 1
        End If
    End If
    
End Function

Public Function CheckDoors(X, Y, imgChar) As Boolean

    X = X / 2
    Y = Y / 2

    Dim I As Integer, sP() As String, sItem As String

    If X = 0 And Y = 0 Then
        Exit Function
    End If

    For I = 1 To cDoors.Count
        sItem = cDoors(I)
        
        sP = Split(sItem, ":")
        
        If sP(0) = X And sP(1) = Y - 1 Then
            'One Square Below
            
            staticProps(sP(2)).Visible = True
            imgChar.Visible = True
            
        ElseIf sP(0) = X And sP(1) = Y + 1.5 Then
            staticProps(sP(2)).Visible = False
            imgChar.Visible = True
            
        ElseIf sP(0) = X And sP(1) = Y Then
            'staticProps(sP(2)).Visible = False
            imgChar.Visible = False
            
        ElseIf sP(0) = X And sP(1) = Y + 1 Then
        
            staticProps(sP(2)).Visible = True
            imgChar.Visible = False
            
        ElseIf sP(0) = X And sP(1) = Y - 1.5 Then
            staticProps(sP(2)).Visible = False
            
        ElseIf sP(0) = X - 0.5 And sP(1) = Y - 1 Or _
            sP(0) = X + 0.5 And sP(1) = Y - 1 Then
        
            staticProps(sP(2)).Visible = False
        End If
    Next

End Function

Public Function CheckEventSquares(ByVal X As Double, ByVal Y As Double) As Boolean

Dim I As Integer, sP() As String, sItem As String

    If bIgnoreEvent = True Then
    
        bIgnoreEvent = False
        Exit Function
    End If

    X = X / 2
    Y = Y / 2

    For I = 1 To cEvents.Count
        sItem = cEvents.GetItem(I)
        
        sP = Split(sItem, ":")

        If sP(0) = X And sP(1) = Y Then
            If CBol(sP(3)) = True Then
                cEvents.Remove I
            End If
            
            bIgnoreEvent = True
            CheckEventSquares = True
            
            ExeStoryCmd sP(2)
            Exit Function
        End If
    Next
    
    CheckEventSquares = False

End Function

Public Function LookDir(sName As String, lDir As Integer)

    Dim iChar As Integer
    iChar = getCharByName(sName)

    If iChar = -1 Then
        WarnUser "Character not found: " & sName & " ."
        Exit Function
    End If

    Set imgChars(iChar).Picture = Chars(iChar).ShowWalk(lDir)

End Function

Public Function Sleep(iMili As Integer)

    On Error Resume Next
    
    HaltStory
    
    timSleep.Interval = iMili
    timSleep.Enabled = True
    
    If Err Then
        WarnUser Err.Description
    End If

End Function

Public Function HideChar(sName As String)

    Dim iChar As Integer
    iChar = getCharByName(sName)

    If iChar = -1 Then
        WarnUser "HideChar failed: " & sName & " does not exist"
        Exit Function
    End If

    'Hide from scen
    imgChars(iChar).Visible = False
    picChars(iChar).Visible = False

End Function

Public Function ShowChar(sName As String)

Dim iChar As Integer

    iChar = getCharByName(sName)

    If iChar = -1 Then
        WarnUser "ShowChar failed: " & sName & " does not exist"
        Exit Function
    End If
    
    'Show from scen
    imgChars(iChar).Visible = True
    picChars(iChar).Visible = True

End Function

Public Function PositionChar(sName As String, X As Integer, Y As Integer, lDir As Integer)

Dim iChar As Integer

    iChar = getCharByName(sName)
    
    If iChar = 0 Then
        ReposistionMainChar X, Y, lDir
        Exit Function
    End If
    
    'Only Utililise big square
    X = X * 2
    Y = Y * 2
    
    If iChar = -1 Then
        WarnUser "PositionChar failed: " & sName & " does not exist"
        Exit Function
    End If
    
    'Position on Map
    With picChars(iChar)
        .Top = (picMap.ScaleHeight - Y) - 2
        .Left = X
    End With
    
    'Position on Field
    With imgChars(iChar)
        .MaskColor = lTrans
    
        .Top = (imgS.ScaleHeight - (Y * 8)) - 16
        .Left = X * 8

        .Visible = False
        
        Set .Picture = Chars(.Index).ShowWalk(lDir)
        .Tag = ""
        
        Set .Container = imgS
    End With

End Function

Public Function SetPrologue(sPath As String, sPicture As String, lSpeed As Long)
    HaltStory
    
    pPrologue.Visible = True
    pPrologue.StartPrologue sPath, sPicture, lSpeed
End Function

Private Sub keyReciever1_Timer()
    'Use the Keys timer event to prevent conflict
    MoveProps cWalkListX, cWalkListY
End Sub

Private Sub pPrologue_OnFinished()
    pPrologue.Visible = False
    ResumeStory
End Sub

Private Sub timBounce_Timer()

Static iState As Integer, bDown As Boolean
    
    Debug.Print iState & ":" & bDown
    
    If bDown = True Then
        If iState > 0 Then
            imgChars(iBounceCharID).Top = imgChars(iBounceCharID).Top + 4
            iState = iState - 1
        Else
            bDown = False
            timBounce.Enabled = False
            
            ResumeStory
        End If
            
    ElseIf iState < 5 Then
        imgChars(iBounceCharID).Top = imgChars(iBounceCharID).Top - 4
        iState = iState + 1
    ElseIf iState = 5 Then
        bDown = True
    End If

End Sub

Private Sub timBounce2_Timer()

Static iState As Integer, bDown As Boolean
    
    Debug.Print iState & ":" & bDown

    If bDown = True Then
        If iState > 0 Then
            imgChars(iBounceCharID).Top = imgChars(iBounceCharID).Top + 4
            staticProps(iSystem_Exlem).Top = staticProps(iSystem_Exlem).Top + 4
            
            iState = iState - 1
        Else
            bDown = False
            timBounce2.Enabled = False
            
            staticProps(iSystem_Exlem).Visible = False
            
            ResumeStory
        End If
            
    ElseIf iState < 5 Then
        imgChars(iBounceCharID).Top = imgChars(iBounceCharID).Top - 4
        staticProps(iSystem_Exlem).Top = staticProps(iSystem_Exlem).Top - 4
        
        iState = iState + 1
    ElseIf iState = 5 Then
        bDown = True
    End If

End Sub

Private Sub timFrameProp_Timer(Index As Integer)
        
    With staticProps(timFrameProp(Index).Tag)
        If .StanceIndex = .FrameUbound And .FrameLoop = False Then
            timFrameProp(Index).Enabled = False
            Exit Sub
        End If
        
        .NextStance
    End With
    
End Sub

Private Function MoveTarget(Target As PictureBox, lDir As Integer)

    Dim X As Integer, Y As Integer
    X = Target.Left: Y = Target.Top

    Select Case lDir
        
    Case 2
        'Move Right
        Target.Left = Target.Left + 1
            
    Case 1
        'Move Left
        Target.Left = Target.Left - 1
            
    Case 3
        'Move Down
        Target.Top = Target.Top + 1
            
    Case 4
        'Move Up
        Target.Top = Target.Top - 1
        
    End Select

End Function

Private Function SafeToMove(Target As PictureBox, lDir As Integer) As Boolean

    Dim X As Integer, Y As Integer
    X = Target.Left: Y = Target.Top

    SafeToMove = False

        Select Case lDir
        
        Case 2
            'Move Right
            If GetPixel(picMap.hdc, X + 2, Y) = 16777215 And _
                GetPixel(picMap.hdc, X + 2, Y + 1) = 16777215 Then
                
                SafeToMove = True
            End If
            
        Case 1
            'Move Left
            If GetPixel(picMap.hdc, X - 1, Y) = 16777215 And _
               GetPixel(picMap.hdc, X - 1, Y + 1) = 16777215 Then
               
                SafeToMove = True
            End If
            
        Case 3
            'Move Down
            If GetPixel(picMap.hdc, X, Y + 2) = 16777215 And _
               GetPixel(picMap.hdc, X + 1, Y + 2) = 16777215 Then
            
                SafeToMove = True
            End If
            
        Case 4
            'Move Up
            If GetPixel(picMap.hdc, X, Y - 1) = 16777215 And _
               GetPixel(picMap.hdc, X + 1, Y - 1) = 16777215 Then
    
                SafeToMove = True
            End If
        
        End Select

End Function

Function MoveProps(colSrcX As SuperCollection, colSrcY As SuperCollection)

    'Debug.Print "MoveProps"
    'On Error GoTo CatchErr

Dim sP() As String, iCurrentX As Single, iCurrentY As Single, charIndex As Integer, iSpeed As Integer, I As Integer, _
    sItem, colIndex As Integer

    'Find Movable chars

    For I = 1 To colSrcX.Count
        sItem = colSrcX.GetItem(I, "X")
        
        If sItem = "X" Then
            'Out of bounds
            Exit For
        End If
    
        sP = Split(sItem, ":")
        iSpeed = CInt(sP(2))
        
        charIndex = getCharByName(sP(0))
        
        If charIndex = -1 Then
            WarnUser "Character not found: " & sP(0) & " ."
            Exit Function
        End If
        
        iCurrentX = picChars(charIndex).Left / 2

        If iCurrentX = sP(1) Then
            colSrcX.Remove sP(0)

            If CBol(sP(3)) = True Then
                ResumeStory
            End If
        Else
            If iCurrentX < sP(1) Then
                If imgChars(charIndex).Tag = "Self" Then
                    'Move Screen (instead)
                    MoveChar 2, True, iSpeed
                Else
                    If CheckPropCollision(charIndex, 2) = False Then
                        imgChars(charIndex).Left = imgChars(charIndex).Left + iSpeed
                        
                        If MapChar(charIndex) = True Then
                            Set imgChars(charIndex).Picture = Chars(charIndex).ShowWalk(2)
                        End If
                    Else
                        WarnUser "[" & sP(0) & "] Collides when trying to move right", False
                        LookDir sP(0), 2
                    End If
                End If
            Else
                If imgChars(charIndex).Tag = "Self" Then
                    'Move Screen (instead)
                    MoveChar 1, True, iSpeed
                Else
                    If CheckPropCollision(charIndex, 1) = False Then
                        'Move Character
                        imgChars(charIndex).Left = imgChars(charIndex).Left - iSpeed
                        
                        If MapChar(charIndex) = True Then
                            Set imgChars(charIndex).Picture = Chars(charIndex).ShowWalk(1)
                        End If
                    Else
                        WarnUser "[" & sP(0) & "] Collides when trying to move left", False
                        LookDir sP(0), 1
                    End If
                End If
            End If
        End If
    Next
    
    'Do same for WalkListY
    
    For I = 1 To colSrcY.Count
        sItem = colSrcY.GetItem(I, "X")
        If sItem = "X" Then
            'Out of bounds
            Exit For
        End If
    
        sP = Split(sItem, ":")
        charIndex = getCharByName(sP(0))
        
        If charIndex = -1 Then
            WarnUser "Character not found: " & sP(0) & " ."
            Exit Function
        End If
        
        iSpeed = CInt(sP(2))
        iCurrentY = picChars(charIndex).Top / 2
        
        If iCurrentY = sP(1) Then
        
            colSrcY.Remove sP(0)

            If CBol(sP(3)) = True Then
                ResumeStory
            End If
        Else
            If iCurrentY < sP(1) Then
                If imgChars(charIndex).Tag = "Self" Then
                    'Move Screen (instead)
                    MoveChar 3, True, iSpeed
                Else
                    If CheckPropCollision(charIndex, 3) = False Then
                
                        imgChars(charIndex).Top = imgChars(charIndex).Top + iSpeed
                            
                        If MapChar(charIndex) = True Then
                            Set imgChars(charIndex).Picture = Chars(charIndex).ShowWalk(3)
                        End If
                    Else
                        WarnUser "[" & sP(0) & "] Collides when trying to move down", False
                        LookDir sP(0), 3
                    End If
                End If
                    
            ElseIf iCurrentY > sP(1) Then
                If imgChars(charIndex).Tag = "Self" Then
                    'Move Screen (instead)
                    MoveChar 4, True, iSpeed
                Else
                    If CheckPropCollision(charIndex, 4) = False Then
                    
                        imgChars(charIndex).Top = imgChars(charIndex).Top - iSpeed
        
                        If MapChar(charIndex) = True Then
                            Set imgChars(charIndex).Picture = Chars(charIndex).ShowWalk(4)
                        End If
                    Else
                        WarnUser "[" & sP(0) & "] Collides when trying to move up", False
                        LookDir sP(0), 4
                    End If
                End If
            End If
        End If
    Next
    
    CheckDoors picChars(charIndex).Left, picChars(charIndex).Top, imgChars(charIndex)

    Exit Function
    
CatchErr:
    WarnUser "Sys_MoveProps has failed: " & Err.Number

End Function

Function getCharByName(sName As String) As Integer

Dim I As Integer
    
    If LCase(sName) = "self" Then
        getCharByName = 0
        Exit Function
    End If

    For I = 0 To imgChars.UBound
        If Chars(I).Tag = sName Then
            getCharByName = I
            Exit Function
        End If
    Next

    getCharByName = -1
    
End Function

Function getFramePropTimerByName(sName As String) As Integer

    Dim I As Integer
    
    For I = 0 To timFrameProp.UBound
        If timFrameProp(I).Tag = sName Then
            getFramePropTimerByName = I
            Exit Function
        End If
    Next
    getFramePropTimerByName = -1

End Function

Function GetPropByName(sName As String) As Integer

    Dim I As Integer
    
    For I = iProp_Ubound To staticProps.UBound
        If staticProps(I).Tag = sName Then
            GetPropByName = I
            Exit Function
        End If
    Next
    GetPropByName = -1
    
End Function

Function DivIgnoreDec(iNum, iNum2) As Integer

    Dim s As String
    s = iNum / iNum2
    
    If InStr(s, ".") Then
        DivIgnoreDec = Split(s, ".")(0)
    Else
        DivIgnoreDec = s
    End If

End Function

Private Sub timSleep_Timer()

    timSleep.Enabled = False
    ResumeStory

End Sub

Private Sub timTerm_Timer()

Dim I As Integer

    timTerm.Enabled = False

    'Reset Ignores
    bIgnoreEvent = False

    'Remove Fiend Chances
    SetFiendChance 0, 0

    UserControl.BackColor = vbBlack

    'Kill All Doors
    Echo "Unloading Doors"
    
    Set cDoors = New Collection
    
    'Kill all Unprotected Aliases
    Aliases.ClearList

    'Kill All Props
    Echo "Unloading Props"
    For I = iProp_Ubound To staticProps.UBound
        Unload staticProps(I)
    Next
    
    Echo "Unloading Frame Props"
    'Kill all timers
    For I = 1 To timFrameProp.UBound
        Unload timFrameProp(I)
    Next
    
    'Refresh Walk Commands
    Echo "Clearing Walk Commands"
    
    Set cWalkListX = New SuperCollection
    Set cWalkListY = New SuperCollection
    
    'Kill All Current Chars
    Echo "Unloading Characters"
    For I = 1 To Chars.UBound
        Unload Chars(I)
        Unload imgChars(I)
        Unload picChars(I)
    Next
    
    'Kill All Event Squares
    Echo "Unloading Event Squares"
    Set cEvents = New SuperCollection
    
    Echo "Unloading Async Events"
    
    Set cAsyncActions = New SuperCollection
    Set cCollisionEvents = New SuperCollection
    
    Set cEventsA = New SuperCollection
    Set cEvents = New SuperCollection
    
    Echo "Waiting for Start Position... "
    
End Sub

Private Sub Transition_OnFinishedGrow()

    timTerm_Timer
    ResumeStory

End Sub

Private Sub Transition_OnFinishedShrink()

    keyReciever1.Enabled = True
    ResumeStory

End Sub

Private Sub UserControl_Initialize()

    If NotInDbg = False Then
        Exit Sub
    End If

    'txtOut.Move 0, 0
    
    keyReciever1.ClearKeys
    
    keyReciever1.AddKey Control_Select
    keyReciever1.AddKey 27 'Escape
    
    keyReciever1.AddKey Control_Up
    keyReciever1.AddKey Control_Down
    keyReciever1.AddKey Control_Left
    keyReciever1.AddKey Control_Right
    
    keyReciever1.AddKey Control_Equip1
    keyReciever1.AddKey Control_Equip2
    keyReciever1.AddKey Control_Equip3
    keyReciever1.AddKey Control_Equip4
    
    imgChars(0).MaskColor = lTrans
    imgS_Top.MaskColor = lTrans
    
    Chars(0).LoadFrames sPath_Story & "characters\main.chr"
    'UserControl.Picture = imgS.Picture
    
    Set imgCur.Picture = frmLib.imgCur.Picture
    imgCur.MaskColor = lTrans
    
    iSystem_Exlem = StaticProp("", sPath_System & "!.bmp", 0, 40, True)
    
    keyReciever1.Enabled = False
    
End Sub

Function PlayMusic()
    
    PlayAudio sMusicCurrent, True
    
End Function

Function GetKeyAscii(lDir As Integer) As Integer

    Select Case lDir
    
    Case 1
        GetKeyAscii = Control_Left
    
    Case 2
        GetKeyAscii = Control_Right
        
    Case 3
        GetKeyAscii = Control_Down
        
    Case 4
        GetKeyAscii = Control_Up
        
    Case Else
        'WarnUser lDir
    
    End Select

End Function

Private Function CheckPropCollision(charIndex As Integer, lDir As Integer) As Boolean

    CheckPropCollision = False

Dim I As Integer, mChar As PictureBox

    Set mChar = picChars(charIndex)
    

    For I = 0 To picChars.UBound
    
        If picChars(I).Visible = True Then
    
            Select Case lDir
            
            Case CD_Right
                If (mChar.Left + 2 = picChars(I).Left) And (mChar.Top = picChars(I).Top) Or _
                    (mChar.Left + 2 = picChars(I).Left) And (mChar.Top = picChars(I).Top - 1) Or _
                      (mChar.Left + 2 = picChars(I).Left) And (mChar.Top = picChars(I).Top + 1) Then
                    
                    CheckPropCollision = True
                End If
                
            Case CD_Left
                If (mChar.Left - 2 = picChars(I).Left) And (mChar.Top = picChars(I).Top) Or _
                    (mChar.Left - 2 = picChars(I).Left) And (mChar.Top = picChars(I).Top - 1) Or _
                      (mChar.Left - 2 = picChars(I).Left) And (mChar.Top = picChars(I).Top + 1) Then
                    
                    CheckPropCollision = True
                End If
                
            Case CD_Up
                If (mChar.Top - 2 = picChars(I).Top) And (mChar.Left = picChars(I).Left) Or _
                    (mChar.Top - 2 = picChars(I).Top) And (mChar.Left = picChars(I).Left - 1) Or _
                      (mChar.Top - 2 = picChars(I).Top) And (mChar.Left = picChars(I).Left + 1) Then
                    
                    CheckPropCollision = True
                End If
            
            Case CD_Down
                If (mChar.Top + 2 = picChars(I).Top) And (mChar.Left = picChars(I).Left) Or _
                    (mChar.Top + 2 = picChars(I).Top) And (mChar.Left = picChars(I).Left - 1) Or _
                      (mChar.Top + 2 = picChars(I).Top) And (mChar.Left = picChars(I).Left + 1) Then
                    
                    CheckPropCollision = True
                End If
            
            End Select
            
            If CheckPropCollision = True Then
                If cCollisionEvents.Exists(Chars(charIndex).Tag) = True And I = 0 Then
                    ExeStoryCmd cCollisionEvents(Chars(charIndex).Tag)
                    cCollisionEvents.Remove Chars(charIndex).Tag
                End If
                
                Exit Function
            End If
            
        End If
        
    Next

End Function

Private Function PlaceExclemES()

    With staticProps(iSystem_Exlem)

        If CheckProximitiesES(False) = True Then
            .Visible = True
                            
            .Top = ((picChars(0).Top) * 8) - 16
            .Left = (picChars(0).Left) * 8
            
            .ZOrder 0
        Else
            .Visible = False
        End If
        
    End With

End Function

Private Function CheckProximitiesES(Optional bExecute As Boolean = True) As Boolean

    CheckProximitiesES = False

    Dim Item As String, sP() As String, I As Integer

    For I = 1 To cEventsA.Count
        Item = cEventsA(I)
        
        sP = Split(Item, ":")
    
        If sP(0) = picChars(0).Left And sP(1) = picChars(0).Top Then
            CheckProximitiesES = True
            
            If bExecute = True Then
                'RemoveCol cEventsA, sP(0) & ":" & sP(1)
                    
                cEvents.Remove I
                ExeStoryCmd sP(2)
            Else
                'Returning TRUE, Means ! Will Show
                CheckProximitiesES = CBol(sP(3))
            End If
            
            'Can only have one ActionSquare
            Exit Function
        End If
    Next

End Function

Private Function CheckProximitiesNPC() As Boolean

'Check if near char
Dim I As Integer, bTrue As Boolean, sKey As String

    CheckProximitiesNPC = False

    'Each Character has an Async Event, [just it might be empty]
    For I = 1 To cAsyncActions.Count
        sKey = Num2Col(I)
        
        'Check if char has action to activate, and that he/she is visible
        If cAsyncActions(sKey) <> "" And picChars(I).Visible = True Then
            
            bTrue = True
            
            'Face that way
            If (picChars(0).Left = picChars(I).Left And _
                (picChars(0).Top - 2 = picChars(I).Top Or picChars(0).Top + 2 = picChars(I).Top)) Or _
                (picChars(0).Top = picChars(I).Top And _
                (picChars(0).Left - 2 = picChars(I).Left Or picChars(0).Left + 2 = picChars(I).Left)) Then
                        
                    'Face that way
                    If picChars(0).Left + 2 = picChars(I).Left Then
                        Debug.Print "#1 " & picChars(0).Left
                        
                        Set imgChars(I).Picture = Chars(I).ShowWalk(1)
                    ElseIf picChars(0).Left - 2 = picChars(I).Left Then
                        Debug.Print "#2"
                        
                        Set imgChars(I).Picture = Chars(I).ShowWalk(2)
                    ElseIf picChars(0).Top - 2 = picChars(I).Top Then
                        Debug.Print "#3 " & picChars(0).Top
                        
                        Set imgChars(I).Picture = Chars(I).ShowWalk(3)
                    ElseIf picChars(0).Top + 2 = picChars(I).Top Then
                        Debug.Print "#4"
                        
                        Set imgChars(I).Picture = Chars(I).ShowWalk(4)
                    End If
                
            Else
                bTrue = False
            End If
                
            If bTrue Then
                DoAsyncEvents sKey
                CheckProximitiesNPC = True
                    
                Exit For
            End If

        End If
    Next

End Function

Private Function InProximity(Obj As Variant) As Boolean

    InProximity = True
    
    'Face that way
    If (picChars(0).Left = Obj.Left And _
        (picChars(0).Top - 2 = Obj.Top Or picChars(0).Top + 2 = Obj.Top)) Or _
        (picChars(0).Top = Obj.Top And _
        (picChars(0).Left - 2 = Obj.Left Or picChars(0).Left + 2 = Obj.Left)) Then
    Else
        InProximity = False
    End If

End Function

Private Sub ClearReport()
    'txtOut.Text = ""
End Sub

Private Sub Echo(sItem As String)

    'txtOut.Text = txtOut.Text & _
        sItem & _
        vbCrLf
        
    'txtOut.SelStart = Len(txtOut.Text)
        
    DoEvents

End Sub

'Pass KeyEvents on
Private Sub keyReciever1_OnKeyPressed(ByVal KeyAscii As Integer)
    OnKeyPressed KeyAscii
End Sub

Private Sub keyReciever1_OnKeyDown(ByVal KeyAscii As Integer)
    OnKeys KeyAscii
End Sub

Private Sub OnKeys(KeyCode As Integer)

Dim iDir As Integer
    
    If Player_Control = False And User_State <> State_Option Then
        Exit Sub
    End If

    iDir = GetDirection(KeyCode)
    If iDir = 0 Then
        Exit Sub
    End If
    
    If User_State = State_Option Then
        If iDir = 4 Then
            Options.GoUp
        ElseIf iDir = 3 Then
            Options.GoDown
        End If
    Else
        If MoveChar(iDir) <> 0 Then
            RaiseEvent OnFiend
        End If
    End If

End Sub

Private Sub OnKeyPressed(KeyCode As Integer)

    If bInBattle = True Then
        Exit Sub
    End If
    
    'Check for Equip
    If (Player_Control = True) And _
        (KeyCode >= Control_Equip1 And KeyCode <= Control_Equip4) Then
        Player_Control = False
        
        RaiseEvent OnEquip(KeyCode - 112)
            
        Exit Sub
    End If

    If Not KeyCode = Control_Select Then Exit Sub

    Select Case User_State
    
    Case State_Option
        User_State = State_NoWait
    
        imgCur.Visible = False
        Options.Visible = False
        Speech.Visible = False
    
        If Options.GetSelectedTag <> "" Then
            ExeStoryCmd Options.GetSelectedTag
        End If
        modStory.ResumeStory
    
    Case State_Notify
        Notify.Visible = False
        User_State = State_NoWait
    
        ResumeStory
    
    Case State_AsyncSpeech
        If Speech.CanScroll = True Then
            Speech.Scroll
        Else
            Speech.Visible = False
            User_State = State_NoWait
        End If
        
    Case State_Speech
        If Speech.CanScroll = True Then
            Speech.Scroll
        Else
            Speech.Visible = False
            User_State = State_NoWait
            
            'MsgBox "OnKeyPressed"
            If modStory.FunctionExists("&speech") Then
                ExeStoryCmd "&speech"
            Else
                modStory.ResumeStory
            End If
        End If
        
    Case Else
        If Player_Control = False Then
            'The things that follow are only applicable (if the player has control)
            Exit Sub
        End If
    
        If CheckProximitiesNPC = False Then
            CheckProximitiesES
        End If
    
    End Select

End Sub

Public Function CharX(ByRef Obj As Variant) As Long
    CharX = Obj.Left / 2
End Function

Public Function CharY(ByRef Obj As Variant) As Long
    CharY = ((picMap.ScaleHeight - Obj.Top) / 2) - 1
End Function

Private Function MapChar(Index As Integer) As Boolean

Dim iChar, oldLeft As Integer, oldTop As Integer

    Set iChar = imgChars(Index)
    
    oldLeft = picChars(Index).Left
    oldTop = picChars(Index).Top

    picChars(Index).Left = iChar.Left / 8
    picChars(Index).Top = iChar.Top / 8

    If oldLeft <> picChars(Index).Left Or oldTop <> picChars(Index).Top Then
        MapChar = True
    End If

End Function
