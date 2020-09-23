VERSION 5.00
Begin VB.UserControl usrEquip 
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LockControls    =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin prjLEngine.keyReciever KeyEvents 
      Left            =   4380
      Top             =   60
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin prjLEngine.usrTransPic imgAvatar 
      Height          =   480
      Left            =   420
      TabIndex        =   56
      Top             =   660
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   847
      MaskColor       =   16777215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   210
      TabIndex        =   57
      Tag             =   "C:0"
      Top             =   420
      Width           =   855
   End
   Begin VB.Label lblPrev 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   4320
      TabIndex        =   55
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblPrev 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   3780
      TabIndex        =   54
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image Image7 
      Height          =   120
      Left            =   2340
      Picture         =   "usrMenuScreen.ctx":0000
      Top             =   1380
      Width           =   120
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   4080
      TabIndex        =   53
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   4080
      TabIndex        =   52
      Top             =   3060
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   4080
      TabIndex        =   51
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   50
      Top             =   2700
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   49
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   48
      Top             =   2340
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   47
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   46
      Top             =   1980
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   45
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   44
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   43
      Top             =   3060
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   42
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   41
      Top             =   2700
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   40
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   39
      Top             =   2340
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   38
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   37
      Top             =   1980
      Width           =   495
   End
   Begin VB.Label Qty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   36
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblPrev 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   3420
      TabIndex        =   35
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblCurr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   2880
      TabIndex        =   34
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image Image6 
      Height          =   90
      Left            =   4140
      Picture         =   "usrMenuScreen.ctx":034C
      Top             =   1380
      Width           =   90
   End
   Begin VB.Label lblPrev 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   2520
      TabIndex        =   33
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblCurr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   1980
      TabIndex        =   32
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image Image5 
      Height          =   90
      Left            =   3240
      Picture         =   "usrMenuScreen.ctx":0693
      Top             =   1380
      Width           =   90
   End
   Begin VB.Label lblPrev 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   1620
      TabIndex        =   31
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblCurr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   1080
      TabIndex        =   30
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   120
      Left            =   1440
      Picture         =   "usrMenuScreen.ctx":09D9
      Top             =   1380
      Width           =   105
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2940
      TabIndex        =   29
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2940
      TabIndex        =   28
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2940
      TabIndex        =   27
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2940
      TabIndex        =   26
      Top             =   540
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2940
      TabIndex        =   25
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblCurr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   24
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblPrev 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   720
      TabIndex        =   23
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   120
      Left            =   540
      Picture         =   "usrMenuScreen.ctx":0D2F
      Top             =   1380
      Width           =   120
   End
   Begin VB.Image Cur 
      Height          =   240
      Left            =   60
      Picture         =   "usrMenuScreen.ctx":1080
      Tag             =   "0"
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Type 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   1200
      TabIndex        =   22
      Tag             =   "-1:10"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Type 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   1200
      TabIndex        =   21
      Tag             =   "C:1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Type 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   1200
      TabIndex        =   20
      Tag             =   "-1:10"
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Type 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   1200
      TabIndex        =   19
      Tag             =   "-1:10"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Type 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   1200
      TabIndex        =   18
      Tag             =   "-1:10"
      Top             =   900
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   1395
      Left            =   0
      Picture         =   "usrMenuScreen.ctx":1428
      Top             =   270
      Width           =   4800
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   2520
      TabIndex        =   17
      Tag             =   "-1:10"
      Top             =   3240
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   2520
      TabIndex        =   16
      Tag             =   "-1:10"
      Top             =   3060
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   2520
      TabIndex        =   15
      Tag             =   "-1:10"
      Top             =   2880
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   2520
      TabIndex        =   14
      Tag             =   "-1:10"
      Top             =   2700
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   2520
      TabIndex        =   13
      Tag             =   "-1:10"
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   2520
      TabIndex        =   12
      Tag             =   "-1:10"
      Top             =   2340
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   2520
      TabIndex        =   11
      Tag             =   "-1:10"
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   2520
      TabIndex        =   10
      Tag             =   "-1:10"
      Top             =   1980
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   2520
      TabIndex        =   9
      Tag             =   "-1:10"
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Tag             =   "-1:10"
      Top             =   3240
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Tag             =   "-1:10"
      Top             =   3060
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Tag             =   "-1:10"
      Top             =   2880
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Tag             =   "-1:10"
      Top             =   2700
      Width           =   1815
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Tag             =   "-1:10"
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Tag             =   "-1:10"
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2340
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1980
      Width           =   1500
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "usrMenuScreen.ctx":1712A
      Top             =   1680
      Width           =   4800
   End
End
Attribute VB_Name = "usrEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mVar_Name As String = "usrEquip"

'Aligns HotSpot with Super Collections
Private cPos As New Collection

Private Equip_Index(5) As Integer
Private HotSpot_Offset As Integer
Private TopMenu_Index As Integer
Private OldSel_Index As Integer
Private OldPrev_Index As Integer

Private CharID As String
Private BattleIndex As Integer

Private Const StrIndex As Integer = 0
Private Const DefIndex As Integer = 1
Private Const MagIndex As Integer = 2
Private Const SprIndex As Integer = 3
Private Const AtbIndex As Integer = 4

Public Event OnFinished()

Public Function RefreshStats()

    lblCurr(StrIndex) = pBattleHu(BattleIndex).Strength
    lblCurr(DefIndex) = pBattleHu(BattleIndex).Defence
    lblCurr(MagIndex) = pBattleHu(BattleIndex).Magic
    lblCurr(SprIndex) = pBattleHu(BattleIndex).Spirit

    ListSubOptions 0

End Function

Public Function ComitStats()

    pBattleHu(BattleIndex).Strength = lblCurr(StrIndex)
    pBattleHu(BattleIndex).Defence = lblCurr(DefIndex)
    pBattleHu(BattleIndex).Magic = lblCurr(MagIndex)
    pBattleHu(BattleIndex).Spirit = lblCurr(SprIndex)

    RaiseEvent OnFinished

End Function

Public Property Let Enabled(bNewE As Boolean)
    KeyEvents.Enabled = bNewE
    
    If bNewE = True Then
        'Were on show, lets refresh!
        ListSubOptions CInt(Split(Hotspot(Cur.Tag).Tag, ":")(1))
    End If
End Property

Public Property Get Char() As String

    Char = CharID

End Property

Public Property Let Char(sNew As String)

Dim Menu As New clsIniObj, I As Integer

    If sNew = "" Then
        Me.Enabled = False
        Exit Property
    End If
    
    CharID = sNew
    
    BattleIndex = FindBattleCharIndex(sNew)
    
    If BattleIndex = -1 Then
        OutErrMsg "Bad Battle Character Index", "usrEquip", "Char"
        Exit Property
    End If
    
    lblName.Caption = Aliases.GetItem(pBattleHu(BattleIndex).Name, pBattleHu(BattleIndex).Name)

    On Error GoTo Catch_E
    
    Menu.File = sPath_BattleChars & sNew & "\menu.ini"
    Menu.Section = "Equipment"
    
    For I = 1 To 5
        Menu.Key = "Type" & CStr(I)
        Hotspot(17 + I).Caption = Menu.Read
    Next
    
    Set imgAvatar.Picture = LoadPicture(sPath_BattleChars & sNew & "\avatar.bmp")
    
    Exit Property
    
Catch_E:
    NotifyUser "Equipment Load Failed: " & Err.Description
    
End Property

Private Sub ClearHotspots()

Dim I As Integer
    For I = 0 To 17
        Hotspot(I).Caption = ""
        Qty(I).Caption = ""
    Next

End Sub

Private Sub ClearPrev()

Dim I As Integer
    For I = 0 To lblPrev.UBound
        lblPrev(I).Caption = "0"
    Next

End Sub

Sub OverHotspot(bEmpty As Boolean)

Dim sIni As String, eIndex As Integer

    'On Error Resume Next
    If OnTopMenu = True Or cSubOptions(TopMenu_Index).Count = 0 Or bEmpty = True Then
        PrevEquip ""
    
        Exit Sub
    End If

    eIndex = cPos(Cur.Tag + HotSpot_Offset + 1)
    sIni = cSubOptions(TopMenu_Index).Key(eIndex)
    
    If OldSel_Index > 0 Then
        'Debug.Print ":P"
        UnPrevEquip cSubOptions(TopMenu_Index).Key(OldSel_Index)
    End If
        
    PrevEquip sIni

    OldSel_Index = eIndex

End Sub

Function EquipForMe(sUsers As String) As Boolean

Dim sP() As String, I As Integer
    sP = Split(sUsers, ",")
    
    For I = 0 To UBound(sP)
        If LCase(sP(I)) = CharID Then
            EquipForMe = True
            Exit Function
        End If
    Next
    
    EquipForMe = False

End Function

Function PrevEquip(sIni As String)

    Debug.Print "PrevEquip: " & sIni

Dim Data As New clsIniObj, iType As Integer
    Data.File = sIni

    iType = Data.Read("Type", "Structure", , 0)
    
    Data.Section = "Stats"
    
    If sIni = "" Then
        If OldPrev_Index > 0 Then
            UnPrevEquip cSubOptions(TopMenu_Index).Key(OldPrev_Index)
        
            OldSel_Index = 0
            OldPrev_Index = 0
        End If
        
        Exit Function
    End If

    lblPrev(StrIndex).Caption = lblPrev(StrIndex).Caption + CInt(lblCurr(StrIndex).Caption) + CInt(Data.Read("Strength", , , 0))
    lblPrev(DefIndex).Caption = lblPrev(DefIndex).Caption + CInt(lblCurr(DefIndex).Caption) + CInt(Data.Read("Defence", , , 0))
    lblPrev(MagIndex).Caption = lblPrev(MagIndex).Caption + CInt(lblCurr(MagIndex).Caption) + CInt(Data.Read("Magic", , , 0))
    lblPrev(SprIndex).Caption = lblPrev(SprIndex).Caption + CInt(lblCurr(SprIndex).Caption) + CInt(Data.Read("Spirit", , , 0))

    OldPrev_Index = cPos(Cur.Tag + HotSpot_Offset + 1)
    
End Function

Function UnPrevEquip(sIni As String, Optional bNew As Boolean = False)

    If sIni = "" Then
        Exit Function
    End If

    Debug.Print "UnPrevEquip: " & sIni

Dim Data As New clsIniObj, iType As Integer
    Data.File = sIni

    iType = Data.Read("Type", "Structure", , 0)
    
    Data.Section = "Stats"

    If bNew = True Then
        lblPrev(StrIndex).Caption = 0 - CInt(Data.Read("Strength", , , 0))
        lblPrev(DefIndex).Caption = 0 - CInt(Data.Read("Defence", , , 0))
        lblPrev(MagIndex).Caption = 0 - CInt(Data.Read("Magic", , , 0))
        lblPrev(SprIndex).Caption = 0 - CInt(Data.Read("Spirit", , , 0))
    Else
        lblPrev(StrIndex).Caption = CInt(lblPrev(StrIndex).Caption) - CInt(Data.Read("Strength", , , 0)) - CInt(lblCurr(StrIndex).Caption)
        lblPrev(DefIndex).Caption = CInt(lblPrev(DefIndex).Caption) - CInt(Data.Read("Defence", , , 0)) - CInt(lblCurr(DefIndex).Caption)
        lblPrev(MagIndex).Caption = CInt(lblPrev(MagIndex).Caption) - CInt(Data.Read("Magic", , , 0)) - CInt(lblCurr(MagIndex).Caption)
        lblPrev(SprIndex).Caption = CInt(lblPrev(SprIndex).Caption) - CInt(Data.Read("Spirit", , , 0)) - CInt(lblCurr(SprIndex).Caption)
    End If

End Function

Function Equip(sIni As String, bEmpty As Boolean)

    Debug.Print "Equip: " & sIni

Dim Data As New clsIniObj, sP() As String, iQty As Integer
    Data.File = sIni

    OldSel_Index = 0
    OldPrev_Index = 0
    
    If cSubOptions(TopMenu_Index).Count = 0 Then
        Exit Function
    End If

    UnEquipCurrent
    
    If bEmpty = True Then
        lblInfo(TopMenu_Index).Caption = ""
        Equip_Index(TopMenu_Index) = 0
        ListSubOptions TopMenu_Index

        ClearPrev
        Exit Function
    End If
    
    lblInfo(TopMenu_Index).Caption = Data.Read("Name", "Visual")
    
    Data.Section = "Stats"
    
    lblCurr(StrIndex).Caption = CInt(lblPrev(StrIndex).Caption)
    lblCurr(DefIndex).Caption = CInt(lblPrev(DefIndex).Caption)
    lblCurr(MagIndex).Caption = CInt(lblPrev(MagIndex).Caption)
    lblCurr(SprIndex).Caption = CInt(lblPrev(SprIndex).Caption)

    'MsgBox cPos(Cur.Tag + 1)
    Equip_Index(TopMenu_Index) = cPos(Cur.Tag + 1)
    'MsgBox cSubOptions(iType).GetItem(Equip_Index(TopMenu_Index))
    
    sP = Split(cSubOptions(TopMenu_Index).GetItem(Equip_Index(TopMenu_Index)), ":")
    iQty = CInt(sP(1)) - 1
    
    cSubOptions(TopMenu_Index).Item(Equip_Index(TopMenu_Index)) = sP(0) & ":" & CStr(iQty)

    ListSubOptions TopMenu_Index
    
    ClearPrev

End Function

Function UnEquipCurrent()

    Debug.Print "UnEquipCurrent: " & TopMenu_Index

Dim sP() As String, iQty As Integer

    If Equip_Index(TopMenu_Index) = 0 Then
        'Not Equiped
        Exit Function
    End If

    'MsgBox cSubOptions(TopMenu_Index).GetItem(Equip_Index(TopMenu_Index))

    sP = Split(cSubOptions(TopMenu_Index).GetItem(Equip_Index(TopMenu_Index)), ":")
    iQty = CInt(sP(1)) + 1
    
    cSubOptions(TopMenu_Index).Item(Equip_Index(TopMenu_Index)) = sP(0) & ":" & CStr(iQty)

    lblCurr(StrIndex).Caption = CInt(lblCurr(StrIndex).Caption) + CInt(lblPrev(StrIndex).Caption)
    lblCurr(DefIndex).Caption = CInt(lblCurr(DefIndex).Caption) + CInt(lblPrev(DefIndex).Caption)
    lblCurr(MagIndex).Caption = CInt(lblCurr(MagIndex).Caption) + CInt(lblPrev(MagIndex).Caption)
    lblCurr(SprIndex).Caption = CInt(lblCurr(SprIndex).Caption) + CInt(lblPrev(SprIndex).Caption)

End Function

Function AddEquip(ByRef sIni As String)

Dim Data As New clsIniObj, lQty As Integer, sP() As String, vType As Variant
    
    If Len(sIni) < 4 Then
        sIni = sIni & ".ini"
    Else
        If LCase(Right(sIni, 4)) <> ".ini" Then
            sIni = sIni & ".ini"
        End If
    End If
    
    Data.File = sIni
    
    If FileExist(sIni) = False Then
        WarnUser "The file does not exist.", True
        Exit Function
    End If
    
    vType = Data.Read("Type", "Structure", , 0)
    
    If IsNumeric(vType) = True Then
        If vType < 1 Or vType > 5 Then
            WarnUser "Structure - Type in Weapon INI must be greator then zero, but less then 5", True
            Exit Function
        End If
    Else
        WarnUser "Structure - Type in Weapon INI must be numerical", True
        Exit Function
    End If
    
    On Error GoTo Catch_E
    
    If cSubOptions(vType).Exists(sIni) = True Then
        'Item Exists, Increase Quantity
        sP = Split(cSubOptions(vType).GetItem(sIni), ":")
        lQty = CInt(sP(1)) + 1
        
        cSubOptions(vType).Item(sIni) = sP(0) & ":" & CStr(lQty)
    Else
        cSubOptions(vType).Add Data.Read("Name", "Visual") & ":1", sIni
    End If
    
    Exit Function
Catch_E:
    WarnUser "AddEquip Failed:: " & Err.Description & " {" & sIni & "}"
    
End Function

Function OnTopMenu() As Boolean

    If Cur.Tag < 18 Then
        OnTopMenu = False
    Else
        OnTopMenu = True
    End If
    
End Function

Private Sub OnKey(ByVal KeyAscii As Integer)

Dim sCmd As String, sP() As String, cIndex As Integer

    sCmd = Hotspot(Cur.Tag).Tag
    sP = Split(sCmd, ":")
    
    cIndex = Cur.Tag + HotSpot_Offset + 1

    Select Case KeyAscii
    
    Case Control_Cancel
        If OnTopMenu = False Then
            'ClearPrev
            'Align_Hotspot 17 + TopMenu_Index
            Exit Sub
        Else
            ComitStats
        End If
    
    Case Control_Select
        If (UBound(sP) = -1) Then
            Exit Sub
        ElseIf (cIndex > cSubOptions(TopMenu_Index).Count And OnTopMenu = False) Then
            Equip "", True
            Align_Hotspot 17 + TopMenu_Index
            Exit Sub
        End If
        
        If sP(0) = "C" Then
            'ListSubOptions CInt(sP(1))
            Align_Hotspot 0
            
            'Preview UnEquip Current
            If Equip_Index(TopMenu_Index) > 0 Then
                UnPrevEquip cSubOptions(TopMenu_Index).Key(Equip_Index(TopMenu_Index)), True
            End If
        Else
            If cIndex <= cPos.Count Then
                Equip cSubOptions(TopMenu_Index).Key(cPos(cIndex)), IIf(Hotspot(Cur.Tag).Caption = "", True, False)
            Else
                Equip cSubOptions(TopMenu_Index).Key(cIndex), IIf(Hotspot(Cur.Tag).Caption = "", True, False)
            End If
            
            Align_Hotspot 17 + TopMenu_Index
        End If
    
    Case Control_Right
        If OnTopMenu = False Then
            Align_Hotspot CInt(sP(1))
        End If
    
    Case Control_Left
        If OnTopMenu = False Then
            Align_Hotspot CInt(sP(0))
        End If
    
    Case Control_Up
        If Cur.Tag = 0 And HotSpot_Offset > 0 Then
            ListSubOptions TopMenu_Index, HotSpot_Offset - 18
            Align_Hotspot 17
        Else
            If Cur.Tag = 18 Then
                'Do Nothing
            Else
                Align_Hotspot Cur.Tag - 1
            
                If OnTopMenu = True Then
                    'ClearHotspots
                    ListSubOptions CInt(Split(Hotspot(Cur.Tag).Tag, ":")(1))
                End If
            End If
        End If
    
    Case Control_Down
        If Cur.Tag = 17 Then
            ListSubOptions TopMenu_Index, HotSpot_Offset + 18
        Else
            Align_Hotspot Cur.Tag + 1
        
            If OnTopMenu = True Then
                'ClearHotspots
                ListSubOptions CInt(Split(Hotspot(Cur.Tag).Tag, ":")(1))
            End If
        End If
        
    Case Else
        Debug.Print KeyAscii
    
    End Select
    
    OverHotspot IIf(Hotspot(Cur.Tag).Caption = "", True, False)

End Sub

Sub ListSubOptions(cIndex As Integer, Optional iOffset As Integer = 0)

Dim I As Integer, sP() As String, Data As New clsIniObj, iQty As Integer, sDesc As String, _
    cItemIndex As Integer, colIndex As Integer
    
    HotSpot_Offset = iOffset
    TopMenu_Index = cIndex

    iOffset = iOffset + 1
    
    Set cPos = New Collection

    For I = 0 To 17
        Hotspot(I).Caption = ""
        Qty(I).Caption = ""
    
        If (colIndex + iOffset) <= cSubOptions(cIndex).Count Then
            Data.File = cSubOptions(cIndex).Key(colIndex + iOffset)
            Data.Section = "Structure"

            sP = Split(cSubOptions(cIndex).GetItem(colIndex + iOffset), ":")
            
            iQty = CInt(sP(1))
            sDesc = CStr(sP(0))
            
            If EquipForMe(Data.Read("Users")) = True And iQty > 0 Then
                If UBound(sP) = 1 Then
                    cPos.Add colIndex + iOffset
                
                    Hotspot(I).Caption = sDesc
                    Qty(I).Caption = iQty
                End If
            Else
                'Re-Use HotSpot
                I = I - 1
            End If

            colIndex = colIndex + 1
        End If
        
    Next

    'Align_Hotspot 0

End Sub

Private Sub KeyEvents_OnKeyPressed(ByVal KeyAscii As Integer)
    OnKey KeyAscii
End Sub

Private Sub UserControl_Initialize()

Dim I As Integer
    
    For I = 0 To 8
        Hotspot(I).Tag = "-1:" & CStr(9 + I)
    Next
    For I = 9 To 17
        Hotspot(I).Tag = CStr(I - 9) & ":-1"
    Next
    For I = 18 To 22
        Hotspot(I).Tag = "C:" & I - 17
    Next

    Align_Hotspot 18
    
    'Catch Key Events
    InitialiseKeys
    ClearHotspots

End Sub

Private Sub InitialiseKeys()

    'Notified on these key changes
    KeyEvents.ClearKeys
    
    KeyEvents.AddKey Control_Cancel
    KeyEvents.AddKey Control_Select
    
    KeyEvents.AddKey Control_Down
    KeyEvents.AddKey Control_Up
    KeyEvents.AddKey Control_Right
    KeyEvents.AddKey Control_Left

    KeyEvents.Enabled = False

End Sub

Sub Align_Hotspot(hIndex As Integer)
    
    If hIndex < 0 Or hIndex > Hotspot.UBound Then
        Exit Sub
    End If
    
    With Cur
        .Top = Hotspot(hIndex).Top + 4
        .Left = Hotspot(hIndex).Left - 16
        
        .Tag = hIndex
    End With
    
End Sub
