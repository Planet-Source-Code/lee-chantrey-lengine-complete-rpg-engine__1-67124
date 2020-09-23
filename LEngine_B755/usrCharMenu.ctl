VERSION 5.00
Begin VB.UserControl usrCharMenu 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "usrCharMenu.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   3360
      Picture         =   "usrCharMenu.ctx":038A
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   0
      Width           =   1440
      Begin VB.Timer timProg 
         Enabled         =   0   'False
         Index           =   3
         Left            =   360
         Top             =   1560
      End
      Begin VB.Timer timProg 
         Enabled         =   0   'False
         Index           =   0
         Left            =   360
         Top             =   720
      End
      Begin VB.Timer timProg 
         Enabled         =   0   'False
         Index           =   2
         Left            =   360
         Top             =   1440
      End
      Begin VB.Timer timProg 
         Enabled         =   0   'False
         Index           =   1
         Left            =   360
         Top             =   1080
      End
      Begin VB.Label Mp 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
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
         Height          =   225
         Index           =   3
         Left            =   420
         TabIndex        =   12
         Top             =   450
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Mp 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
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
         Height          =   225
         Index           =   2
         Left            =   420
         TabIndex        =   11
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Mp 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
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
         Height          =   225
         Index           =   1
         Left            =   420
         TabIndex        =   10
         Top             =   150
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Mp 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
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
         Height          =   225
         Index           =   0
         Left            =   420
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Hp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
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
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   8
         Top             =   450
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Hp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
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
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Hp 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
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
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   150
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Hp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
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
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image imgLoading 
         Height          =   45
         Left            =   0
         Picture         =   "usrCharMenu.ctx":39CC
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgComplete 
         Height          =   75
         Left            =   360
         Picture         =   "usrCharMenu.ctx":3A32
         Top             =   3120
         Width           =   60
      End
      Begin VB.Image imgP 
         Height          =   45
         Index           =   3
         Left            =   780
         Picture         =   "usrCharMenu.ctx":3D6E
         Stretch         =   -1  'True
         Top             =   555
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgPro 
         Height          =   105
         Index           =   3
         Left            =   720
         Picture         =   "usrCharMenu.ctx":3DD4
         Top             =   525
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgP 
         Height          =   45
         Index           =   2
         Left            =   780
         Picture         =   "usrCharMenu.ctx":4141
         Stretch         =   -1  'True
         Top             =   405
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgPro 
         Height          =   105
         Index           =   2
         Left            =   720
         Picture         =   "usrCharMenu.ctx":41A7
         Top             =   375
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgP 
         Height          =   45
         Index           =   1
         Left            =   780
         Picture         =   "usrCharMenu.ctx":4514
         Stretch         =   -1  'True
         Top             =   255
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgPro 
         Height          =   105
         Index           =   1
         Left            =   720
         Picture         =   "usrCharMenu.ctx":457A
         Top             =   225
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgP 
         Height          =   45
         Index           =   0
         Left            =   780
         Picture         =   "usrCharMenu.ctx":48E7
         Stretch         =   -1  'True
         Top             =   105
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgPro 
         Height          =   105
         Index           =   0
         Left            =   720
         Picture         =   "usrCharMenu.ctx":494D
         Top             =   75
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.Label Items 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing"
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
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Items 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing"
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
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   300
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Items 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing"
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
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Items 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing"
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
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image imgCur 
      Height          =   240
      Left            =   0
      Picture         =   "usrCharMenu.ctx":4CBA
      Top             =   30
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "usrCharMenu.ctx":5030
      Top             =   0
      Width           =   60
   End
   Begin VB.Image imgback 
      Height          =   720
      Left            =   0
      Picture         =   "usrCharMenu.ctx":52B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "usrCharMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iAtm(3) As Integer
Private ioScale As Integer
Private outCur As Image

'Cursor on me ?
Private bMe As Boolean
Private hIndex As Integer
Private lCount As Integer

Private oIndex As Integer
Private oiScale As Integer

Private bPaused As Boolean

Private iOldInterval(3) As Integer
Private bEnabled(3) As Boolean
Private bDead(3) As Boolean

Public Event OnReady(iIndex As Integer)

Public Function UpdateATB(Index As Integer, iAtb As Integer)
    timProg(Index).Interval = iAtb
End Function

Public Function UpdateMP(Index As Integer, iMpo As Integer)
    Mp(Index).Caption = iMpo
End Function

Public Function UpdateHP(Index As Integer, iHp As Integer)
    Hp(Index).Caption = iHp
End Function

Public Function HideCursor()
    If IsSomething(outCur) = False Then
        Exit Function
    End If
    
    imgCur.Visible = False
    outCur.Visible = False
End Function

Property Get SelectedCaption() As String
    SelectedCaption = Items(oIndex)
End Property

Property Get ListIndex() As Integer
    ListIndex = oIndex
End Property

Property Get HighlightedIndex() As Integer
    HighlightedIndex = hIndex
End Property

Function ClearList()

    'Reset everything

    bMe = False
    hIndex = 0
    lCount = 0
    ioScale = 0
    
    Dim I As Integer
    For I = 0 To 3
        iAtm(I) = 0
        iOldInterval(I) = 0
        
        timProg(I).Interval = 0
        timProg(I).Enabled = False
        
        imgP(I).Width = 1
        Items(I).Caption = "Nothing"
        
        imgP(I).Picture = imgLoading.Picture
        
        bEnabled(I) = False
        bDead(I) = False
        
        bPaused = False
        
        Set outCur = Nothing
    Next

End Function

Function ListCount() As Integer
    ListCount = lCount
End Function

Function ListUbound() As Integer
    ListUbound = lCount - 1
End Function

Function AliveCount() As Integer

Dim I As Integer
    For I = 0 To 3
        If bDead(I) = False Then
            AliveCount = AliveCount + 1
        End If
    Next

End Function

Sub Finished()
    bMe = False
End Sub

Sub ResetATM(Index As Integer)
    
    Set imgP(Index).Picture = imgLoading.Picture
    
    imgP(Index).Width = 1
    timProg(Index).Enabled = True
    
End Sub

Sub PauseATM()

    bPaused = True
    
    Dim I As Integer
    For I = 0 To timProg.UBound
        If timProg(I).Enabled = True Then
            timProg(I).Enabled = False
            bEnabled(I) = True
        Else
            bEnabled(I) = False
        End If
    Next
    
End Sub

Sub CharDied(Index As Integer)

    imgP(Index).Width = 1
    imgP(Index).Picture = imgLoading.Picture
    bDead(Index) = True

    Items(Index).ForeColor = vbRed
    timProg(Index).Enabled = False

End Sub

Sub CharRevived(ByVal Index As Integer)

    imgP(Index).Width = 1
    Items(Index).ForeColor = vbWhite
    bDead(Index) = False

    timProg(Index).Enabled = True

End Sub

Sub Highlight(Index As Integer)

    'Highlighting is different from selecting
    'thus listindex will return Selected

    Dim I As Integer
    For I = 0 To Items.UBound
        If Items(I).ForeColor = vbYellow Then
            Items(I).ForeColor = vbWhite
        End If
    Next
    
    Items(Index).ForeColor = vbYellow
    hIndex = Index

End Sub

Sub Nullight(Index As Integer)

    Items(Index).ForeColor = vbWhite

End Sub

Sub AttachCur(ByRef Cur As Image, myLeft As Integer, myTop As Integer, Optional iScale As Integer = 15)
    Set outCur = Cur
    
    imgCur.Visible = True
    outCur.Left = myLeft - (outCur.Width - (imgCur.Width * iScale))
    outCur.Top = myTop + (imgCur.Top * iScale)
    outCur.Visible = True
    
    oiScale = iScale
End Sub

Sub GoDown()

    If (oIndex + 1) > Items.UBound Then
        Exit Sub
    End If
    
    imgCur.Top = imgCur.Top + 10
    outCur.Top = outCur.Top + 10 * oiScale
    
    oIndex = oIndex + 1

End Sub

Sub GoUp()
    
    If (oIndex - 1) < 0 Then
        Exit Sub
    End If
 
    imgCur.Top = imgCur.Top - 10
    outCur.Top = outCur.Top - 10 * oiScale
    
    oIndex = oIndex - 1

End Sub

Sub AddCharacter(sName As String, iIndex As Integer, iHp As Integer, iMp2 As Integer, iAtm As Integer)

    lCount = lCount + 1

    Items(iIndex).Caption = sName
    
    Items(iIndex).Visible = True
    Mp(iIndex).Visible = True
    
    timProg(iIndex).Interval = iAtm
    iOldInterval(iIndex) = iAtm
    
    imgPro(iIndex).Visible = True
    imgP(iIndex).Visible = True

    imgP(iIndex).Width = 0
    
    Hp(iIndex).Caption = iHp
    Mp(iIndex).Caption = iMp2
    
    Hp(iIndex).Visible = True
    
    bDead(iIndex) = IIf(pBattleHu(iIndex).Alive = True, False, True)

End Sub

Sub StartATM()

    bPaused = False

    Dim I As Integer
    For I = 0 To timProg.UBound
        If bEnabled(I) = True And bDead(I) = False Then
            timProg(I).Enabled = True
        End If
    Next
    
End Sub

Sub ResumeATM()

    bPaused = False

    Dim I As Integer
    For I = 0 To timProg.UBound
        If bEnabled(I) = True And bDead(I) = False Then
            timProg(I).Enabled = True
        End If
    Next

End Sub

Sub RemoveCharacter(Index As Integer)

    Dim lUbound As Integer
    lUbound = lCount - 1

    'Swap with end
    If Index < lUbound Then
        Items(Index).Caption = Items(lUbound).Caption
        Items(Index).Tag = Items(lUbound).Tag
        
        Hp(Index).Caption = Hp(lUbound).Caption
        
        iOldInterval(Index) = iOldInterval(lUbound)
        bDead(Index) = bDead(lUbound)
        
        bEnabled(Index) = bEnabled(lUbound)
        
        timProg(Index).Enabled = timProg(lUbound).Enabled
        timProg(Index).Interval = timProg(lUbound).Interval
    End If
        
    bEnabled(lUbound) = False
    bDead(lUbound) = True
    
    iOldInterval(lUbound) = 0
    timProg(lUbound).Enabled = False
    timProg(lUbound).Interval = 0
        
    Items(lUbound).Caption = "Nothing"
    Items(lUbound).Tag = ""
    
    Items(lUbound).Visible = False
    Mp(lUbound).Visible = False
    
    Hp(lUbound).Visible = False
    imgPro(lUbound).Visible = False
    imgP(lUbound).Visible = False
    
    imgP(lUbound).Width = 4
    timProg(lUbound).Enabled = False
    
    'Minus 1
    lCount = lUbound

End Sub

Private Sub timProg_Timer(Index As Integer)

    If bPaused = False Then
        imgP(Index).Width = imgP(Index).Width + 2
    Else
        Exit Sub
    End If
    
    If imgP(Index).Width = 33 Then
        imgP(Index).Width = 31
    
        timProg(Index).Enabled = False
        Set imgP(Index).Picture = imgComplete.Picture

        RaiseEvent OnReady(Index)
    End If
    
End Sub

Private Sub UserControl_Initialize()

    'Empty space, is dead
Dim I As Integer
    For I = 0 To 3
        bDead(I) = True
    Next

End Sub

Private Sub UserControl_Resize()
    imgBack.Width = UserControl.Width
End Sub
