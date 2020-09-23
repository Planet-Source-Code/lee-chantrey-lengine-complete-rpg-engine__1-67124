VERSION 5.00
Begin VB.UserControl usrChar 
   BackColor       =   &H0080C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Image imgA 
      Height          =   240
      Index           =   4
      Left            =   2160
      Picture         =   "usrChar.ctx":0000
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgB 
      Height          =   240
      Index           =   4
      Left            =   2160
      Picture         =   "usrChar.ctx":03BA
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image imgB 
      Height          =   240
      Index           =   3
      Left            =   1800
      Picture         =   "usrChar.ctx":0772
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image imgA 
      Height          =   240
      Index           =   3
      Left            =   1800
      Picture         =   "usrChar.ctx":0B2E
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image imgB 
      Height          =   240
      Index           =   1
      Left            =   1080
      Picture         =   "usrChar.ctx":0EE6
      Top             =   2280
      Width           =   210
   End
   Begin VB.Image imgB 
      Height          =   240
      Index           =   2
      Left            =   1440
      Picture         =   "usrChar.ctx":129B
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image imgA 
      Height          =   240
      Index           =   2
      Left            =   1440
      Picture         =   "usrChar.ctx":164E
      Top             =   1920
      Width           =   195
   End
   Begin VB.Image imgS 
      Height          =   240
      Left            =   1080
      Picture         =   "usrChar.ctx":19F9
      Top             =   1560
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgA 
      Height          =   240
      Index           =   1
      Left            =   1080
      Picture         =   "usrChar.ctx":1CFB
      Top             =   1920
      Width           =   195
   End
End
Attribute VB_Name = "usrChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const iLeft As Integer = 1
Private Const iRight As Integer = 3

Private lastDir As Long
Private s_srPath As String

Public Property Get srPath() As String
    srPath = s_srPath
End Property

Public Property Let srPath(sNewPath As String)
    s_srPath = sNewPath
End Property

Function LastDirection() As Long

    LastDirection = lastDir

End Function

Function LoadFrames(sPath As String)

    If NotInDbg = False Then
        Exit Function
    End If

    On Error Resume Next

    Dim I As Integer
    For I = 1 To 4
        imgA(I).Picture = LoadPicture(sPath & "\a" & (I - 1) & ".bmp")
        imgB(I).Picture = LoadPicture(sPath & "\b" & (I - 1) & ".bmp")
    Next
    
    If Err Then
        If Err.Number = 53 Then
            MsgBox "LoadChar(): Pictures could not be found. Check Path and try again.", vbCritical
        End If
    End If

End Function

Function ShowStance(lDir As Integer) As IPictureDisp

    Set ShowStance = imgA(lDir).Picture

End Function

Function ShowWalk(lDir As Integer) As IPictureDisp

Static pDir As Integer

    If pDir = lDir Then
        Set ShowWalk = imgA(lDir).Picture
        pDir = 0
    Else
        Set ShowWalk = imgB(lDir).Picture
        pDir = lDir
    End If

    lastDir = lDir
    
End Function

Private Sub UserControl_Initialize()

    lastDir = 1

End Sub
