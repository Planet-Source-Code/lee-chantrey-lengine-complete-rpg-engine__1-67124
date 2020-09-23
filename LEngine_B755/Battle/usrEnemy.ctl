VERSION 5.00
Begin VB.UserControl usrEnemy 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image imgA 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "usrEnemy.ctx":0000
      Top             =   2280
      Width           =   450
   End
   Begin VB.Image imgA 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "usrEnemy.ctx":04CD
      Top             =   2280
      Width           =   450
   End
End
Attribute VB_Name = "usrEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private bFloats As Boolean
Private sPath As String

Property Get Floats() As Boolean
    Floats = bFloats
End Property

Property Let Floats(b_Floats As Boolean)
    bFloats = b_Floats
End Property

Function ShowStance(lStance As Integer) As IPictureDisp
    Set ShowStance = imgA(lStance).Picture
End Function

Sub LoadPath(s_Path As String)
    
    'Snatch Path
    sPath = s_Path
    
    'Load Aval Images
    Dim I As Integer
    For I = 0 To imgA.UBound
        imgA(I).Picture = LoadPicture(s_Path & "\a" & I & ".gif")
    Next

End Sub
