VERSION 5.00
Begin VB.UserControl usrBattleChar 
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "usrBattleChar.ctx":0000
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Image imgS 
      Height          =   615
      Index           =   0
      Left            =   720
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "usrBattleChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function ShowStance(lStance As Integer) As IPictureDisp
    Set ShowStance = imgS(lStance).Picture
    UserControl.Picture = imgS(lStance).Picture
End Function

Function LoadStore(sPath As String, Optional iMax As Integer)

Dim I As Integer
    'Unload existing
    For I = 1 To imgS.UBound
        Unload imgS
    Next
    If iMax = 0 Then
        iMax = imgS.UBound
    End If
    
    imgS(0).Picture = LoadPicture(sPath & "\a0.bmp")
    
    For I = 1 To iMax
        Load imgS(imgS.Count)
        With imgS(imgS.UBound)
            .Picture = LoadPicture(sPath & "\a" & I & ".bmp")
        End With
    Next

End Function

Private Sub UserControl_Initialize()
    UserControl.MaskColor = RGB(255, 0, 128)
End Sub
