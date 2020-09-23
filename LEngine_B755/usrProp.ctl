VERSION 5.00
Begin VB.UserControl usrProp 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image imgA 
      Height          =   195
      Index           =   1
      Left            =   840
      Picture         =   "usrProp.ctx":0000
      Top             =   1680
      Width           =   225
   End
   Begin VB.Image imgA 
      Height          =   195
      Index           =   0
      Left            =   480
      Picture         =   "usrProp.ctx":02B2
      Top             =   1680
      Width           =   225
   End
End
Attribute VB_Name = "usrProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function Animate() As IPictureDisp

    'Previous Dir
    Static iIndex As Integer
    iIndex = iIndex + 1
    
    If iIndex = imgA.Count Then
        iIndex = 0
    End If
        
    Set Animate = imgA(iIndex).Picture

End Function

