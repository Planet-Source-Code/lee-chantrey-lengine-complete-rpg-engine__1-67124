VERSION 5.00
Begin VB.UserControl usrNumero 
   BackColor       =   &H008000FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "usrNumero.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   9
      Left            =   1200
      Picture         =   "usrNumero.ctx":0102
      Top             =   2160
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   8
      Left            =   1080
      Picture         =   "usrNumero.ctx":0204
      Top             =   2160
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   7
      Left            =   1440
      Picture         =   "usrNumero.ctx":0306
      Top             =   2040
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   6
      Left            =   1320
      Picture         =   "usrNumero.ctx":0408
      Top             =   2040
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   5
      Left            =   1200
      Picture         =   "usrNumero.ctx":050A
      Top             =   2040
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   4
      Left            =   1080
      Picture         =   "usrNumero.ctx":060C
      Top             =   2040
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   3
      Left            =   1440
      Picture         =   "usrNumero.ctx":070E
      Top             =   1920
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   2
      Left            =   1320
      Picture         =   "usrNumero.ctx":0810
      Top             =   1920
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   1
      Left            =   1200
      Picture         =   "usrNumero.ctx":0912
      Top             =   1920
      Width           =   120
   End
   Begin VB.Image imgGNum 
      Height          =   120
      Index           =   0
      Left            =   1080
      Picture         =   "usrNumero.ctx":0A14
      Top             =   1920
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   9
      Left            =   1200
      Picture         =   "usrNumero.ctx":0B16
      Top             =   1440
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   8
      Left            =   1080
      Picture         =   "usrNumero.ctx":0C18
      Top             =   1440
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   7
      Left            =   1440
      Picture         =   "usrNumero.ctx":0D1A
      Top             =   1320
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   6
      Left            =   1320
      Picture         =   "usrNumero.ctx":0E1C
      Top             =   1320
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   5
      Left            =   1200
      Picture         =   "usrNumero.ctx":0F1E
      Top             =   1320
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   4
      Left            =   1080
      Picture         =   "usrNumero.ctx":1020
      Top             =   1320
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   2
      Left            =   1320
      Picture         =   "usrNumero.ctx":1122
      Top             =   1200
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   1
      Left            =   1200
      Picture         =   "usrNumero.ctx":1224
      Top             =   1200
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   0
      Left            =   1080
      Picture         =   "usrNumero.ctx":1326
      Top             =   1200
      Width           =   120
   End
   Begin VB.Image imgNum 
      Height          =   120
      Index           =   3
      Left            =   1440
      Picture         =   "usrNumero.ctx":1428
      Top             =   1200
      Width           =   120
   End
End
Attribute VB_Name = "usrNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function GetNumero(iNum As Integer, Optional bOptionalGood As Boolean = False) As IPictureDisp

    If bOptionalGood = False Then
        'Get Damage (White)
        Set GetNumero = imgNum(iNum).Picture
        
    Else
        'Get Good (Green)
        Set GetNumero = imgGNum(iNum).Picture
    End If

End Function

