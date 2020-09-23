VERSION 5.00
Begin VB.UserControl usrNotify 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "usrNotify.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Text 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "usrNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Initialize()
    'Make Pink Trans
    UserControl.MaskPicture = UserControl.Picture
    UserControl.MaskColor = lTrans
End Sub

Private Sub UserControl_Resize()
    'Center Text
    Text.Move Text.Left, 0, UserControl.ScaleWidth - (Text.Left * 2)
End Sub

Property Let Caption(newCaption As String)
    Text.Caption = newCaption
End Property

Property Get Caption() As String
    Caption = Text.Caption
End Property

