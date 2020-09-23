VERSION 5.00
Begin VB.UserControl usrName 
   BackColor       =   &H00808080&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "usrName.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin prjLEngine.keyReciever Keyboard 
      Left            =   3840
      Top             =   120
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   3000
      Width           =   2535
   End
End
Attribute VB_Name = "usrName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event OnConfirmed(ByVal sName As String)

Public Function SetPicture(ByVal sNewPicture As String)
    Set UserControl.Picture = LoadPicture(sNewPicture)
End Function

Public Property Get Keys() As keyReciever
    Set Keys = Keyboard
End Property

Public Property Let CharName(sNewName As String)
    Text.Caption = sNewName & "_"
End Property

Public Property Get CharName() As String
    CharName = Left(Text.Caption, Len(Text.Caption) - 1)
End Property

Private Sub Keyboard_OnKeyPressed(ByVal KeyAscii As Integer)

    Debug.Print KeyAscii

    If KeyAscii = 8 Then
        If Text.Caption <> "_" Then
            Text.Caption = Left(Text.Caption, Len(Text.Caption) - 2) & "_"
        End If
    ElseIf KeyAscii = Control_Select Then
        RaiseEvent OnConfirmed(Me.CharName)
        Text.Caption = "_"
        
        Keyboard.Enabled = False
    Else
        If GetAsyncKeyState(16) Then
            Text.Caption = Left(Text.Caption, Len(Text.Caption) - 1) & UCase(Chr(KeyAscii)) & "_"
        Else
            Text.Caption = Left(Text.Caption, Len(Text.Caption) - 1) & LCase(Chr(KeyAscii)) & "_"
        End If
    End If
    
    'Text.Caption = StrConv(Text.Caption, vbProperCase)

End Sub

Private Sub UserControl_Initialize()
    UserControl.BackColor = vbBlack
    
    Keyboard.SetSize 29
    
Dim I As Integer
    
    Keyboard.AddKey 27
    Keyboard.AddKey Control_Select
    Keyboard.AddKey 8
    
    For I = 65 To 90
        Keyboard.AddKey I
    Next
    
    Keyboard.Enabled = False
End Sub
