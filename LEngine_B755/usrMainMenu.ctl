VERSION 5.00
Begin VB.UserControl usrMainMenu 
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin prjLEngine.keyReciever KeyEvents 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin VB.Image Cur 
      Height          =   240
      Left            =   0
      Picture         =   "usrMainMenu.ctx":0000
      Tag             =   "0"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Load"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   3780
      TabIndex        =   4
      Tag             =   "-1:10"
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Save"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   3780
      TabIndex        =   3
      Tag             =   "-1:10"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   3780
      TabIndex        =   2
      Tag             =   "-1:10"
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Equipment"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   3780
      TabIndex        =   1
      Tag             =   "C:0"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Hotspot 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   3780
      TabIndex        =   0
      Tag             =   "-1:10"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   3660
      Picture         =   "usrMainMenu.ctx":03A8
      Top             =   0
      Width           =   1140
   End
End
Attribute VB_Name = "usrMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub KeyEvents_OnKeyPressed(ByVal KeyAscii As Integer)
    OnKey KeyAscii
End Sub

Private Sub OnKey(ByVal KeyAscii As Integer)

    Select Case KeyAscii

    Case Control_Up
        If Cur.Tag = 0 Then
            Align_Hotspot Hotspot.UBound
        Else
            Align_Hotspot Cur.Tag - 1
        End If
    
    Case Control_Down
        If Cur.Tag = Hotspot.UBound Then
            Align_Hotspot 0
        Else
            Align_Hotspot Cur.Tag + 1
        End If
        
    End Select

End Sub

Private Sub UserControl_Initialize()

    SetJoypad

    'Catch Key Events
    InitialiseKeys

    Align_Hotspot 0

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
    
    KeyEvents.Enabled = True

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
