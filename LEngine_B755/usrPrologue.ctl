VERSION 5.00
Begin VB.UserControl usrPrologue 
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "usrPrologue.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Timer timUpdater 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   120
      Top             =   3000
   End
End
Attribute VB_Name = "usrPrologue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Put This code in a module
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Const EM_GETLINECOUNT = &HBA

Private Declare Function DrawText Lib "user32" _
    Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Event OnFinished()

Private vTop         As Long     'Stores the Text Top pos
Private TxtSource      As String   'Stores the text lines
Private TxtHeight As Long

Public Function StartPrologue(sPath As String, Optional sPicture As String = "", Optional lSpeed As Long = 70)

    On Error GoTo Catch_E

Dim NoOfLines As Long

    vTop = UserControl.ScaleHeight

'  Loading the data from txtfile
    TxtSource = Fload(sPath)
    
    If TxtSource = "" Then
        WarnUser "StartPrologue:: No Text to Prologue", True
        Exit Function
    End If
    
    Text1.Text = TxtSource
    
    NoOfLines = SendMessage(Text1.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&)
    
    TxtHeight = UserControl.TextHeight("A") * NoOfLines

    If sPicture <> "" Then
        UserControl.Picture = LoadPicture(sPicture)
    End If

    timUpdater.Interval = lSpeed
    timUpdater.Enabled = True
    
    Exit Function
Catch_E:
    WarnUser "StartPrologue::" & Err.Description
    
End Function

'Returns the RGB values
Private Sub GetRGB(ByVal LngCol As Long, R As Long, G As Long, B As Long)
  R = LngCol Mod 256
  G = (LngCol And vbGreen) / 256 'Green
  B = (LngCol And vbBlue) / 65536 'Blue
End Sub

'Drawing the Text
Public Function SendPrologue(Txt As String, _
                        ByVal x As Integer, ByVal Y As Integer)
                        
Dim hLength   As Integer 'Region over which the text fades
Dim DrawCol   As Long    'The current faded color
Dim rctDraw   As RECT

    With rctDraw
        .Left = x
        .Top = Y
        .Right = UserControl.Width / 15
        .Bottom = UserControl.Height / 15
    End With

    DrawText UserControl.hdc, Txt, -1, rctDraw, &H10
End Function

Private Sub timUpdater_Timer()

Dim x As Integer, lHeight As Long
    UserControl.Cls
    
    SendPrologue TxtSource, 2, vTop

    If -vTop = TxtHeight Then
        timUpdater.Enabled = False
        RaiseEvent OnFinished
    End If
    
    vTop = vTop - 1

End Sub

