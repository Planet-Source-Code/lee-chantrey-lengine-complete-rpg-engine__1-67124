VERSION 5.00
Begin VB.UserControl usrSpeech 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Picture         =   "usrSpeech.ctx":0000
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   Begin prjLEngine.usrTransPic imgAvatar 
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      MaskColor       =   16777215
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "usrSpeech.ctx":B982
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Timer timScroll 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2580
      Tag             =   "0"
      Top             =   1740
   End
   Begin VB.Image imgOnTopOfTxt 
      Height          =   75
      Left            =   60
      Picture         =   "usrSpeech.ctx":B988
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4380
   End
   Begin VB.Label lblSay 
      BackStyle       =   0  'Transparent
      Caption         =   "[ Nothing Was Said"
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
      Height          =   615
      Left            =   705
      TabIndex        =   0
      Top             =   60
      Width           =   3735
   End
End
Attribute VB_Name = "usrSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Put This code in a module
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Const EM_GETLINECOUNT = &HBA

Private lNoLines As Long

Private Function NoOfLines() As Long
   On Error Resume Next
   NoOfLines = SendMessage(Text1.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&)
   On Error GoTo 0
   
End Function
Private Sub timScroll_Timer()

    timScroll.Tag = timScroll.Tag + 1
    
    If timScroll.Tag < 19 Then
        lblSay.Top = lblSay.Top - 2
        lblSay.Height = lblSay.Height + 2
    Else
        lNoLines = lNoLines - 2

        timScroll.Enabled = False
        timScroll.Tag = 0
    End If

End Sub

Function CanScroll() As Boolean

    If lNoLines > 2 Then
        CanScroll = True
    Else
        CanScroll = False
    End If

End Function

Sub Scroll()
    
    lNoLines = lNoLines - 2
    
    lblSay.Top = lblSay.Top - 36
    lblSay.Height = lblSay.Height + 36
    
End Sub

Private Sub UserControl_Initialize()
    imgAvatar.MaskColor = lTrans
    
    'Make Pink Trans
    UserControl.MaskPicture = UserControl.Picture
    UserControl.MaskColor = lTrans
End Sub

Private Sub UserControl_Resize()
    lblSay.Move lblSay.Left, lblSay.Top, (UserControl.Width / 15) - lblSay.Left - 6
End Sub

Sub Say(sName As String, sText As String)

    On Error Resume Next
    Set imgAvatar.Picture = LoadPicture(sName & "\avatar.bmp")

    If Err Then
        'Hide Avatar
    
        lblSay.Left = 8
        imgAvatar.Visible = False
    Else
        'Show It
    
        lblSay.Left = 47
        imgAvatar.Visible = True
    End If
    
    lblSay.Width = (UserControl.Width / 15) - lblSay.Left - 6
    
    Text1.Left = lblSay.Left
    Text1.Width = lblSay.Width

    lblSay.Top = 4
    lblSay.Height = 40
    
    lblSay.Caption = sText
    Text1.Text = sText
    
    lNoLines = NoOfLines
    
    lblSay.Visible = True

End Sub
