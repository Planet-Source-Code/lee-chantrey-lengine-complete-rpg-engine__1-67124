VERSION 5.00
Begin VB.UserControl usrTransition 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer timShrink 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3120
      Top             =   2400
   End
   Begin VB.Timer timGrow 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   2640
      Top             =   2400
   End
   Begin VB.PictureBox picSrc 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Shape shp 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   15
         Left            =   0
         Top             =   1800
         Visible         =   0   'False
         Width           =   4800
      End
   End
End
Attribute VB_Name = "usrTransition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event OnFinishedGrow()
Public Event OnFinishedShrink()

Private bShrank As Boolean
Private bGrow As Boolean

Private Declare Function BitBlt Lib "gdi32" _
        (ByVal hDestDC As Long, ByVal X As Long, _
        ByVal Y As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc _
        As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Sub Capture()

    BitBlt picSrc.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hdc, 0, 0, vbSrcCopy
    picSrc.Visible = True

End Sub

Private Sub timShrink_Timer()
    
    On Error GoTo Catch_E
    
    shp.Height = shp.Height - 20
    shp.Top = shp.Top + 10
    
    If shp.Height = 1 Then
        timShrink.Enabled = False
        shp.Visible = False

        bShrank = True
        RaiseEvent OnFinishedShrink
    End If
    
    Exit Sub
    
Catch_E:

    shp.Height = 1
    shp.Visible = False
    
    bShrank = True

End Sub

Private Sub timGrow_Timer()

    On Error GoTo Catch_E

    shp.Height = shp.Height + 20
    shp.Top = shp.Top - 10
    
    If shp.Height = UserControl.ScaleHeight Then
        timGrow.Enabled = False
    
        bShrank = False
        bGrow = False
        
        RaiseEvent OnFinishedGrow
    End If
    
    Exit Sub
    
Catch_E:
    shp.Height = UserControl.ScaleHeight
    
    bShrank = False
    bGrow = False
        
    RaiseEvent OnFinishedGrow
    
End Sub

Function Shrink() As Boolean

    Shrink = True

    If bShrank = True Then
        Shrink = False
        Exit Function
    End If

    picSrc.Visible = False
    Set shp.Container = Me
    
    timShrink.Enabled = True
    
End Function

Sub Grow()

    If bGrow = True Then
        Exit Sub
    End If

    bGrow = True

    Set shp.Container = picSrc

    Capture

    shp.Height = 1
    shp.Top = 120
    
    shp.Visible = True

    timGrow.Enabled = True
End Sub

Private Sub UserControl_Initialize()
    bShrank = True
    bGrow = False
End Sub
