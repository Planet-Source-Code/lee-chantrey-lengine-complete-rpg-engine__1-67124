VERSION 5.00
Begin VB.UserControl usrNumbers 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer timAnim 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timAnim 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timAnim 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timAnim 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timAnim 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timAnim 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   20
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "usrNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private imgNumbers(5) As Image
Private iimgStart As Integer
Private Parent

Public Event OnDamageFinished()

Sub CreateDamage(lAmount As Integer, newParent, iTarget)

Set Parent = newParent

    Dim lLen As Long, I As Integer
    lLen = Len(CStr(lAmount)) - 1
    
    iimgStart = Parent.CreateImageArray
    For I = 0 To 5
        Set imgNumbers(I) = Parent.imgNumbers(I + iimgStart)
        
        imgNumbers(I).Tag = 0

        timAnim(I).Tag = 0
        timAnim(I).Interval = 60
    Next
    
    imgNumbers(0).Visible = True
    imgNumbers(0).Top = (iTarget.Top + iTarget.Height) - imgNumbers(0).Height
    
    If iTarget.Width > lLen * 8 Then
        'Center of itarget, Center of Damage
        imgNumbers(0).Left = iTarget.Left + (iTarget.Width / 2) - (((lLen + 1) * 8) / 2)
    Else
        imgNumbers(0).Left = iTarget.Left
    End If

    With imgNumbers(0)
        .Picture = GetNumero(CLng(Mid(CStr(lAmount), 1, 1)))
        .Tag = "0"
    End With
    
    With timAnim(0)
        .Interval = 60
    End With
    
    For I = 1 To lLen
        With imgNumbers(I)
            '.Caption = Mid(CStr(lAmount), I + 1, 1)
            .Top = imgNumbers(0).Top
            .Picture = GetNumero(Mid(CStr(lAmount), I + 1, 1))
            
            .Left = imgNumbers(I - 1).Left + imgNumbers(I - 1).Width
            
            .Tag = "0"
            .ZOrder 0
        End With
        
        With timAnim(I)
            .Interval = 60
        End With
    Next

    timAnim(0).Enabled = True

End Sub

Private Sub timAnim_Timer(Index As Integer)

    imgNumbers(Index).Visible = True

    If timAnim(Index).Tag = "wait" Then
        Parent.DeleteImageArray iimgStart
        
        timAnim(Index).Tag = ""
        timAnim(Index).Enabled = False
    
        RaiseEvent OnDamageFinished
        Exit Sub
    End If

    If imgNumbers(Index).Tag = 8 Then
        If timAnim.UBound = Index + 1 Or timAnim.Count = 1 Then
            timAnim(Index).Tag = "wait"
            timAnim(Index).Interval = 1000
        Else
            timAnim(Index).Enabled = False
        End If
    End If

    If imgNumbers(Index).Tag < 4 Then
        'Move imgnumbers up
        imgNumbers(Index).Tag = imgNumbers(Index).Tag + 1
        imgNumbers(Index).Top = imgNumbers(Index).Top - 2
        
    ElseIf imgNumbers(Index).Tag > 3 Then

        If timAnim.Count > (Index + 1) Then
            timAnim(Index + 1).Enabled = True
        End If
    
        'Move imgnumbers down
        imgNumbers(Index).Tag = imgNumbers(Index).Tag + 1
        imgNumbers(Index).Top = imgNumbers(Index).Top + 2
        
    End If

End Sub

