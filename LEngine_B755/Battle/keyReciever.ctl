VERSION 5.00
Begin VB.UserControl keyReciever 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   2430
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   1320
      Tag             =   "0"
      Top             =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Keys"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "keyReciever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private iKeys() As Integer
Private iKeysA() As Integer

Public Event OnKeyDown(ByVal KeyAscii As Integer)
Public Event OnKeyPressed(ByVal KeyAscii As Integer)
Public Event Timer()

Private Index As Integer

Option Explicit

Property Let Enabled(newEnabled As Boolean)
    Timer1.Enabled = newEnabled
End Property

Sub SetSize(iSize As Integer)
    ReDim iKeys(iSize)
    ReDim iKeysA(iSize)
End Sub

Sub ClearKeys()

Dim I As Integer
    
    For I = 0 To UBound(iKeysA)
        iKeys(I) = 0
        iKeysA(I) = 0
    Next
    Index = 0

End Sub

Sub AddKey(KeyAscii As Integer)

    iKeys(Index) = KeyAscii
    
    Index = Index + 1
    
End Sub

Private Sub Timer1_Timer()

    Dim I As Integer
    
    For I = 0 To UBound(iKeysA)
        If iKeysA(I) <> 0 Then
            If iKeysA(I) = 1 And GetAsyncKeyState(iKeys(I)) = 0 Then
            
                If iKeys(I) = 27 Then
                    'Esc
                    iKeysA(I) = 0
                    Call frmMain.ToggleFileMenu
                Else
                
                    iKeysA(I) = 0
                    RaiseEvent OnKeyPressed(iKeys(I))
                End If
            End If
        End If
    Next
    
    For I = 0 To UBound(iKeysA)
        If iKeys(I) <> 0 Then
            If GetAsyncKeyState(CLng(iKeys(I))) <> 0 Then
            
                iKeysA(I) = 1
                RaiseEvent OnKeyDown(iKeys(I))

                'Only Report 1 KeyDown at a time
                Exit For
            End If
        End If
    Next
    
    RaiseEvent Timer

End Sub

Private Sub UserControl_Initialize()

    If NotInDbg = True Then
        Timer1.Enabled = True
    End If
    
    ReDim iKeys(12)
    ReDim iKeysA(12)
    
    'Add Esc Key
    Me.AddKey 27
        
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 810
    UserControl.Width = 810
End Sub

Private Sub UserControl_Terminate()
    Timer1.Enabled = False
End Sub
