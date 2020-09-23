VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.UserControl usrSprite 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin prjLEngine.usrTransSpritePic Images 
      Height          =   240
      Index           =   0
      Left            =   3540
      TabIndex        =   0
      Top             =   660
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      MaskColor       =   16777215
   End
   Begin VB.Timer Timers 
      Index           =   0
      Left            =   4080
      Top             =   1200
   End
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "usrSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event OnFinished()

Private cTimers As New Collection
Private cFreeTimer As New Collection
Private cImages As New Collection

Public Function ExecuteStatement(Statement)
    
    Script.ExecuteStatement Statement
    
End Function

Public Function AddObject(sName As String, Object)

    Script.AddObject sName, Object, False
    
End Function

Public Function Finished()

    RaiseEvent OnFinished
    Reset False
    
End Function

Private Function GetFreeTimer() As Integer

Dim I As Integer

    If cFreeTimer.Count > 0 Then
        GetFreeTimer = cFreeTimer(1)
        cFreeTimer.Remove 1
        
        Exit Function
    End If
    
    GetFreeTimer = Timers.Count

End Function

Sub Reset(Optional bResetScript As Boolean = True)

On Error GoTo Catch_E

Dim I As Integer

    For I = 1 To Images.UBound
        If IsObject(Images(I)) Then
            Unload Images(I)
        End If
    Next
    For I = 1 To Timers.UBound
        If IsObject(Images(I)) Then
            Unload Timers(I)
        End If
    Next
    
    Set cTimers = New Collection
    Set cFreeTimer = New Collection
    Set cImages = New Collection

    If bResetScript = True Then
        Script.Reset
        Script.AddObject "Base", Me, True
    End If
        
    Exit Sub
Catch_E:
    MsgBox "Reset: " & Err.Description, vbCritical

End Sub

Private Sub Images_GetX(Index As Integer, iX As Single)

    iX = Images(Index).Left

End Sub

Private Sub Images_GetY(Index As Integer, iY As Single)

    iY = UserControl.ScaleHeight - (Images(Index).Top + Images(Index).Height)

End Sub

Private Sub Images_OnUnload(Index As Integer)

    If modColLib.ColExists(cImages, "a" & Index) = False Then

        Unload Images(Index)
        cImages.Add Index, "a" & Index
        
        UserControl.Refresh
        
    Else
    
        MsgBox "ImageHolder " & Index & " is already empty!", vbCritical
    
    End If

End Sub

Private Sub Images_OnX(Index As Integer, iX As Single)

    With Images(Index)
        .Left = iX
    End With

End Sub

Private Sub Images_OnY(Index As Integer, iY As Single)

    With Images(Index)
        .Top = UserControl.ScaleHeight - (iY + Images(.Index).Height)
    End With

End Sub

Private Sub Images_Visibability(Index As Integer, bVisible As Boolean)
    Images(Index).Visible = bVisible
End Sub

Private Sub Timers_Timer(Index As Integer)
    On Error Resume Next
    
    Script.Run Timers(Index).Tag
    
    If Err Then
        MsgBox "Error: " & Script.Error.Description & vbCr & _
               "Line: " & Script.Error.Line, vbCritical
               
        Timers(Index).Enabled = False
    End If
End Sub

Function LoadScript(sPath As String) As Boolean

Dim sData As String

    LoadScript = False
    
    On Error Resume Next
    
    sData = Fload(sPath_SFX & sPath, False)
    
    If sData <> "" Then
        Script.AddCode sData
        LoadScript = True
    End If

    If Err Then
        MsgBox "Error: " & Script.Error.Description & vbCr & _
               "Line: " & Script.Error.Line
    End If
End Function

Sub TimedEvent(sName As String, iInterval As Integer)

    On Error Resume Next

    Dim tIndex As Integer
    tIndex = GetFreeTimer

    Load Timers(tIndex)
    With Timers(tIndex)
        .Interval = iInterval
        .Tag = sName
        .Enabled = True
        
        cTimers.Add .Index, sName
    End With
    
    If Err Then
        If Err.Number = 457 Then
            MsgBox "TimedEvent Failed: " & sName & " is already a timed event."
        Else
            MsgBox "TimedEvent Failed: " & Err.Description, Err.Number
        End If
    End If

End Sub

Function LoadImage(Optional sPath As String, Optional X = 0, Optional Y = 0) As usrTransSpritePic

    On Error Resume Next

Dim iFreeImg As Integer
    iFreeImg = GetFreeImage

    Load Images(iFreeImg)
    Dim clsImg As New clsImage
    
    With Images(iFreeImg)
        .Visible = False
    
        clsImg.AttachClient Images(.Index)
        
        If sPath <> "" Then
            clsImg.Source = sPath
        End If
    
        '.Picture = LoadPicture(sPath)
        '.Stretch = False
        .Tag = sName

        .MaskColor = lTrans
        
        Set LoadImage = Images(Images.UBound)
        .AutoSize = True
        
        .Left = X
        .Top = UserControl.ScaleHeight - (Y + Images(.Index).Height)
        
        .Visible = True
    End With
    
    If Err Then
        MsgBox "LoadImage Failed: " & Err.Description
        Exit Function
    End If

End Function

Function GetFreeImage() As Integer

    If cImages.Count = 0 Then
        GetFreeImage = Images.Count
    Else
        GetFreeImage = cImages(1)
        cImages.Remove 1
    End If

End Function

Function StopTimed(sName As String)

    On Error Resume Next

    Timers(cTimers(sName)).Enabled = False
    
    'Freeup Timer
    cFreeTimer.Add cTimers(sName)
    Unload Timers(cTimers(sName))
    
    cTimers.Remove sName
    
    If Err Then
        If Err.Number = 5 Then
            MsgBox "StopTimed Failed: " & sName & " is not a timed event."
        Else
            MsgBox Err.Description
        End If
    End If

End Function

