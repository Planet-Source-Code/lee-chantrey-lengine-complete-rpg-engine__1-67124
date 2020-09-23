VERSION 5.00
Begin VB.UserControl usrTransSpritePic 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Image imgS 
      Height          =   615
      Index           =   0
      Left            =   420
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "usrTransSpritePic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sPath As String
Private lIndex As Integer
Private bFrameLoop
Private bVisible As Boolean

Public Event OnX(iX As Single)
Public Event OnY(iY As Single)
Public Event GetX(ByRef iX As Single)
Public Event GetY(ByRef iY As Single)
Public Event Visibability(bVisible As Boolean)
Public Event OnUnload()

Public Function Unload()
    RaiseEvent OnUnload
End Function

Public Property Let Visible(bNewVisible As Boolean)
    bVisible = bNewVisible
    
    RaiseEvent Visibability(bVisible)
End Property

Public Property Get Visible() As Boolean
    Visible = bVisible
End Property

Public Property Get Width() As Single
    Width = UserControl.Width
End Property

Public Property Get Height() As Single
    Height = UserControl.Height
End Property

Public Property Let Source(newPath As String)
    sPath = sPath_SFX & newPath

    Set Me.Picture = LoadPicture(sPath)
    Me.MaskColor = lTrans
End Property

Public Property Get Source() As String
    Source = Replace(sPath, sPath_SFX, "")
End Property

Public Property Let X(iNewX As Single)
    RaiseEvent OnX(iNewX)
End Property

Public Property Let Y(iNewY As Single)
    RaiseEvent OnY(iNewY)
End Property

Public Property Get X() As Single

Dim iRet As Single
    RaiseEvent GetX(iRet)
    
    X = iRet
    
End Property

Public Property Get Y() As Single

Dim iRet As Single
    RaiseEvent GetY(iRet)
    
    Y = iRet

End Property

Public Function FrameUbound() As Integer
    FrameUbound = imgS.UBound
End Function

Public Property Let FrameLoop(sNew_bFrameLoop As Boolean)
    bFrameLoop = sNew_bFrameLoop
End Property

Public Property Get FrameLoop() As Boolean
    FrameLoop = bFrameLoop
End Property

Public Property Get Frame() As Integer
    Frame = lIndex
End Property

Public Sub NextFrame()
    
    If lIndex < imgS.UBound Then
        Me.Frame = lIndex + 1
    Else
        Me.Frame = 0
    End If
    
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackStyle = 0
    bVisible = True
End Sub

Public Property Let AutoSize(newB As Boolean)

    Dim pSize As picSize
    
    If newB = True Then
        pSize = GetImageSize(Me.Picture)
        
        UserControl.Height = pSize.Height
        UserControl.Width = pSize.Width
    End If
    
End Property

Public Property Get MaskColor() As Long
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As Long)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

Public Property Get Picture() As IPictureDisp
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As IPictureDisp)
    Set UserControl.Picture = New_Picture
    Set UserControl.MaskPicture = New_Picture
    PropertyChanged "Picture"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", vbWhite)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, lTrans)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

Property Let Frame(lFrameIndex As Integer)

    Set UserControl.Picture = imgS(lFrameIndex).Picture
    Set UserControl.MaskPicture = imgS(lFrameIndex).Picture
    lIndex = lFrameIndex
    
End Property

Function RetrieveImage(lStance As Integer)
    Set RetrieveImage = imgS(lStance).Picture
End Function

Function LoadAdditional(sPath As String)

    Load imgS(imgS.Count)
    With imgS(imgS.UBound)
        .Picture = LoadPicture(sPath)
    End With

End Function

Function LoadStore(sNewPath As String, Optional iMax As Integer)

Dim I As Integer

    On Error GoTo Catch_E

    'Set Path Property
    sPath = sPath_SFX & sNewPath
    
    'Add trailing slash if its not present
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    
    'Unload existing
    For I = 1 To imgS.UBound
        Unload imgS(I)
    Next
    If iMax = 0 Then
        iMax = imgS.UBound
    End If

    imgS(0).Picture = LoadPicture(sPath & "a0.bmp")
    
    'Find the size
    frmLib.Picture1 = LoadPicture(sPath & "a0.bmp")
    
    UserControl.Height = frmLib.Picture1.Height
    UserControl.Width = frmLib.Picture1.Width
    
    For I = 1 To iMax
        Load imgS(imgS.Count)
        With imgS(imgS.UBound)
            .Picture = LoadPicture(sPath & "a" & I & ".bmp")
        End With
    Next
    
    Exit Function
    
Catch_E:
    If Err.Number = 53 Then
        WarnUser "Frame " & I & " is missing: " & sNewPath
    Else
        WarnUser Err.Description
    End If

End Function
