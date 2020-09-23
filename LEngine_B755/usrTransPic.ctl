VERSION 5.00
Begin VB.UserControl usrTransPic 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H008000FF&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image imgS 
      Height          =   615
      Index           =   0
      Left            =   600
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "usrTransPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sPath As String
Private lIndex As Integer
Private bFrameLoop

Public Function FrameUbound() As Integer
    FrameUbound = imgS.UBound
End Function

Public Property Let FrameLoop(sNew_bFrameLoop As Boolean)
    bFrameLoop = sNew_bFrameLoop
End Property

Public Property Get FrameLoop() As Boolean
    FrameLoop = bFrameLoop
End Property

Public Property Get StanceIndex() As Integer
    StanceIndex = lIndex
End Property

Public Property Let Path(newPath As String)
    sPath = newPath
End Property

Public Property Get Path() As String
    Path = sPath
End Property

Public Sub NextStance()
    
    If lIndex < imgS.UBound Then
        ShowStance lIndex + 1
    Else
        ShowStance 0
    End If
    
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackStyle = 0
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

Function ShowStance(lStance As Integer)

    Set UserControl.Picture = imgS(lStance).Picture
    Set UserControl.MaskPicture = imgS(lStance).Picture
    lIndex = lStance
    
End Function

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

    On Error GoTo Catch_E

    'Set Path Property
    sPath = sNewPath
    
    'Add trailing slash if its not present
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

Dim I As Integer
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

