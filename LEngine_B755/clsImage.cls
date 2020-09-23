VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Img
Dim sSource As String

Function AttachClient(ByRef newImg)
    Set Img = newImg
End Function

Property Get Source() As String
    Source = sSource
End Property

Property Let Source(sNewPicture As String)
    sSource = sNewPicture
    
    Set Img.Picture = LoadPicture(sPath_SFX & "\" & sNewPicture)
End Property
