VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cTypes As New SuperCollection

Public Property Let Types(cNewTypes As SuperCollection)
    Set cTypes = cNewTypes
End Property

Public Property Get Types() As SuperCollection
    Set Types = cTypes
End Property

Public Function TypeCount(sType As String) As Integer
    If cTypes.Exists(sType) = False Then
        TypeCount = 0
        
        MsgBox "clsInventory:TypeCount: Type does not exist! '" & sType & "'", vbCritical
        Exit Function
    End If

    TypeCount = cTypes.GetItem(sType)
End Function

Public Function TakeItem(sType As String)
    If cTypes.Exists(sType) = False Then
        MsgBox "clsInventory:TakeItem: Type does not exist! '" & sType & "'", vbCritical
        Exit Function
    End If
    
    cTypes.Item(sType) = cTypes.GetItem(sType) - 1
End Function

Public Function AddItem(sType As String)
    If cTypes.Exists(sType) = False Then
        MsgBox "clsInventory:AddItem: Type does not exist! '" & sType & "'", vbCritical
        Exit Function
    End If
    
    cTypes.Item(sType) = cTypes(sType) + 1
End Function
