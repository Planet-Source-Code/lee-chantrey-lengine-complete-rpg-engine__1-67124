VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cValues As New Collection
Private cKeys As New Collection
Private cLock As New Collection
Private iCount As Integer

Option Explicit

Public Property Let Locked(lIndex, bLocked As Boolean)
    ColSetValue cLock, lIndex, bLocked
End Property

Public Function Exists(Index) As Boolean
    Exists = ColExists(cValues, Index)
End Function

Public Function ClearList()
    
Dim I As Integer, cRemove As New Collection
    
    For I = 1 To cValues.Count
        If ColExists(cLock, cKeys(I)) = False Then
            'Remove I
            cRemove.Add cKeys(I)
        End If
    Next
    
    While cRemove.Count > 0
        If Remove(cRemove(1)) = False Then
            MsgBox "Critical Error [Collection::ClearList] Memory Failed!", vbCritical
            End
        End If
        
        cRemove.Remove 1
    Wend
    
End Function

Public Function Remove(Index) As Boolean

    Remove = True

    If Me.Exists(Index) = False Then
        Remove = False
        Exit Function
    End If

    iCount = iCount - 1
    
    cValues.Remove Index
    cKeys.Remove Index
End Function

Public Function Count()
    Count = cValues.Count
End Function

Public Function Add(Item, Key, Optional Index = 0)

    'MsgBox Item & ":" & Key

    iCount = iCount + 1

    If ColExists(cValues, Key) = True Then
        Remove Key
    End If

    If (cValues.Count + 1 > Index) And (Index > 0) Then
        'Put into specified index
        cValues.Add Item, Key, Index
        cKeys.Add Key, Key, Index
    Else
        cValues.Add Item, Key
        cKeys.Add Key, Key
    End If
        
End Function

Public Function GetItem(Index, Optional Default = "")

    On Error GoTo Catch_E
    GetItem = cValues(Index)
    
    Exit Function
Catch_E:

    If Default = "" Then
        MsgBox "GetItem Failed: ", vbCritical
    End If
        
    GetItem = Default

End Function

Property Get Key(Index)
    If Index = 0 Then
        MsgBox "SuperCollection Invalid Index: 0", vbCritical
        Exit Property
    End If
    
    Key = cKeys(Index)
End Property

Property Get Item(Index)
Attribute Item.VB_UserMemId = 0
    Item = GetItem(Index)
End Property

Property Let Item(Index, newValue)
    
    On Error GoTo Catch_E
    
    Dim sKey, sValue
    
    cValues.Remove Index
        
    sKey = cKeys(Index)
    cKeys.Remove Index
    
    Add newValue, sKey, Index
    
    Exit Property
Catch_E:

End Property
