Attribute VB_Name = "modColLib"
'Usefull Collection Functions
'Lee Matthew Chantrey

Public Function UpdateColDirty(ByRef Col As Collection, Key, sUpdate As String)
    'Updates Collection Key, Ignores numerical indexes

    Col.Remove Key
    Col.Add sUpdate, Key

End Function

Public Function UpdateCol(ByRef Col As Collection, Index, sUpdate As String)
    'Updates Collection Key, Keeps Numerical Index intact

    Col.Remove Index
    
    If Col.Count + 1 > Index Then
        Col.Add sUpdate, , Index
    Else
        Col.Add sUpdate
    End If

End Function

Public Function ColExists(ByRef Col As Collection, Key) As Boolean

    'Check Key Exists in Collection
    ColExists = True

Dim sTemp As String

    On Error GoTo CatchErr
    ColExists = True
    
    sTemp = Col(Key)
    Exit Function
    
CatchErr:
    ColExists = False

End Function

Public Function ColAmend(ByRef Col As Collection, ByVal Key, ByVal sAmend As String)

    'Ammend a previous value

    On Error GoTo CatchErr
    
    If ColExists(Col, Key) Then
        sAmend = Col(Key) & sAmend
        Col.Remove Key
    End If
    
    Col.Add sAmend, Key
    
    Exit Function
CatchErr:
    MsgBox Err.Description

End Function

Public Function ColSetValue(ByRef Col As Collection, Key, Value)

    'Set Col Value (Remove if it exists)

    On Error GoTo CatchErr
    
    If ColExists(Col, Key) Then
        Col.Remove Key
    End If
    
    Col.Add Value, Key
    
    Exit Function
CatchErr:
    MsgBox Err.Description

End Function

Public Function ColValue(ByRef Col As Collection, Key, Default) As Variant

    'Retrieve Col Value (Dont Error If it doesnt exist)

    On Error GoTo CatchErr
    ColValue = Col(Key)
    
    Exit Function
CatchErr:
    ColValue = Default

End Function

Public Function RemoveCol(ByRef Col As Collection, Key) As Boolean

    'Remove Key (Dont Error If it doesnt exist)

    On Error Resume Next
    Col.Remove Key
    
    If Err Then
        RemoveCol = False
        Exit Function
    End If
    
    RemoveCol = True

End Function
