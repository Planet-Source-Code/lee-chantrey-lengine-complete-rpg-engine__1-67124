Attribute VB_Name = "modInventorySupport"
Public Sub GetItemTypes()

    On Error GoTo Catch_E

Dim I As Integer, cItems As New SuperCollection, sItem As String, sP() As String
    With frmLib.File1
        .Path = sPath_Items
        .Pattern = "*.def"
        
        For I = 0 To .ListCount - 1
            sItem = .List(I)
            sP = Split(sItem, ".def", 2)

            cItems.Add 0, StrConv(sP(0), vbProperCase)
        Next
    End With
    Inventory.Types = cItems
    
    Exit Sub
    
Catch_E:
    WarnUser "modInventorySupport:GetItemTypes(): '" & sPath_Items & "' " & Err.Description, False

End Sub
