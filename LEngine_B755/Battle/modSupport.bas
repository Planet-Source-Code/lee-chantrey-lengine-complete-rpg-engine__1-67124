Attribute VB_Name = "modSupport"
Public Function GetDirection(KeyCode As Integer) As Integer

    Select Case KeyCode
    
    Case 37
        GetDirection = 1
    
    Case 39
        GetDirection = 2
        
    Case 40
        GetDirection = 3
        
    Case 38
        GetDirection = 4
        
    Case Else
        GetDirection = 17
    
    End Select

End Function

Private Function InDbg() As Boolean

  On Error Resume Next
  Debug.Assert 1 / 0
  InDbg = (Err <> 0)
  
End Function

