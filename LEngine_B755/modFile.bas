Attribute VB_Name = "modFile"
Function ChangeLine(sPath As String, iLine As Integer)

    Dim iFree As Integer, strBuff As String
    iFree = FreeFile
    
    Open sPath For Binary As iFree
        While EOF(iFree) = False
        
            strBuff = Space(LOF(iFree))
            Get #1, , strBuff
        
            MsgBox strBuff
            
        Wend
    Close #iFree

End Function
