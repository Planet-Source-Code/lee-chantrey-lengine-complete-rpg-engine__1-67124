Attribute VB_Name = "modStoryOps"
Option Explicit
Private cUsrVar As New Collection

Function ReplaceOutBrack(ByRef sData As String)

    sData = Trim(sData)

    Dim I As Integer, lC As Integer, C As String, bInQ As Boolean, bInBrack As Boolean
    lC = Len(sData) + 1
    
    'Shave Off Brackets
    If Left(sData, 1) = "(" Then
        sData = Mid(sData, 2)
    End If
    
    If Right(sData, 1) = ")" Then
        sData = Mid(sData, 1, Len(sData) - 1)
    End If
    
    I = 1
    While I < lC
        C = Mid(sData, I, 1)

        If C = "'" Then
            If bInQ = False Then
                bInQ = True
            Else
                bInQ = False
            End If
        End If

        If C = "(" Then
            bInBrack = True
        ElseIf C = ")" Then
            bInBrack = False
        End If
    
        If C = "'" And bInBrack = False Then
            Mid(sData, I, 1) = """"
        End If
    
        I = I + 1
    Wend

End Function

Public Function VarScan(ByVal sData As String)

Dim p1 As Integer, p2 As Integer, SVar As String, sP1 As String, sP2 As String, _
    sVarValue As String
    
    p1 = 1
    p2 = 1
    
    While p1 > 0
        p1 = InStr(p1, sData, "%")
        p2 = p1 + 1
        p1 = InStr(p1 + 1, sData, "%")
        
        If p1 > 0 And p2 > 0 Then
            'Found Variable
            SVar = Mid(sData, p2, p1 - p2)

            sP1 = Mid(sData, 1, p2 - 2)
            sP2 = Mid(sData, Len(sP1) + Len(SVar) + 3)
            
            'Put it into data
            sVarValue = UsrVarValue(SVar, "%" & SVar & "%")
            
            'Check if variable has variable
            While IsVariable(sVarValue) = True
                sVarValue = UsrVarValue(sVarValue, "<" & SVar & ">")
            Wend

            sData = sP1 & sVarValue & sP2
            p1 = InStr(Len(sP1) + Len(sVarValue), sData, "%")
        End If
    Wend
    
    VarScan = sData

End Function

Public Function OutErr(ByVal sParam As String, ByVal iMsg As Integer)

    Dim sOut As String

    Select Case iMsg
    
    Case 1
        sOut = "'" & sParam & "' is not a valid variable name"
        
    Case 2
        sOut = "'" & sParam & "' is not numeric"
        
    Case 3
        sOut = sParam
        
    End Select
    
    MsgBox sOut, vbCritical, "Story Script Error"

End Function

Private Function VarCheck(ByRef sValue1 As String, Optional bWarn As Boolean = True) As Boolean
    'Strip %

    VarCheck = False

    If IsVariable(sValue1) = True Then
        sValue1 = Mid(sValue1, 2, Len(sValue1) - 2)
        
        If UsrVarExists(sValue1) = True Then
            sValue1 = UsrVarValue(sValue1)
            VarCheck = True
        Else
            'Give variable atributes back
            sValue1 = "%" & sValue1 & "%"
        End If
    Else
        If UsrVarExists(sValue1) = True Then
            sValue1 = UsrVarValue(sValue1)
            VarCheck = True
        Else
            If bWarn = True Then
                OutErr sValue1, 1
            End If
            VarCheck = False
        End If
    End If

End Function

Function IsSkip(sData As String) As Boolean

    IsSkip = False
    
    Dim sChar As String
    sChar = Left(sData, 1)
    
    If sChar = "+" Then
        IsSkip = True
    End If

End Function

Private Function UsrVarExists(sName As String) As Boolean

Dim sTemp As String

    On Error GoTo CatchErr
    UsrVarExists = True
    
    sTemp = cUsrVar(sName)
    Exit Function
    
CatchErr:
    If IsVariable(sName) = True Then
        MsgBox "Unacceptable variable pass! [VarExists]"
        End
    End If

    UsrVarExists = False

End Function

Function UsrVarValue(sName As String, Optional Default = "") As String

    On Error GoTo CatchErr
    
    If Len(cUsrVar(sName)) > 0 Then
        UsrVarValue = cUsrVar(sName)
        Exit Function
    End If
    
    Exit Function
    
CatchErr:
    If IsVariable(sName) = True Then
        WarnUser "Invalid Variable: '" & sName & "'", False
    End If

    UsrVarValue = Default

End Function

Private Function IsVariable(ByVal sData As String) As Boolean

    If Left(sData, 1) = "%" And Right(sData, 1) = "%" Then
        IsVariable = True
    End If

End Function

Function Op_IntAdd(ByVal sName As String, ByVal sValue As String)

    If IsNumeric(sValue) = False Then
        OutErr sValue, 2
        Exit Function
    End If
    
    Dim sValue2 As String
    sValue2 = UsrVarValue(sName, "0")
        
    If IsNumeric(sValue2) = False Then
        OutErr sValue2, 2
        Exit Function
    End If

    Op_Variable sName, CInt(sValue2) + CInt(sValue), True
    
    If Err Then
        MsgBox Err.Description
    End If
    
End Function

Function Op_Variable(ByVal sName As String, ByVal sValue As String, Optional bAllowPure As Boolean = True)
    'Make the variable sName, but if it exists overwrite

    If IsVariable(sName) = True Then
        sName = Mid(sName, 2, Len(sName) - 2)
    ElseIf bAllowPure = False Then
        OutErr sName, 1
        Op_Variable = False
    End If
    
    If UsrVarExists(sName) = False Then
        VarsWin.Add sName, sValue
        cUsrVar.Add sValue, sName
    Else
        cUsrVar.Remove sName

        If sValue <> "" Then
            cUsrVar.Add sValue, sName
            VarsWin.Add sName, sValue
        Else
            VarsWin.Remove sName
        End If
        
        'VarsWin.Update sName, sValue
    End If

End Function

Function Op_StrCmp(ByVal sValue1 As String, ByVal sValue2 As String, ByVal sTrue As String, Optional ByVal sFalse As String)

    If sValue1 = sValue2 Then
        'Do True
        ExeStoryCmd sTrue
    Else
        If sFalse <> vbNullString Then
            'Do False
            ExeStoryCmd sFalse
        End If
    End If

End Function
