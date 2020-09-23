Attribute VB_Name = "modSupport"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Integer
    
Private Declare Sub SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ShowCursor Lib "user32" _
    (ByVal bShow As Long) As Long
    
Public NotInDbg As Boolean
Public sFloadLast As String

Public Type picSize
    Height As Integer
    Width As Integer
End Type

Public Type Cords
    X As Long
    Y As Long
End Type

Public Function OutErrMsg(sDescription As String, sModule As String, sProcedure As String)

Dim sData As String

    sData = sProcedure & "::" & sModule & vbCrLf & _
            sDescription
            
    MsgBox sData, vbCritical, "Error"

End Function

Public Function InDbg() As Boolean

  On Error Resume Next
  Debug.Assert "TESTDEBUG"
  If Err = 0 Then
    ' Runtime variant
    InDbg = False
  Else
    ' Debug variant
    InDbg = True
  End If

End Function

Public Function Num2Col(iData As Integer) As String
    'You cant have collections with numerical keys
    'So this function convers numerical to alphanumerical
    
    Num2Col = CStr("a" & iData)
End Function

Public Function Qo(sData As String) As String

    Qo = """" & sData & """"

End Function

Public Function CopyBattleChar(ByRef Src As clsBattlePlayer, ByVal Dest As clsBattlePlayer)
    Set Src = Dest
End Function

Public Function FindBattleCharIndex(sId As String) As Integer
'Finds the human battle character's array index

Dim I As Integer

    sId = LCase(sId)
    
    FindBattleCharIndex = -1

    For I = 0 To UBound(pBattleHu)
        If LCase(pBattleHu(I).ID) = sId Then
            FindBattleCharIndex = I
            
            Exit Function
        End If
    Next

End Function

Public Function Hide_Cursor()

    While ShowCursor(False) >= 0
    Wend
    
End Function

Public Function Show_Cursor()

    While ShowCursor(True) <> 0
    Wend

End Function

Public Function FormOnTop(lHwnd As Long)

    SetWindowPos lHwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1

End Function

Public Function KillHome(ByVal sPath As String) As String

    KillHome = Replace(sPath, App.Path & "\resources\", "")

End Function

Public Function FileCount(sPath As String, Optional fPattern As String = "*.*") As Integer

Dim I As Integer, cItems As New Collection, sItem As String, sP() As String, iC As Integer

    On Error GoTo Catch_E

    With frmLib.File1
        .Path = sPath
        .Pattern = fPattern
        
        FileCount = .ListCount - 1
    End With

    Exit Function
    
Catch_E:
    FileCount = -1

End Function

Public Function FileExist(asPath As String) As Boolean

    On Error GoTo Catch_E

    If FileLen(asPath) > 0 Then
        FileExist = True
    Else
        FileExist = False
    End If
    
    Exit Function
Catch_E:
    FileExist = False
    
End Function
Public Function TrimPath(ByVal asPath As String) As String

    If Len(asPath) = 0 Then Exit Function
    Dim X As Integer
    
    Do
    X = InStr(asPath, "\")
    If X = 0 Then Exit Do
    asPath = Right(asPath, Len(asPath) - X)
    Loop
    TrimPath = asPath
    
End Function

Public Function NotifyUser(ByRef sMsg As String)

    With DebugWin
        .AddItem KillHome(sMsg)
    End With

End Function

Public Function WarnUser(ByRef sMsg As String, Optional bHalt As Boolean = True)

    'sMsg = Replace(sMsg, sPath_Resources, "")
    sMsg = KillHome(sMsg)

    If bHalt = True Then
        HaltStory
        
        sMsg = "Halt Exception: " & sMsg
        DebugWin.txtCur.ForeColor = vbRed
    Else
        sMsg = "Warning: " & sMsg
    End If

    With DebugWin
        .AddItem sMsg
        '.ScrollLast = False
    End With
    
    If frmMain.WindowState = vbMaximized Then
        MsgBox sMsg, vbCritical, "From Messages Window"
    End If

End Function

Public Function StrFront(sData As String, sDelim As String)

Dim sP() As String, I As Integer
    sP = Split(sData, sDelim)
    
    For I = 0 To UBound(sP) - 1
        StrFront = StrFront & sP(I)
    Next
    
    If StrFront = "" Then
        StrFront = sData
    End If

End Function

Public Function StrEnd(sData As String, sDelim As String, Optional iOffset As Integer = 1)

    If InStr(sData, sDelim) = 0 Then
        'Delim not present
    
        StrEnd = sData
        Exit Function
    End If

Dim iLen As Integer, iDLen As Integer

    iLen = Len(sData) + 1
    iDLen = Len(sDelim)

    If iLen = 1 Or iDLen = 0 Then
        StrEnd = False
        Exit Function
    End If

    While Mid(sData, iLen, iDLen) <> sDelim And iLen > 1
        iLen = iLen - 1
    Wend

    If iLen = 0 Then
        StrEnd = False
        Exit Function
    End If
    
    StrEnd = Mid(sData, iLen + iOffset)

End Function

Public Function Hex2VB(ByVal HexColor As String) As String

    'The input at this point could be HexColor = "#00FF1F"

Dim Red As String
Dim Green As String
Dim Blue As String

HexColor = Replace(HexColor, "#", "")
    'Here HexColor = "00FF1F"

Red = Val("&H" & Mid(HexColor, 1, 2))
    'The red value is now the long version of "00"

Green = Val("&H" & Mid(HexColor, 3, 2))
    'The red value is now the long version of "FF"

Blue = Val("&H" & Mid(HexColor, 5, 2))
    'The red value is now the long version of "1F"


Hex2VB = RGB(Red, Green, Blue)
    'The output is an RGB value

End Function

Public Function IsSomething(ByRef Object) As Boolean

    IsSomething = False

    If IsObject(Object) = False Then
        Exit Function
    End If

    On Error Resume Next
    
    'Every object should have a name
    If Object.Name <> "" Then
        IsSomething = True
        
        Exit Function
    End If
    
    If Err Then
        IsSomething = False
    End If

End Function

Public Function GetImageSize(tPicture As IPictureDisp) As picSize
    
    Set frmLib.Picture1.Picture = tPicture
    
    With frmLib.Picture1
        GetImageSize.Height = .Height
        GetImageSize.Width = .Width
    End With
    
End Function

Public Function Fload(sPath As String, Optional bDefaults As Boolean = True) As String

Dim NewLine As String
Dim sOut As String

Dim iFree As Integer

    'Load entire text file
    On Error GoTo CatchErr

    iFree = FreeFile

    Open sPath For Input As #iFree
    sFloadLast = sPath
    
    While Not EOF(iFree)
        Line Input #iFree, NewLine
        sOut = sOut & NewLine & vbCrLf
    Wend
    Close #iFree
    
    Fload = sOut
    
    Exit Function
    
CatchErr:
    If bDefaults = True Then
        NotifyUser Err.Description & ": " & sPath
        Fload = Fload(sPath_Defaults & StrEnd(sPath, "\"), False)
    Else
        WarnUser "Fload:: " & Err.Description & ": " & sPath
    End If

End Function

Public Function RandomNumber(Upper As Integer, _
    Lower As Integer) As Integer
  On Error GoTo LocalError
  
  'Generates a Random Number BETWEEN then LOWER
  'and UPPER values
  
  Randomize Time
  RandomNumber = Int((Upper - Lower + 1) * Rnd + Lower)
  
  Exit Function
LocalError:
  RandomNumber = Lower
End Function

Public Function Evaluate(sMath As String) As Integer

    Evaluate = frmLib.Math.Eval(sMath)

End Function

Public Function BolAsNO(bData As Boolean) As Integer

    BolAsNO = 0
    If bData = True Then
        BolAsNO = 1
    End If

End Function

Public Function CBol(Data) As Boolean

    CBol = False

    If IsNumeric(Data) = True Then
        If Data = "1" Then
            CBol = True
        End If
    ElseIf LCase(Data) = "true" Then
        CBol = True
    End If

End Function

Public Function GetNumero(iNum As Integer, Optional bGood As Boolean) As IPictureDisp
    On Error Resume Next
    Set GetNumero = frmLib.usrNumero.GetNumero(iNum, bGood)
End Function

Public Function GetDirection(KeyCode As Integer) As Integer

    Select Case KeyCode
    
    Case Control_Left
        GetDirection = 1
    
    Case Control_Right
        GetDirection = 2
        
    Case Control_Down
        GetDirection = 3
        
    Case Control_Up
        GetDirection = 4
    
    End Select

End Function

Function FindExistingMusic(ByRef sPath As String, sFileName As String) As Boolean

    On Error GoTo Catch_E

    FindExistingMusic = False

    'Find music file with any extension
    frmLib.File1.Path = sPath
    frmLib.File1.Pattern = sFileName & ".*.*"

    If frmLib.File1.ListCount > 0 Then
        sPath = sPath & "\" & frmLib.File1.List(0)

        FindExistingMusic = True
    End If
    
    Exit Function
Catch_E:
    NotifyUser "Finding Extension for Music Failed: " & Err.Description
    

End Function

Public Function PlayNewMusic(sPath As String)
    
    frmLib.PlayMusic sPath

End Function

Public Function InDbgTry() As Boolean
  On Error Resume Next
  Debug.Assert 1 / 0
  InDbgTry = (Err <> 0)
End Function

Public Function DivB(Val1, Val2, Default) As Integer

    If Val1 Mod Val2 <> 0 Then
        DivB = Default
    Else
        DivB = Val1 / Val2
    End If

End Function

Public Function DivA(Val1, Val2) As Integer

    Dim nRes, iStr As Integer
    nRes = Val1 / Val2

    iStr = InStr(nRes, ".")
    
    If iStr > 0 Then
        DivA = Mid(nRes, 1, iStr)
    Else
        DivA = CInt(nRes)
    End If

End Function
