Attribute VB_Name = "modMCI"
Option Explicit
' Declarations and such needed for the example:
' (Copy them to the (declarations) section of a module.)
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal _
    lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength _
    As Long, ByVal hwndCallback As Long) As Long
    
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal _
    fdwError As Long, ByVal lpszErrorText As String, ByVal cchErrorText As Long) As Long
    
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal _
    lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    
Private Const MM_MCINOTIFY = &H3B9
Private Const MCI_NOTIFY_ABORTED = &H4
Private Const MCI_NOTIFY_FAILURE = &H8
Private Const MCI_NOTIFY_SUCCESSFUL = &H1
Private Const MCI_NOTIFY_SUPERSEDED = &H2

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal _
lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd _
As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private pOldProc As Long
Private pFHwnd As Long

Private cFiles As New SuperCollection

Private bLoops As Boolean
Private sLastPlayed As String

Public Function DumpAliases(ByRef bcIni As clsIniObj)

    bcIni.Section = "Audio"
    DumpSuperCollection bcIni, cFiles

End Function

Public Function RestoreAliases(ByRef bcIni As clsIniObj)

    'We dont use RestoreSuperCollection, because we need to reload

Dim I As Integer, sKeys() As String

    CloseAll

    bcIni.Section = "Audio"
    sKeys = Split(bcIni.Read("Keys"), ",")
    
    For I = 0 To UBound(sKeys)
        LoadFile sPath_Resources & bcIni.Read(sKeys(I)), sKeys(I)
    Next

End Function

Private Function GetShortPath(LongPath As String) As String

Dim s As String
Dim I As Long
Dim PathLength As Long

        I = Len(LongPath) + 1

        s = String(I, 0)

        PathLength = GetShortPathName(LongPath, s, I)

        GetShortPath = Left$(s, PathLength)

End Function

Public Function HookWindow(lHwnd As Long)

    If InDbg = True Then
        Exit Function
    End If

    pOldProc = SetWindowLong(lHwnd, GWL_WNDPROC, AddressOf WindowProc)
    pFHwnd = lHwnd

End Function

Public Function Unhook()

    SetWindowLong pFHwnd, GWL_WNDPROC, pOldProc

End Function

Public Function LoadFile(ByVal sPath As String, sAlias As String)

Dim errcode As Long, sPathShort As String  ' MCI error code

    If InDbg = True Then
        Exit Function
    End If

    If cFiles.Exists(sAlias) = True Then
        CloseFile sAlias
    End If
    
    sPathShort = GetShortPath(sPath)
    errcode = mciSendString("open " & Qo(sPathShort) & " alias " & Qo(sAlias), "", 0, 0)
    
    If errcode <> 0 Then
        DisplayError errcode, sPath
        WarnUser sPath, True
        
        Exit Function
    End If
    
    cFiles.Add sPath, sAlias

End Function

Public Function PlayAudio(sAlias As String, Optional bLoop As Boolean = False) As Boolean

Dim errcode As Long  ' MCI error code

    If InDbg = True Then
        Exit Function
    End If

    If cFiles.Exists(sAlias) = False Then
        WarnUser "PlayAudio: ID was not found: '" & sAlias & "'", True
    
        PlayAudio = False
        Exit Function
    End If

    Debug.Print "PlayAudio: " & sAlias

    mciSendString "seek " & Qo(sAlias) & " to start", "", 0, 0
    errcode = mciSendString("play " & Qo(sAlias) & " notify", "", 0, pFHwnd)
    
    If errcode <> 0 Then
        DisplayError errcode, ""
        Exit Function
        
    End If
    
    If bLoop = True Then
        sLastPlayed = sAlias
    End If
        
    bLoops = bLoop
        
End Function

Public Function CloseFile(sAlias As String)

    If mciSendString("close " & Qo(sAlias), "", 0, 0) = 0 Then
        cFiles.Remove sAlias
    End If

End Function

Public Function PauseAudio(sAlias As String)

    mciSendString "stop " & Qo(sAlias), "", 0, 0

End Function

Public Function StopAudio(sAlias As String)

    mciSendString "stop " & Qo(sAlias), "", 0, 0
    mciSendString "seek " & Qo(sAlias) & " to start", "", 0, 0

End Function

Public Function CloseAll()

Dim errcode As Long

    While cFiles.Count > 0
        errcode = mciSendString("close " & cFiles.Key(1), "", 0, 0)
        cFiles.Remove 1
    Wend

End Function

Private Sub DisplayError(ByVal errcode As Long, sFileName As String)
    ' This subroutine displays a dialog box with the text of the MCI error.  There's
    ' no reason to use the MessageBox API function; VB's MsgBox function will suffice.
    Dim errstr As String  ' MCI error message text
    Dim retval As Long    ' return value
    
    ' Get a string explaining the MCI error.
    errstr = Space(128)
    retval = mciGetErrorString(errcode, errstr, Len(errstr))
    ' Remove the terminating null and empty space at the end.
    errstr = Left(errstr, InStr(errstr, vbNullChar) - 1)

    retval = WarnUser("MCI Error: " & sFileName & " :" & errstr, True)
End Sub

' Custom window procedure for Form1.
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long
Dim mbtext As String ' text of message box
Dim retval As Long ' return value

    ' If the notification message is received, tell the user how
    ' the playback of the MIDI file concluded.
    If uMsg = MM_MCINOTIFY Then
    
        Select Case wParam
            
        Case MCI_NOTIFY_SUCCESSFUL
            CheckLoopers

        End Select

        WindowProc = 0
            
    Else
            
        retval = CallWindowProc(pOldProc, hWnd, uMsg, wParam, lParam)
        
    End If
    
End Function

Private Function CheckLoopers()

    If bLoops = True Then
        PlayAudio sLastPlayed, True
    End If

End Function

