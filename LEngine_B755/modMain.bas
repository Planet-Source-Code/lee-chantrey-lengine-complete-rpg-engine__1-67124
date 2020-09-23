Attribute VB_Name = "modMain"
Option Explicit

Public bUnload As Boolean

Sub Main()

Dim sCmd As String
    sCmd = Command

    On Error Resume Next

    Call SetVars
    
    DebugWin.Show
    VarsWin.Show

    DoEvents

    'Free file
    StoryCursor = FreeFile()

    NotInDbg = True
    
    sCmd = Replace(LCase(sCmd), LCase(App.Path & "\resources\story\"), "")
    
    frmMain.Show
    
    If sCmd = "" Then
        sCmd = "main.sty"
    End If
    
    LoadDefaults
    
    NotifyUser "Complete! Starting '" & sCmd & "'"
    NotifyUser ""

    frmMain.Start CStr(sCmd)

End Sub

Sub LoadDefaults()

Dim sMusic(3) As String

    sMusic(0) = sPath_Battle & "music"
    sMusic(1) = sPath_Battle & "music"
    sMusic(2) = sPath_Battle & "music"

    'Load Default Music
    If FindExistingMusic(sMusic(0), "boss") = True Then
        NotifyUser "Preloading Boss Music"
        modMCI.LoadFile sMusic(0), "boss"
    End If
    
    If FindExistingMusic(sMusic(1), "victory") = True Then
        NotifyUser "Preloading Victory Music"
        modMCI.LoadFile sMusic(1), "victory"
    End If
    
    If FindExistingMusic(sMusic(2), "fiend") = True Then
        NotifyUser "Preloading Fiend Music"
        modMCI.LoadFile sMusic(2), "fiend"
    End If

End Sub
