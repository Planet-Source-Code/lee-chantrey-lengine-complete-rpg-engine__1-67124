Attribute VB_Name = "modStory"
Option Explicit

Private bResume As Boolean

Public Script_Halt As Boolean
Public Player_Control As Boolean
Public StoryCursor As Integer
Public iLineSkip As Integer
Public iMinusPoint As Integer
Public sCurrentStory As String
Public StoryFile As New clsFileObject

Private bStoryEnd As Boolean

'Holds Function Alias's
Private cPublicFunctions As New Collection
Private cPrivateFunctions As New Collection

'Holds Function Data
Private cFunctions As New Collection

Public cStory_ToDo As New Collection

Function FunctionExists(sFunction As String) As Boolean

    If ColExists(cFunctions, sFunction) Then
        FunctionExists = True
        Exit Function
    End If
    
    FunctionExists = False
    Exit Function

End Function

Sub RetrieveFunctions(ByRef fIni As clsIniObj)

Dim I As Integer, sFunctions() As String

    Set cPrivateFunctions = New Collection
    Set cPublicFunctions = New Collection
    Set cFunctions = New Collection

    fIni.Section = "Story_Private_Functions"
    sFunctions = Split(fIni.Read("Keys"), ",")
    
    For I = 0 To UBound(sFunctions)
        cPrivateFunctions.Add sFunctions(I), sFunctions(I)
        cFunctions.Add ReplaceN(fIni.Read(sFunctions(I))) & vbCrLf, sFunctions(I)
    Next

    fIni.Section = "Story_Public_Functions"
    sFunctions = Split(fIni.Read("Keys"), ",")
    
    For I = 0 To UBound(sFunctions)
        cPublicFunctions.Add sFunctions(I), sFunctions(I)
        cFunctions.Add ReplaceN(fIni.Read(sFunctions(I))) & vbCrLf, sFunctions(I)
    Next

End Sub

Sub DumpFunctions(ByRef fIni As clsIniObj)

Dim I As Integer, sKeys As String

    fIni.Section = "Story_Private_Functions"

    If cPrivateFunctions.Count > 0 Then

        For I = 1 To cPrivateFunctions.Count - 1
            fIni.WriteData cFunctions(cPrivateFunctions(I)), cPrivateFunctions(I)
            sKeys = sKeys & cPrivateFunctions(I) & ","
        Next
        
        If cPrivateFunctions.Count > 0 Then
            fIni.WriteData cFunctions(cPrivateFunctions(cPrivateFunctions.Count)), CStr(cPrivateFunctions(cPrivateFunctions.Count))
            sKeys = sKeys & cPrivateFunctions(cPrivateFunctions.Count)
        End If
        
        fIni.WriteData sKeys, "Keys"
        sKeys = ""
        
    End If
    
    If cPublicFunctions.Count > 0 Then
    
        fIni.Section = "Story_Public_Functions"
        
        For I = 1 To cPublicFunctions.Count - 1
            fIni.WriteData ReplaceNL(cFunctions(cPublicFunctions(I))), cPublicFunctions(I)
            sKeys = sKeys & cPublicFunctions(I) & ","
        Next
        
        If cPublicFunctions.Count > 0 Then
            fIni.WriteData ReplaceNL(cFunctions(cPublicFunctions(cPublicFunctions.Count))), cPublicFunctions(cPublicFunctions.Count)
            sKeys = sKeys & cPublicFunctions(cPublicFunctions.Count)
        End If
        
        fIni.WriteData sKeys, "Keys"

    End If

End Sub

Function ReplaceN(sSrc As String) As String

    ReplaceN = Replace(sSrc, Chr(14), vbCrLf)

End Function

Function ReplaceNL(sSrc As String) As String

    ReplaceNL = Replace(sSrc, vbCrLf, Chr(14))

End Function

Sub Reset()

    HaltStory

    Set cPublicFunctions = New Collection
    Set cPrivateFunctions = New Collection
    Set cFunctions = New Collection

End Sub

Sub HaltStory()
    'Halt Story
    DebugWin.Caption = StrFront(DebugWin.Caption, "#") & "# Halted"
    
    Debug.Print "HaltStory"
    Script_Halt = True
    
End Sub

Sub ResumeStory()
    
    Debug.Print "ResumeStory"
       
    'Resume Story
    If cStory_ToDo.Count = 0 Then
        If Script_Halt = False Or bStoryEnd = True Then
            Debug.Print "Story cannot be resumed"
            Exit Sub
        End If

        DebugWin.Caption = StrFront(DebugWin.Caption, "#") & "# Running [From File]"

        Script_Halt = False
        StartStory
    Else
        'Use cache
        DebugWin.Caption = StrFront(DebugWin.Caption, "#") & "# Running [From Memory]"

        Script_Halt = False
        StartStory , True
    End If
End Sub

Sub ReturnStory()

    If StoryCursor > 1 Then
        StoryCursor = StoryCursor - 1
        
        bResume = True
        StartStory
    Else
        Exit Sub
    End If

End Sub

Sub ExeStoryCmd(sCmd As String)
    
    Debug.Print "ExeStoryCmd: " & sCmd

    If sCmd = "" Then
        Debug.Print "ExeStoryCmd_Exit"
    
        Exit Sub
    End If

    If IsSkip(sCmd) Then
        iLineSkip = iLineSkip + CInt(sCmd)
        
        MsgBox "ExeStoryCmd"
        ResumeStory
    Else
        ReplaceOutBrack sCmd

        If cStory_ToDo.Count > 1 Then
            'Move to front of the line
            cStory_ToDo.Add sCmd, , 1
        Else
            cStory_ToDo.Add sCmd
        End If

        'Force resume
        bResume = True
        StartStory "", True
    End If

End Sub

Sub StartNewStory(sNewStory As String)

    On Error GoTo CatchErr

    If bStoryEnd = False Then
        StoryCursor = StoryCursor + 1
    End If
    
    sNewStory = sPath_Story & sNewStory
    
    StoryFile.OpenStream sNewStory
    DebugWin.Caption = KillHome(sNewStory)
    
    bResume = True
    bStoryEnd = False
    
    StartStory
    
    Exit Sub
    
CatchErr:
    If Err.Number = 53 Then
        WarnUser "StartNewStory Failed: " & sPath_Story & sNewStory & " does not exist"
    Else
        MsgBox "StartNewStory Failed: " & Err.Description & Err.Number
    End If

End Sub

Sub StartStory(Optional sNewStory As String, Optional bUseCache As Boolean = False)

Static bGoto As Boolean, sGoto As String, _
        bFunction As Boolean, sFunction As String
    
'Load Main
Dim NewLine As String, sP() As String, sP2() As String, CArgs As New Collection, _
    I As Integer, sChar As String, sParam As String, fChar As String, sL() As String, sItem As String
    
    'Story Script Flows ONE WAY!
    
    If sNewStory <> "" Then
        'Close previous
        sCurrentStory = sPath_Story & sNewStory

        'Resume must be false
        bStoryEnd = False

        bResume = False
    End If
    
    On Error GoTo CatchErr

    If bResume = False Then
        If sCurrentStory <> "" Then
            Debug.Print "Opening New Stream: " & sCurrentStory
            
            'Remove non public functions
            While cPrivateFunctions.Count > 0
                cFunctions.Remove cPrivateFunctions(1)
                cPrivateFunctions.Remove 1
            Wend
            
            Set cStory_ToDo = New Collection
            
            StoryFile.OpenStream sCurrentStory
            DebugWin.Caption = KillHome(sCurrentStory)
        End If
    Else
        bResume = False
    End If
    
    While Not StoryFile.EOS Or bUseCache = True
StartLine:
    
        If cStory_ToDo.Count > 0 Then
            NewLine = cStory_ToDo(1)
            cStory_ToDo.Remove 1
        Else
            If bUseCache = True Then
                MsgBox "Expected error #2 has occured: Told to use cache, when cache is empty.", vbCritical
                Exit Sub
            End If
        
            NewLine = StoryFile.ReadLine
        End If
        
        NewLine = Trim(NewLine)
        fChar = Left(NewLine, 1)

        'Check for skips
        If iLineSkip > 0 Then
        
            iLineSkip = iLineSkip - 1
            GoTo NextLine
            
        End If
    
         'Coments are annoying * sigh *
        If fChar = "#" Or NewLine = "" Then
            'Ingore
            GoTo NextLine
        End If
    
        'Check for gotos
        If bGoto = True Then
            If NewLine = sGoto Then
                'Weve arrived
                bGoto = False
                'Resume script
                Script_Halt = False
            End If
        
            GoTo NextLine
            
        ElseIf bFunction = True Then
            'Goto's could direction go into a function (unclean)
            If NewLine = "_" Then
                bFunction = False
            Else
                'UpdateCol cFunctions, sFunction, NewLine
                ColAmend cFunctions, sFunction, NewLine & vbCrLf
            End If
            
            GoTo NextLine
        End If
        
        'Do the boring check, la la
        If fChar = ":" Then
            GoTo NextLine
        End If
        
        sP = Split(NewLine, " ")
        
        If fChar = "!" Then
            sFunction = LCase(Mid(sP(0), 2))

            If ColExists(cFunctions, sFunction) = True Or _
                ColExists(cPublicFunctions, sFunction) Then
                
                WarnUser "Function (" & sFunction & " ) exists."
                Exit Sub
            End If
            
            bFunction = True
            
            cFunctions.Add "", sFunction
            cPublicFunctions.Add sFunction, sFunction
              
            GoTo NextLine
        End If
        
        If fChar = "." Then
            sFunction = LCase(Mid(sP(0), 2))
            
            If ColExists(cFunctions, sFunction) = True Or _
                ColExists(cPrivateFunctions, sFunction) Then
                
                WarnUser "Function (" & sFunction & " ) exists."
                Exit Sub
            End If
                
            bFunction = True
            
            cFunctions.Add "", sFunction
            cPrivateFunctions.Add sFunction, sFunction
              
            GoTo NextLine
        End If
        
        'Show on DebugWindow

        If UBound(sP) > 0 Then
            NewLine = Mid(NewLine, InStr(NewLine, " ") + 1)
        Else
            NewLine = ""
        End If
        
    Dim Bq As Boolean
    
        Bq = False
        Set CArgs = New Collection
        
        For I = 1 To Len(NewLine)
            sChar = Mid(NewLine, I, 1)
            
            If sChar = """" Then
                If Bq = False Then
                    Bq = True
                    
                ElseIf Bq = True Then
                    CArgs.Add VarScan(sParam)
                    sParam = ""
                    
                    Bq = False
                End If
                
            ElseIf Bq = True Then
                sParam = sParam & sChar
            Else
                If sChar <> " " Then
                
                    MsgBox "Data found between a parameter, check command parameters" & vbCrLf & _
                           "Line: " & NewLine, vbCritical, "StartStory: Parse Error"
                           
                    Exit Sub
                End If
            End If
        Next

        If LCase(sP(0)) = "goto" Then
            bGoto = True
            sGoto = ":" & CArgs(1)
        ElseIf LCase(sP(0)) = "nextevent" Then
            GoTo StartLine
        Else
        
            If ColExists(cFunctions, "catch:" & sP(0) & ":before") Then
                IncludeFunction "catch:" & sP(0) & ":before"
                StartStory , True

            End If
            'Do after
            
            DebugWin.AddItem sP(0) & " " & NewLine
            DebugWin.txtCur.Text = sP(0) & " " & NewLine

            CallByName frmMain, sP(0), VbMethod, CArgs
            
            If ColExists(cFunctions, "catch:" & sP(0) & ":after") Then
                IncludeFunction "catch:" & sP(0) & ":after"
            End If
        End If
        
        If cStory_ToDo.Count = 0 Then
            'No Cache Left
            bUseCache = False
        End If
        
        Debug.Print "Script_Halt: " & Script_Halt
        
        If Script_Halt = True Then
            bResume = True
            If bGoto = False Then
            
                Debug.Print "Exit: StartStory"
                Exit Sub
            End If
        End If
        
NextLine:
    'Debug.Print "NextLine"
    Wend
    
    DebugWin.Caption = StrFront(DebugWin.Caption, "#") & "# -End of sty file"
    
    bStoryEnd = True
    sCurrentStory = ""

    If bGoto = True Then
        bGoto = False
        
        WarnUser "Section '" & sGoto & "' does not exist."
    End If
    
    Exit Sub
    
CatchErr:
    If Err.Number = 438 Then
        'Check its not a function
        sP(0) = LCase(sP(0))
        
        bResume = True 'Forces story to continue at point of interuption
                       '[A function is an interuption as well]
        
        If ColExists(cFunctions, sP(0)) Then
            For I = 1 To CArgs.Count
                Op_Variable I, CArgs(I)
            Next

            sL = Split(cFunctions(sP(0)), vbCrLf)
            'If UBound(sL) = 0 Then
                'MsgBox sP(0) & " is empty!"
                'MsgBox Len(cFunctions(sP(0)))
            'End If
            
            If cStory_ToDo.Count > 0 Then
                'Move to front of the line
            
                cStory_ToDo.Add sL(0), , 1
                For I = 1 To UBound(sL) - 1
                    cStory_ToDo.Add sL(I), , , I
                    Debug.Print "Adding: " & sL(I)
                Next
            Else
                For I = 0 To UBound(sL) - 1
                    cStory_ToDo.Add sL(I)
                    Debug.Print "_Adding: " & sL(I)
                Next
            End If

            StartStory , True
        Else
        
            WarnUser "'" & sP(0) & "'" & " is not a supported command.", True
        End If
        
    Else
        WarnUser "modStory::StartStory {" & sCurrentStory & "} " & Err.Description
    End If

    Exit Sub

End Sub

Public Sub STYDebug()

    On Error GoTo CatchErr
    
    Shell App.Path & "\STYWritor.exe " & StoryFile.Path & " /" & StoryFile.Position
    
    Exit Sub
CatchErr:

    WarnUser "modStory::STYDebug {" & StoryFile.Path & "} " & Err.Description

End Sub

Private Sub IncludeFunction(sName As String)

    Dim sL() As String, I
    
    sL = Split(cFunctions(sName), vbCrLf)
            
    If cStory_ToDo.Count > 0 Then
        'Move to front of the line
            
        cStory_ToDo.Add sL(0), , 1
        For I = 1 To UBound(sL) - 1
            cStory_ToDo.Add sL(I), , , I
        Next
    Else
        For I = 0 To UBound(sL) - 1
            cStory_ToDo.Add sL(I)
        Next
    End If

End Sub
