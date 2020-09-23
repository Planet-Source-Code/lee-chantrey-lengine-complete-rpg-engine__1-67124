VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iFile As Integer
Private iPos As Long
Private sPath As String

Public Property Get Path() As String
    Path = sPath
End Property

Public Property Get Position() As Long
    Position = iPos
End Property

Public Function SkipTo(iLine As Integer)
    
    While iPos < iLine
        ReadLine
    Wend
    
End Function

Public Function OpenStream(sNewPath As String) As Boolean

    On Error GoTo Catch_E
    OpenStream = True

    iPos = 0
    sPath = sNewPath

    Close #iFile
    Open sPath For Input As #iFile
    
    Exit Function
    
Catch_E:
    OpenStream = False
    WarnUser "FileObject::OpenStream {" & sNewPath & "} " & Err.Description
    
End Function

Public Function ReadLine() As String
    
    Line Input #iFile, ReadLine
    iPos = iPos + 1
        
End Function

Public Function EOS() As Boolean
    
    'Pulling data from STY Writor, then
    'never ending story
    On Error GoTo Catch_E
    
    EOS = True
    EOS = EOF(iFile)
    
    Exit Function
    
Catch_E:
    'End of file (if it doesnt exist / couldnt access it)
    EOS = True
        
End Function

Private Sub Class_Initialize()
    iFile = FreeFile
End Sub
