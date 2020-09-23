VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sPath As String
Private sSection As String
Private sKey As String
Private vDefault As String

Public Property Let File(sNewPath As String)
    sPath = sNewPath
End Property

Public Property Get File() As String
    'might be usefull someday
    File = sPath
End Property

Public Property Let Section(sNewSection As String)
    sSection = sNewSection
End Property

Public Property Get Section() As String
    'might be usefull someday
    Section = sSection
End Property

Public Property Let Default(sNewDefault As String)
    vDefault = sNewDefault
End Property

Public Property Get Default() As String
    'might be usefull someday
    Default = vDefault
End Property

Public Property Let Key(sNewKey As String)
    sKey = sNewKey
End Property

Public Property Get Key() As String
    'might be usefull someday
    Key = sKey
End Property

Public Function Read(Optional Key As String, Optional Section As String, Optional Path As String, Optional Default As String = vbNullString) As String
    If Key <> "" Then
        sKey = Key
    End If
    
    If Section <> "" Then
        sSection = Section
    End If
    
    If Path <> "" Then
        sPath = Path
    End If
    
    If Default <> "" Then
        vDefault = Default
    End If
    
    Read = ReadINIValue(sPath, sSection, sKey, vDefault)
End Function

Public Function WriteData(Value As String, Optional Key As String, Optional Section As String, Optional Path As String) As String
    If Key <> "" Then
        sKey = Key
    End If
    
    If Section <> "" Then
        sSection = Section
    End If
    
    If Path <> "" Then
        sPath = Path
    End If
    
    If Default <> "" Then
        vDefault = Default
    End If
    
    'Read = ReadINIValue(sPath, sSection, sKey, vDefault)
    WriteINIValue sPath, sSection, sKey, KillHome(Value)
End Function

