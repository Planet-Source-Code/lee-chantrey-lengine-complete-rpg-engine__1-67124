Attribute VB_Name = "modWinini"
Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" _
(ByVal grpnm As String, ByVal parnm As String, _
ByVal deflt As String, ByVal parvl As String, _
ByVal parlen As Long, ByVal INIPath As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" _
(ByVal grpnm As String, ByVal parnm As String, _
ByVal parvl As String, ByVal INIPath As String) As Long

Public Function ReadINIValue(ByVal INIPath As String, _
ByVal SectionName As String, ByVal KeyName As String, _
ByVal DefaultValue As String) As String
Dim sBuff As String
Dim x As Long
sBuff = Space$(1024)
x = GetPrivateProfileString(SectionName, KeyName, DefaultValue, _
sBuff, Len(sBuff), INIPath)
ReadINIValue = Left$(sBuff, x)
End Function

Public Sub WriteINIValue(ByVal INIPath As String, _
ByVal SectionName As String, ByVal KeyName As String, _
ByVal KeyValue As String)
Dim x As Long
x = WritePrivateProfileString(SectionName, KeyName, KeyValue, INIPath)
End Sub

