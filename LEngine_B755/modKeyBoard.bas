Attribute VB_Name = "modKeyBoard"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Integer

Public Control_Select As Integer
Public Control_Cancel As Integer

Public Control_Up As Integer
Public Control_Down As Integer
Public Control_Left As Integer
Public Control_Right As Integer

Public Const Control_Equip1 As Integer = 112
Public Const Control_Equip2 As Integer = 113
Public Const Control_Equip3 As Integer = 114
Public Const Control_Equip4 As Integer = 115

Public User_State As Integer

Sub SetJoypad()

    Control_Select = 17
    Control_Cancel = 16
    
    Control_Up = 38
    Control_Down = 40
    Control_Left = 37
    Control_Right = 39

End Sub
