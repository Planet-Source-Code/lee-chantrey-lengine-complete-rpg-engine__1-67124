VERSION 5.00
Begin VB.Form frmVariables 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Variables #0"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstVars 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CurToEnd()
    lstVars.Selected(lstVars.ListCount - 1) = True
    
    Me.Caption = "Variables #" & lstVars.ListCount
End Sub

Sub Add(ByVal sKey As String, ByVal sValue As String)
    
Dim Index As Integer

    sKey = LCase(sKey)
    Index = getIndex(sKey)

    If Index > -1 Then
        lstVars.RemoveItem (Index)
        lstVars.AddItem sKey & " : " & sValue
    Else
        lstVars.AddItem sKey & " : " & sValue
    End If
        
    CurToEnd
End Sub

Private Function getIndex(ByVal Index As String) As Integer

Dim sP() As String, I As Integer

    If IsNumeric(Index) = False Then
        Index = LCase(Index)
    End If

    For I = 0 To lstVars.ListCount - 1
        sP = Split(lstVars.List(I), " : ", 2)
        
        If sP(0) = Index Then
            getIndex = I
            Exit Function
        End If
    Next

    getIndex = -1

    'MsgBox "Unexpected Error #1:Gui_Variables:getIndex:Out of bounds", vbCritical

End Function

Sub Remove(Index)

Dim I As Integer
    I = getIndex(Index)
    
    If I = -1 Then
        Exit Sub
    End If
    
    lstVars.RemoveItem I
    CurToEnd

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstVars.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
