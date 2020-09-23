VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Messages Window"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4515
   ScaleWidth      =   7710
   Begin VB.TextBox txtCur 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   6015
   End
   Begin VB.ListBox lstOut 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2460
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bScrollLast As Boolean

Public Property Let ScrollLast(bNew As Boolean)
    bScrollLast = bNew
End Property

Sub AddItem(sItem As String)
    lstOut.AddItem sItem
    
    If bScrollLast = True Then
        lstOut.Selected(lstOut.ListCount - 1) = True
    End If
End Sub

Private Sub Form_Load()
    lstOut.AddItem "LEngine By Lee Matthew Chantrey, [Build " & CInt((App.Major * 100) + (App.Minor * 10) + (App.Revision)) & "]"
    lstOut.AddItem "Messages Window v2 (Debug)"
    lstOut.AddItem " "
    
    bScrollLast = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lstOut.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - txtCur.Height
    txtCur.Move 0, Me.ScaleHeight - txtCur.Height, Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bUnload = True Then
        Exit Sub
    End If
    
    Cancel = True
    'Just hide window instead
    Me.Hide
End Sub

Private Sub lstOut_DblClick()
    Clipboard.Clear
    Clipboard.SetText lstOut.Text

    MsgBox "Copied to clipboard"
End Sub

Private Sub txtCur_DblClick()
    STYDebug
End Sub

Private Sub txtCur_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtCur.ForeColor = vbBlack
        cStory_ToDo.Add txtCur.Text
        
        ResumeStory
    End If
        
End Sub
