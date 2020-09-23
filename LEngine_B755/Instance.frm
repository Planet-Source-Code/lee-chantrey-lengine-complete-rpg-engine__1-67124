VERSION 5.00
Begin VB.Form Instance 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrev 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "*"
      Top             =   180
      Width           =   4395
   End
End
Attribute VB_Name = "Instance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim myUniqueID As String, hPrevInst As Long, lPropValue As Long
    
    Me.Hide
    
    myUniqueID = "Lee Matthew Chantrey" & App.Title ' << use something truly unique for each project.
    
    lPropValue = txtPrev.hWnd ' << change name as needed
    
    hPrevInst = IsPrevInstance(Me.hWnd, myUniqueID, lPropValue, True)
    If Not hPrevInst = 0 Then
        End ' don't show this other instance
    End If
    
    ' Change last parameter of IsPrevInstance() to FALSE
    ' if you you want to control what is being passed. But pass something,
    ' otherwise your 1st instance won't be aware it needs to show itself
    Main

End Sub

Private Sub txtPrev_Change()

    If txtPrev.Text = "*" Then
        Exit Sub
    End If
    
    txtPrev.Text = Trim(txtPrev.Text)
    NotifyUser "External Command:: '" & txtPrev.Text & "'"
    
    If txtPrev.Text = "" Then
        frmMain.Start "main.sty"
    ElseIf txtPrev.Text = "temp.run" Then
        frmMain.Start "temp.run"
    End If

    txtPrev.Text = "*"

End Sub
