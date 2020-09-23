VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form2"
   ScaleHeight     =   5700
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin prjLEngine.usrMainMenu usrMainMenu1 
      Height          =   3495
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   4815
      _extentx        =   8493
      _extenty        =   6165
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&First"
      Height          =   315
      Left            =   3660
      TabIndex        =   6
      Top             =   1980
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   915
      Left            =   4920
      TabIndex        =   5
      Top             =   300
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&End"
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   1980
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   4140
      Width           =   735
   End
   Begin VB.Image imgCur 
      Height          =   195
      Left            =   0
      Picture         =   "frmtest.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    usrMenu1.GoDown
End Sub

Private Sub Command2_Click()
    usrMenu1.GoUp
End Sub

Private Sub Command3_Click()
    usrMenu1.RemoveItem usrMenu1.ListIndex
End Sub

Private Sub Command4_Click()
    MsgBox usrMenu1.GetSelectedText
End Sub

Private Sub Command5_Click()
    UC.Visible = False
    UC2.Visible = True
    
    UC.Enabled = False
    UC2.Enabled = True
End Sub

Private Sub Command7_Click()
    UC2.Visible = False
    UC.Visible = True
    
    UC2.Enabled = False
    UC.Enabled = True
End Sub

Private Sub USR_OnFinishedGrow()
    USR.Shrink
End Sub

