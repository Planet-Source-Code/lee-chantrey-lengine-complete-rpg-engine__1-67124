VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   3  'Windows Default
   Begin Project1.usrBattle Battle 
      Height          =   3735
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      _extentx        =   7011
      _extenty        =   6588
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    usrMenu1.GoDown
End Sub

Private Sub Form_Load()
    Battle.CreateEnemy "Evil Eye"
End Sub
