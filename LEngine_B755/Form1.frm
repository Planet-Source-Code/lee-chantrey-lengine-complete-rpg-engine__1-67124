VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Unload"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin prjLEngine.usrBattle Battle 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4800
      _extentx        =   8467
      _extenty        =   6376
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Test"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "S.Tippett (ONLY)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Battle_BattleFinished(iRes As Integer)
    MsgBox "done!"
End Sub

Private Sub Command1_Click()

    Battle.CreatePlayer "main.chr"
    Battle.CreatePlayer "alex.chr", True
    
    '--
    
    Battle.LoadBattFile ""
    Battle.Format
    
    Battle.CreateEnemy "evileye.ply"
    
    Battle.StartATMs

End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

