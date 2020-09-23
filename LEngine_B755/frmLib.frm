VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmLib 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   2385
   ClientTop       =   2340
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSScriptControlCtl.ScriptControl Math 
      Left            =   3360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   2400
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1440
      ScaleHeight     =   615
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   240
      Pattern         =   "clsparty.*.*"
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin prjLEngine.usrNumero usrNumero 
      Left            =   600
      Top             =   960
      _ExtentX        =   1296
      _ExtentY        =   1085
   End
   Begin VB.Image imgCur 
      Height          =   195
      Left            =   0
      Picture         =   "frmLib.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub PlayMusic(sFile As String)

    modMCI.LoadFile sFile, "F"
    modMCI.PlayAudio "F", True

    Debug.Print "Playing!"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    modMCI.CloseAll
    modMCI.Unhook
End Sub

Private Sub Form_Load()
    
Dim I As Integer

    For I = 0 To C_MAX_PLAYERS
        Math.AddObject "Player" & I, pBattleHu(I), False
        Math.AddObject "Enemy" & I, pBattleAI(I), False
    Next
    
    modMCI.HookWindow Me.hWnd
        
End Sub
