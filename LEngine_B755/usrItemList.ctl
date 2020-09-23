VERSION 5.00
Begin VB.UserControl usrItemList 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   4740
      Picture         =   "usrItemList.ctx":0000
      ScaleHeight     =   3600
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   0
      Width           =   60
   End
   Begin VB.Image imgCur 
      Height          =   240
      Left            =   0
      Picture         =   "usrItemList.ctx":0276
      Top             =   30
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "usrItemList.ctx":05EC
      Top             =   0
      Width           =   60
   End
   Begin VB.Image imgback 
      Height          =   720
      Left            =   0
      Picture         =   "usrItemList.ctx":086E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30
   End
   Begin VB.Label Items 
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "usrItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private outCur As Image

Private oIndex As Integer
Private hIndex As Integer
Private oiScale As Integer

Private bFirst As Boolean

Public Function HideCursor()
    imgCur.Visible = False
    outCur.Visible = False
End Function

Private Function AllItemsGone() As Boolean

    AllItemsGone = True

    Dim I As Integer
    For I = 0 To Items.UBound
        If Items(I).Caption <> "" Then
            AllItemsGone = False
            Exit Function
        End If
    Next

End Function

Public Function HighlightNext() As Boolean

    HighlightNext = True

    Dim I As Integer
    For I = (hIndex + 1) To Items.UBound
        If Items(I).Caption <> "" Then
            Debug.Print Items(I).Caption
        
            Me.Highlight I

            Exit Function
        End If
    Next
    
    HighlightNext = False

End Function

Public Function FindByName(sItem As String)
    
    Dim I As Integer
    
    For I = 0 To Items.UBound
        If Items(I).Caption = sItem Then
            FindByName = I
            Exit Function
        End If
    Next
    
    FindByName = 999
    
End Function

Property Get HListIndex() As Integer
    HListIndex = hIndex
End Property

Property Let HListIndex(NewIndex As Integer)
    hIndex = NewIndex
End Property

Property Get ListIndex() As Integer
    ListIndex = oIndex
End Property

Property Let ListIndex(NewIndex As Integer)
    oIndex = NewIndex
End Property

Property Get ListCount() As Integer
    ListCount = Items.UBound
End Property

Property Get SelectedHTag() As String
    SelectedHTag = Items(hIndex).Tag
End Property

Property Get SelectedCaption() As String
    'Same as GetSelected, just doesnt change anything
    SelectedCaption = Items(oIndex).Caption
End Property

Sub ClearList()

    Dim I As Integer
    For I = 1 To Items.UBound
        Unload Items(I)
    Next
    
    bFirst = False
    
    'Clear Variables
    oIndex = 0
    oiScale = 0
    hIndex = 0
    
    Items(0).Tag = ""
    Items(0).Caption = "Nothing"

End Sub

Sub AttachCur(ByRef Cur As Image, myLeft As Integer, myTop As Integer, Optional iScale As Integer = 15)
    Set outCur = Cur
    
    'imgCur.Top = 2
    'oIndex = 0
    
    imgCur.Visible = True
    outCur.Left = myLeft - (outCur.Width - (imgCur.Width * iScale))
    outCur.Top = myTop + (imgCur.Top * iScale)
    outCur.Visible = True
    
    oiScale = iScale
    
    If Items(oIndex).Caption = "" Then
        FindNext
    End If
        
End Sub

Sub GoDown()
    
    If (oIndex + 1) > Items.UBound Then
        'Bring cursor to top
        imgCur.Top = imgCur.Top - (10 * Items.UBound)
        outCur.Top = outCur.Top - 10 * (oiScale * Items.UBound)
        
        oIndex = 0
        
        If AllItemsGone = False Then
            While Items(oIndex).Caption = ""
                GoDown
            Wend
        End If
        
        Exit Sub
    End If
    
    imgCur.Top = imgCur.Top + 10
    outCur.Top = outCur.Top + 10 * oiScale
    
    oIndex = oIndex + 1
    
    If AllItemsGone = False Then
        While Items(oIndex).Caption = ""
            GoDown
        Wend
    End If

End Sub

Sub GoUp()
  
    If (oIndex - 1) < 0 Then
        'Bring cursor to bottom
        imgCur.Top = imgCur.Top + (10 * Items.UBound)
        outCur.Top = outCur.Top + 10 * (oiScale * Items.UBound)
        
        oIndex = Items.UBound
        
        If AllItemsGone = False Then
            While Items(oIndex).Caption = ""
                GoUp
            Wend
        End If
        
        Exit Sub
    End If
 
    imgCur.Top = imgCur.Top - 10
    outCur.Top = outCur.Top - 10 * oiScale
    
    oIndex = oIndex - 1
    
    If AllItemsGone = False Then
        While Items(oIndex).Caption = ""
            GoUp
        Wend
    End If

End Sub

Sub ChangeTag(iIndex As Integer, sTag As String)

    Items(iIndex).Tag = sTag

End Sub

Sub ChangeCaption(iIndex As Integer, sCaption As String)

    Items(iIndex).Tag = sCaption

End Sub

Sub RemoveItem(Index As Integer)

    With Items(Index)
        .Caption = ""
    End With
    
    If oIndex = Index Then 'cursor on item about to remove ?
        FindNext
    End If

End Sub

Private Sub FindNext(Optional bDirUp As Boolean = False)

    'Routines will recursivly call each other

    If bDirUp = False Then
        GoDown
    Else
        GoUp
    End If

End Sub

Sub AddItem(sItem As String, iQty As Integer)
    If bFirst = False Then
        bFirst = True
        
        Items(0).Caption = sItem
        Items(0).Tag = iQty
        Debug.Print iQty
        
        Exit Sub
    End If
    
    Load Items(Items.Count)
    With Items(Items.UBound)
        .Top = Items(.Index - 1).Top + 10
    
        .Caption = sItem & " x" & iQty
        .Tag = iQty
        
        .Visible = True
        .ZOrder 0
    End With
    
End Sub

Sub UpdateItem(sItem As String, iQty As Integer)

    

End Sub

Sub Highlight(Index As Integer)

    Dim I As Integer
    For I = 0 To Items.UBound
        Items(I).ForeColor = vbWhite
    Next
    
    Items(Index).ForeColor = vbYellow
    hIndex = Index

End Sub

Sub Nullight(Index As Integer)

    Items(Index).ForeColor = vbWhite

End Sub

Function GetSelected() As String
    
    GetSelected = Items(oIndex).Tag
    
    imgCur.Visible = False
    outCur.Visible = False
    
End Function

Private Sub UserControl_Resize()
    imgback.Width = UserControl.Width
    Items(0).Width = UserControl.ScaleWidth
End Sub

Property Get Tag()
    Tag = UserControl.Tag
End Property
