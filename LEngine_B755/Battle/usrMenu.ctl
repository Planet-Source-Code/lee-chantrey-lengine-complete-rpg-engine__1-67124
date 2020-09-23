VERSION 5.00
Begin VB.UserControl usrMenu 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   4740
      Picture         =   "usrMenu.ctx":0000
      ScaleHeight     =   3600
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   0
      Width           =   60
   End
   Begin VB.Label Items 
      BackStyle       =   0  'Transparent
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
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   450
      Width           =   735
   End
   Begin VB.Label Number 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Index           =   3
      Left            =   3960
      TabIndex        =   7
      Top             =   450
      Width           =   735
   End
   Begin VB.Label Items 
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   300
      Width           =   735
   End
   Begin VB.Label Number 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   300
      Width           =   735
   End
   Begin VB.Label Items 
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Number 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Number 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgCur 
      Height          =   240
      Left            =   0
      Picture         =   "usrMenu.ctx":0276
      Tag             =   "0"
      Top             =   30
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Items 
      BackStyle       =   0  'Transparent
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
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "usrMenu.ctx":05EC
      Top             =   0
      Width           =   60
   End
   Begin VB.Image imgback 
      Height          =   720
      Left            =   0
      Picture         =   "usrMenu.ctx":086E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30
   End
End
Attribute VB_Name = "usrMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private outCur

Private oIndex As Integer
Private hIndex As Integer
Private oiScale As Integer

Private bFirst As Boolean

Private cItems As New Collection
Private cNumbers As New Collection
Private cTags As New Collection

Private lCount As Integer

Private aIndex As Integer
Private bDown As Boolean

Public Function HideCursor()
    On Error Resume Next

    imgCur.Visible = False
    outCur.Visible = False
End Function

Private Function AllItemsGone() As Boolean

    AllItemsGone = True

    If lCount <> 0 Then
        AllItemsGone = False
    End If
    
End Function

Public Function HighlightNext() As Boolean

    HighlightNext = True

    Dim I As Integer
    For I = (hIndex + 1) To Items.UBound
        If Items(I).Caption <> "" Then

            Me.Highlight I
            Exit Function
        End If
    Next
    
    HighlightNext = False

End Function

Public Function FindByName(sItem As String)
    
    Dim I As Integer
    
    For I = 1 To cItems.Count
        If cItems(I) = sItem Then
            FindByName = I - 1
            Exit Function
        End If
    Next
    
    FindByName = 999
    
End Function

Property Get HListIndex() As Integer
    HListIndex = hIndex
End Property

Property Let HListIndex(newIndex As Integer)
    hIndex = newIndex
End Property

Property Get ListIndex() As Integer
    ListIndex = oIndex
End Property

Property Let ListIndex(newIndex As Integer)
    oIndex = newIndex
End Property

Property Get ListCount() As Integer
    ListCount = cItems.Count - 1
End Property

Property Get SelectedHTag() As String
    SelectedHTag = cTags(hIndex + 1)
End Property

Property Get SelectedCaption() As String
    'Same as GetSelected, just doesnt change anything
    SelectedCaption = cItems(oIndex + 1)
End Property

Sub ClearList()

    Set cTags = New Collection
    Set cItems = New Collection
    Set cNumbers = New Collection

    Dim I As Integer

    'Clear Variables
    oIndex = 0
    oiScale = 0
    hIndex = 0
    lCount = 0
    
    Items(0).Caption = ""
    
    For I = 1 To 3
        Items(I).Caption = ""
    Next
    
    imgCur.Top = 2
    imgCur.Tag = 0

End Sub

Sub AttachCur(ByRef Cur, myLeft As Integer, myTop As Integer, Optional iScale As Integer = 15)
    Set outCur = Cur
    
    'imgCur.Top = 2
    'oIndex = 0
    
    imgCur.Visible = True

    outCur.Left = myLeft - (outCur.Width - (imgCur.Width * iScale))
    outCur.Top = myTop + (imgCur.Top * iScale)
    outCur.Visible = True
    
    oiScale = iScale

End Sub

Sub GoDown()
    
Dim I As Integer

    If (oIndex + 2) > cItems.Count Then
        Exit Sub
    End If

    If imgCur.Top = 32 Then
        'Get The Next Array of items
        oIndex = oIndex + 1
        
        bDown = True
        aIndex = oIndex

        For I = 0 To 3
            Items(I).Caption = cItems(oIndex + I - 2)
            Number(I).Caption = cNumbers(oIndex + I - 2)
        Next
    Else
    
        imgCur.Top = imgCur.Top + 10
        outCur.Top = outCur.Top + 10 * oiScale
        
        imgCur.Tag = imgCur.Tag + 1
        
        oIndex = oIndex + 1
        
    End If
    
    If Items(imgCur.Tag) = "" Then
        GoDown
    End If

End Sub

Sub GoUp()
  
Dim I As Integer

    If (oIndex - 1) < 0 Then
        Debug.Print "Exiting"
    
        Exit Sub
    End If

    If imgCur.Top = 2 Then
        'Get The Next Array of items
        oIndex = oIndex - 1
        
        bDown = False
        aIndex = oIndex
        
        For I = 0 To 3
            Items(I).Caption = cItems(oIndex + I + 1)
            Number(I).Caption = cNumbers(oIndex + I + 1)
        Next
    Else
        imgCur.Top = imgCur.Top - 10
        outCur.Top = outCur.Top - 10 * oiScale
        imgCur.Tag = imgCur.Tag - 1
        
        oIndex = oIndex - 1
        
    End If
    
    If Items(imgCur.Tag) = "" Then
        GoUp
    End If

End Sub

Sub ChangeTag(iIndex As Integer, sTag As String)

    UpdateCol cTags, iIndex + 1, sTag
    Items(iIndex).Tag = sTag

End Sub

Sub ChangeCaption(iIndex As Integer, sCaption As String)

    UpdateCol cItems, iIndex + 1, sCaption
    Items(iIndex).Caption = sCaption

End Sub

Sub RemoveItem(ByVal Index As Integer)

Dim I As Integer

    lCount = lCount - 1
    Index = Index + 1

    UpdateCol cItems, Index, ""
    UpdateCol cTags, Index, ""
    UpdateCol cNumbers, Index, ""

    On Error Resume Next
    'Ignore the fact there might be less then 4 items

    If bDown = False Then
        For I = 0 To 3
            Items(I).Caption = cItems(aIndex + I + 1)
            Number(I).Caption = cNumbers(aIndex + I + 1)
        Next
    Else
        For I = 0 To 3
            Items(I).Caption = cItems(aIndex + I - 2)
            Number(I).Caption = cNumbers(aIndex + I - 2)
        Next
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

Sub AddItem(sItem As String, sTag As String, Optional sNum As String = "")
    
    cItems.Add sItem
    cTags.Add sTag
    cNumbers.Add sNum
    
    If (lCount) < 4 Then
        With Items(lCount)
            .Caption = sItem
            .Visible = True
            .ZOrder 0
            
            .Width = Items(0).Width
        End With
        
        With Number(lCount)
            'Line Up
            .Left = Number(0).Left
        
            .Caption = sNum
            .Visible = True
            .ZOrder 0
        End With
    End If
    
    lCount = lCount + 1
    
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

Function GetSelectedTag() As String
    GetSelectedTag = cTags(oIndex + 1)
End Function

Function GetSelectedText() As String
    If cItems.Count = 0 Then
        Exit Function
    End If
    
    GetSelectedText = cItems(oIndex + 1)
End Function

Function GetSelectedNumber() As Integer
    GetSelectedNumber = cNumbers(oIndex + 1)
End Function

Private Sub UserControl_Resize()
    imgBack.Width = UserControl.Width
    
    Items(0).Width = UserControl.ScaleWidth
    Number(0).Left = UserControl.ScaleWidth - Number(0).Width - 10
End Sub

Property Get Tag()
    Tag = UserControl.Tag
End Property
