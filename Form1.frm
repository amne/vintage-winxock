VERSION 5.00
Object = "{B8C203BC-8274-47EC-96A6-08C6528AB238}#12.0#0"; "WinXock.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin WinXock.aWinXock aWinXock1 
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tListItem
 sName As String
 sIdx As Long
 End Type
Dim sockX() As tListItem
Dim nItems As Long



Private Sub aWinXock1_ConnxStateChange(ByVal Index As Integer, oldState As Long, newState As Long)
If Index > nItems Then
 ReDim Preserve sockX(Index) As tListItem
 nItems = Index
 End If
sockX(Index).sName = "Socket " & Index & " named: " & aWinXock1.getxockmember(Index, sName)
If Index > List1.ListCount Then
 List1.AddItem sockX(Index).sName
 sockX(Index).sIdx = List1.ListCount - 1
 Else
 List1.List(sockX(Index).sIdx) = sockX(Index).sName
 If Index = List1.ListIndex Then List1_Click
 End If
End Sub

Private Sub Command1_Click()
Dim i&
aWinXock1.SetRefreshRate 100
'List1.Clear
'aWinXock1.ScanForDeadSock
'For i = 1 To aWinXock1.GetSockUBound
' List1.AddItem "Socket " & i & " named: " & aWinXock1.getxockmember(i, sname)
' Next i
End Sub

Private Sub Form_Load()
aWinXock1.StartServer 2050
End Sub

Private Sub List1_Click()
List2.Clear
List2.AddItem "Name: " & aWinXock1.getxockmember(List1.ListIndex + 1, sName)
List2.AddItem "Index: " & aWinXock1.getxockmember(List1.ListIndex + 1, sIndex)
List2.AddItem "State: " & aWinXock1.getxockmember(List1.ListIndex + 1, sState)
List2.AddItem "Tag: " & aWinXock1.getxockmember(List1.ListIndex + 1, sTag)
End Sub
