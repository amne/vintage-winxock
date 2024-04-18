VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl aWinXock 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   780
   Picture         =   "xckMain.ctx":0000
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   52
   ToolboxBitmap   =   "xckMain.ctx":0C42
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckConnx 
      Index           =   1
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "aWinXock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TXMsg
 mLen As Long
 mSplits As Long
 mNSplit As Long
 mMsg As String
 End Type

Public Enum eWSockMembers
 sBytesReceived = 1
 sIndex = 2
 sLocalHostName = 3
 sLocalIP = 4
 sLocalPort = 5
 sName = 6
 sProtocol = 7
 sRemoteHost = 8
 sRemoteHostIP = 9
 sRemotePort = 10
 sSocketHandle = 11
 sState = 12
 sTag = 13
 End Enum

Private lastStates() As Long
Dim msg_Separator As String * 5

Public Event ConnxDataArrival(ByVal Index As Integer, ByVal dataLen As Long, ByVal sData As String)
Public Event ConnxStateChange(ByVal Index As Integer, oldState As Long, newState As Long)
Public Event ConnxConnected(ByVal Index As Integer)
'Public Event connxDisconnected(byval Index As Integer)
Public Event ConnxSendComplete(ByVal Index As Integer)
Public Event ConnxSendProgress(ByVal Index As Integer, bytesSend As Long, bytesRemaining As Long)
Public Event srvConnectionRequest(ByVal AssignedIndex As Long)
Public Event srvStarted(ByVal srvPort As Long)
Public Event srvShutdown(ByVal srvPort As Long)

Private Sub sckConnx_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim sData$
Dim bTotal&
bTotal = 0
While bTotal < bytesTotal
 sckConnx(Index).GetData sData
 bTotal = bTotal + Len(sData)
 Wend
RaiseEvent ConnxDataArrival(Index, bTotal, sData)
End Sub

Private Sub sckServer_ConnectionRequest(ByVal requestID As Long)
Dim NextSock As Long
On Local Error GoTo DaEnd_ConnReq_sckServer
NextSock = FindFreeSock
If NextSock > sckConnx.UBound Then
 Load sckConnx(NextSock)
 ReDim Preserve lastStates(NextSock) As Long
 lastStates(NextSock) = 0
 End If
CloseWinSock sckConnx(NextSock)
sckConnx(NextSock).Accept requestID
RaiseEvent srvConnectionRequest(NextSock)
GoTo OutOfHere_ConnReq_sckServer
DaEnd_ConnReq_sckServer:
 ErrMsg Err.Number, Err.Description, Err.Source, "sckServer_ConnectionRequest"
OutOfHere_ConnReq_sckServer:
End Sub

Private Sub UserControl_Initialize()
'ShutDownServer
If sckServer.State <> 0 Then sckServer.Close
ReDim Preserve lastStates(1) As Long
lastStates(1) = 0
tmrRefresh.Enabled = False
End Sub

Public Function StartServer(ByVal lPort As Long) As Long
On Local Error GoTo DaEnd_StartServer
If sckServer.State <> 0 Then
 sckServer.Close
 End If
sckServer.LocalPort = lPort
sckServer.Listen
RaiseEvent srvStarted(lPort)
GoTo OutOfHere_StartServer
DaEnd_StartServer:
 ErrMsg Err.Number, Err.Description, Err.Source, "StartServer"
OutOfHere_StartServer:
End Function

Public Function ShutDownServer(Optional ByVal Forced As Boolean = True)
On Local Error GoTo DaEnd_ShutDownServer
If Forced Then
 If sckServer.State <> 0 Then
  sckServer.Close
  End If
 Else
 If (sckServer.State = mswinsocklib.StateConstants.sckListening) Or _
    (sckServer.State = mswinsocklib.StateConstants.sckError) Then
  sckServer.Close
  End If
 End If
RaiseEvent srvShutdown(sckServer.LocalPort)
GoTo OutOfHere_ShutDownServer
DaEnd_ShutDownServer:
 ErrMsg Err.Number, Err.Description, Err.Source, "ShutDownServer"
OutOfHere_ShutDownServer:
End Function

Private Sub ErrMsg(ByVal Num As Long, ByVal Desc As String, ByVal Src As String, ByVal funcName As String)
MsgBox "Error #" & Num & " occured in WinXock::" & funcName & vbCrLf & "Description: " & Desc & vbCrLf & "Source: " & Src, vbCritical, "WinXock Error"
End Sub

Private Function FindFreeSock() As Long
Dim i&
On Local Error GoTo DaEnd_FFS
FindFreeSock = sckConnx.UBound + 1
ScanForDeadSock
For i = 1 To sckConnx.UBound
 With sckConnx(i)
  If Mid(.Tag, 1, 1) = "F" Then
   FindFreeSock = i
   Exit Function
   End If
  End With
 Next i
GoTo OutOfHere_FFS
DaEnd_FFS:
 ErrMsg Err.Number, Err.Description, Err.Source, "FindFreeSock"
OutOfHere_FFS:
End Function

Public Sub ScanForDeadSock()
Dim i&
For i = 1 To sckConnx.UBound
 With sckConnx(i)
  If .State = 0 Then .Tag = "F"
  If .State = 2 Then
   .Close
   .Tag = "F"
   End If
  If .State = StateConstants.sckClosing Then
   .Close
   .Tag = "FC"
   End If
  If .State = StateConstants.sckError Then
   .Close
   .Tag = "FE"
   End If
  If .State = StateConstants.sckConnectionPending Then
   .Close
   .Tag = "FCP"
   End If
  If .State = StateConstants.sckOpen Then
   .Close
   .Tag = "FO"
   End If
  If .State = StateConstants.sckConnected Then .Tag = "C"
  If .State = StateConstants.sckConnecting Then .Tag = "CX"
  If .State = StateConstants.sckResolvingHost Then .Tag = "RHX"
  If .State = StateConstants.sckHostResolved Then .Tag = "RH"
  End With
 Next i
End Sub

Private Sub CloseWinSock(wSock As Winsock)
If wSock.State <> 0 Then wSock.Close
End Sub

Public Function GetXockMember(ByVal iDX As Long, ByVal sMember As eWSockMembers) As String
On Local Error GoTo DaEnd_GetXockMember
If iDX < 1 Then GoTo OutOfHere_GetXockMember
If iDX > sckConnx.UBound Then GoTo OutOfHere_GetXockMember
GetXockMember = GetSockMember(sckConnx(iDX), sMember)
GoTo OutOfHere_GetXockMember
DaEnd_GetXockMember:
 ErrMsg Err.Number, Err.Description, Err.Source, "GetXockMember"
OutOfHere_GetXockMember:
End Function

Public Function GetSockCount&()
GetSockCount = sckConnx.Count
End Function
Public Function GetSockLBound&()
GetSockLBound = sckConnx.LBound
End Function
Public Function GetSockUBound&()
GetSockUBound = sckConnx.UBound
End Function

Private Function GetSockMember(scK As Winsock, ByVal sMember As eWSockMembers) As String
On Local Error GoTo DaEnd_GetSockMember
Select Case sMember
 Case eWSockMembers.sBytesReceived: GetSockMember = scK.BytesReceived
 Case eWSockMembers.sIndex: GetSockMember = scK.Index
 Case eWSockMembers.sLocalHostName: GetSockMember = scK.LocalHostName
 Case eWSockMembers.sLocalIP: GetSockMember = scK.LocalIP
 Case eWSockMembers.sLocalPort: GetSockMember = scK.LocalPort
 Case eWSockMembers.sName: GetSockMember = scK.Name
 Case eWSockMembers.sProtocol: GetSockMember = scK.Protocol
 Case eWSockMembers.sRemoteHost: GetSockMember = scK.RemoteHost
 Case eWSockMembers.sRemoteHostIP: GetSockMember = scK.RemoteHostIP
 Case eWSockMembers.sRemotePort: GetSockMember = scK.RemotePort
 Case eWSockMembers.sSocketHandle: GetSockMember = scK.SocketHandle
 Case eWSockMembers.sState: GetSockMember = scK.State
 Case eWSockMembers.sTag: GetSockMember = scK.Tag
 End Select
GoTo OutOfHere_GetSockMember
DaEnd_GetSockMember:
 ErrMsg Err.Number, Err.Description, Err.Source, "GetSockMember"
OutOfHere_GetSockMember:
End Function

Public Sub ScanForChanges()
Dim i&
For i = 1 To sckConnx.UBound
 If lastStates(i) <> sckConnx(i).State Then
  Select Case sckConnx(i).State
   Case StateConstants.sckClosed: sckConnx(i).Tag = "F"
   Case StateConstants.sckClosing: sckConnx(i).Tag = "FC"
   Case StateConstants.sckConnected: sckConnx(i).Tag = "C"
   Case StateConstants.sckConnecting: sckConnx(i).Tag = "CX"
   Case StateConstants.sckConnectionPending: sckConnx(i).Tag = "FCP"
   Case StateConstants.sckError: sckConnx(i).Tag = "FE"
   Case StateConstants.sckHostResolved: sckConnx(i).Tag = "RH"
   Case StateConstants.sckOpen: sckConnx(i).Tag = "FO"
   Case StateConstants.sckResolvingHost: sckConnx(i).Tag = "RHX"
   End Select
  RaiseEvent ConnxStateChange(i, lastStates(i), sckConnx(i).State)
  lastStates(i) = sckConnx(i).State
  End If
 Next i
End Sub

Public Sub SetRefreshRate(ByVal interval_MS As Integer)
If interval_MS < 0 Then interval_MS = 0
If interval_MS = 0 Then tmrRefresh.Enabled = False
tmrRefresh.Enabled = True
tmrRefresh.Interval = interval_MS
End Sub

Private Sub tmrRefresh_Timer()
ScanForChanges
End Sub


Private Sub UserControl_Resize()
On Local Error Resume Next
UserControl.Width = 550
On Local Error Resume Next
UserControl.Height = 550
End Sub

Private Sub UserControl_Terminate()
tmrRefresh.Enabled = False
End Sub
