VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TransferModeConstants
    tmBinary
    tmASCII
End Enum

Public ExtraHandle As Long

Dim Buffer As String

Dim privTransferMode As TransferModeConstants
Dim DataLength As Long 'Binary transfer mode
Public Terminator As String 'ASCII transfer mode

Dim WithEvents Winsock As Winsock
Attribute Winsock.VB_VarHelpID = -1

Event Connect()
Event Disconnect()
Event ConnectionRequest(Connection As Connection)
Event Error(Number As Integer, Description As String)
Event ReceiveData(Data As String)

Property Get RemoteHost() As String
    'On Error Resume Next
    
    RemoteHost = Winsock.RemoteHost
End Property

Property Let RemoteHost(Host As String)
    'On Error Resume Next
    
    Winsock.RemoteHost = Host
End Property

Property Get RemoteHostIP() As String
    'On Error Resume Next
    
    RemoteHostIP = Winsock.RemoteHostIP
End Property

Property Get RemotePort() As Integer
    'On Error Resume Next
    
    RemotePort = Winsock.RemotePort
End Property

Property Let RemotePort(Port As Integer)
    'On Error Resume Next
    
    Winsock.RemotePort = Port
End Property

Property Get LocalHost() As String
    'On Error Resume Next
    
    LocalHost = Winsock.LocalHostName
End Property

Property Get LocalHostIP() As String
    'On Error Resume Next
    
    LocalHostIP = Winsock.LocalIP
End Property

Property Get LocalPort() As Integer
    'On Error Resume Next
    
    LocalPort = Winsock.LocalPort
End Property

Property Let LocalPort(Port As Integer)
    'On Error Resume Next
    
    Winsock.LocalPort = Port
End Property

Property Get TransferMode() As TransferModeConstants
    'On Error Resume Next
    
    TransferMode = privTransferMode
End Property

Property Let TransferMode(TransferMode As TransferModeConstants)
    'On Error Resume Next
    
    If TransferMode = tmBinary Then DataLength = 0&
    
    privTransferMode = TransferMode
End Property

Property Get Protocol() As ProtocolConstants
    'On Error Resume Next
    
    Protocol = Winsock.Protocol
End Property

Property Let Protocol(Protocol As ProtocolConstants)
    'On Error Resume Next
    
    Winsock.Protocol = Protocol
End Property

Property Get IsOpen() As Boolean
    'On Error Resume Next
    
    Select Case Winsock.Protocol
        Case sckTCPProtocol: IsOpen = Winsock.State = sckConnected
        Case sckUDPProtocol: IsOpen = Winsock.State = sckOpen
    End Select
End Property

Property Get State() As Integer
    'On Error Resume Next
    
    State = Winsock.State
End Property

Sub Accept(requestID As Long)
    'On Error Resume Next
    
    Winsock.Close
    Winsock.Accept requestID
    RaiseEvent Connect
End Sub

Sub ClearBuffer()
    'On Error Resume Next
    
    Buffer = Empty
End Sub

Sub SendData(Data As String)
    'On Error Resume Next
    
    If Winsock.State = sckConnected Then Winsock.SendData DWordToString(Len(Data)) & Data
End Sub

Sub Connect(Optional Host As String, Optional Port As Integer)
    'On Error Resume Next
    
    Winsock.Close
    
    If Not IsNullString(Host) Then Winsock.RemoteHost = Host
    If Port <> 0 Then Winsock.RemotePort = Port
            
    Select Case Winsock.Protocol
        Case sckTCPProtocol: Winsock.Connect
        Case sckUDPProtocol: Winsock.Bind
    End Select
End Sub

Sub Listen(Optional Port As Integer)
    'On Error Resume Next
    
    Winsock.Close
    
    If Port <> 0 Then Winsock.LocalPort = Port
    
    Select Case Winsock.Protocol
        Case sckTCPProtocol: Winsock.Listen
        Case sckUDPProtocol: Winsock.Bind
    End Select
End Sub

Sub Disconnect()
    'On Error Resume Next
    
    If Winsock.State <> sckClosed Then
        Winsock.Close
        RaiseEvent Disconnect
    End If
End Sub

Private Sub Class_Initialize()
    'On Error Resume Next
    
    Set Winsock = New MSWinsockLib.Winsock
End Sub

Private Sub Class_Terminate()
    'On Error Resume Next
    
    Set Winsock = Nothing
End Sub

Private Sub Winsock_Close()
    'On Error Resume Next
    
    Disconnect
End Sub

Private Sub Winsock_Connect()
    'On Error Resume Next
    
    RaiseEvent Connect
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
    'On Error Resume Next
    
    Dim Connection As Connection
    
    RaiseEvent ConnectionRequest(Connection)
    
    If Connection Is Nothing Then
        Accept requestID
    Else
        Connection.Accept requestID
    End If
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    'On Error Resume Next
    
    Dim Data As String
    
    Winsock.GetData Data
    Buffer = Buffer & Data
    
    Select Case privTransferMode
        Case tmBinary
            Do
                If DataLength = 0 Then
                    If Len(Buffer) >= 4 Then
                        DataLength = StringToDWord(Left$(Buffer, 4))
                        Buffer = Mid$(Buffer, 5)
                    Else
                        'Wait for more data
                        Exit Do
                    End If
                Else
                    If Len(Buffer) >= DataLength Then
                        Data = Left$(Buffer, DataLength)
                        RaiseEvent ReceiveData(Data)
                        
                        Buffer = Mid$(Buffer, DataLength + 1)
                        DataLength = 0
                    Else
                        'Wait for more data
                        Exit Do
                    End If
                End If
            Loop
        Case tmASCII
            Dim Start As Long
            
            Start = InStr(1, Buffer, Terminator)
            Do While Start
                Data = Left$(Buffer, Start - 1)
                RaiseEvent ReceiveData(Data)
                
                Buffer = Mid$(Buffer, Start + Len(Terminator))
                Start = InStr(1, Buffer, Terminator)
            Loop
    End Select
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'On Error Resume Next
    
    RaiseEvent Error(Number, Description)
    Disconnect
End Sub
