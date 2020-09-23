VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectionEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'God this is so stupid.
'Just a big stupid event relay.

Public Owner As Connections
Public Index As Long
Public WithEvents Connection As Connection
Attribute Connection.VB_VarHelpID = -1

Private Sub Connection_Connect()
    'On Error Resume Next
    
    Owner.Connect Index
End Sub

Private Sub Connection_ConnectionRequest(Connection As Connection)
    'On Error Resume Next
    
    Owner.ConnectionRequest Index, Connection
End Sub

Private Sub Connection_Disconnect()
    'On Error Resume Next
    
    Owner.Disconnect Index
End Sub

Private Sub Connection_Error(Number As Integer, Description As String)
    'On Error Resume Next
    
    Owner.Error Index, Number, Description
End Sub

Private Sub Connection_ReceiveData(Data As String)
    'On Error Resume Next
    
    Owner.ReceiveData Index, Data
End Sub
