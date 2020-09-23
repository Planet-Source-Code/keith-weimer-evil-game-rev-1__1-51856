VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Connections() As ConnectionEvents

Event Connect(Index As Long)
Event Disconnect(Index As Long)
Event ConnectionRequest(Index As Long, Connection As Connection)
Event Error(Index As Long, Number As Integer, Description As String)
Event ReceiveData(Index As Long, Data As String)

'--------------------------------------------------------------------------------
'EVENT RELAYS
'--------------------------------------------------------------------------------

Sub Connect(Index As Long)
    'On Error Resume Next
    
    RaiseEvent Connect(Index)
End Sub

Sub Disconnect(Index As Long)
    'On Error Resume Next
    
    RaiseEvent Disconnect(Index)
    Remove Index
End Sub

Sub ConnectionRequest(Index As Long, Connection As Connection)
    'On Error Resume Next
    
    RaiseEvent ConnectionRequest(Index, Connection)
End Sub

Sub Error(Index As Long, Number As Integer, Description As String)
    'On Error Resume Next
    
    RaiseEvent Error(Index, Number, Description)
End Sub

Sub ReceiveData(Index As Long, Data As String)
    'On Error Resume Next
    
    RaiseEvent ReceiveData(Index, Data)
End Sub

'--------------------------------------------------------------------------------
'PROPERTIES
'--------------------------------------------------------------------------------

Property Get Item(Index As Long) As Connection
Attribute Item.VB_UserMemId = 0
    'On Error Resume Next
    
    Set Item = Connections(Index).Connection
End Property

Property Set Item(Index As Long, Item As Connection)
    'On Error Resume Next
    
    Set Connections(Index).Connection = Item
End Property

Property Get Count() As Long
    On Error Resume Next
    
    Count = UBound(Connections) - LBound(Connections) + 1
End Property

Property Get OpenCount() As Long
    'On Error Resume Next
    
    Dim Index As Long
    
    For Index = 1 To Count
        If Not Connections(Index) Is Nothing Then OpenCount = OpenCount + 1
    Next Index
End Property

'--------------------------------------------------------------------------------
'METHODS
'--------------------------------------------------------------------------------

Function Add(Connection As Connection) As Long
    'On Error Resume Next
    
    Dim Index As Long
    
    Dim ConnectionEvents As ConnectionEvents
    Set ConnectionEvents = New ConnectionEvents
    
    With ConnectionEvents
        Set .Owner = Me
        Set .Connection = Connection
    End With
    
    For Index = 1 To Count
        If Connections(Index) Is Nothing Then
            ConnectionEvents.Index = Index
            Add = Index
            Set Connections(Index) = ConnectionEvents
            Exit Function
        End If
    Next Index
    
    ReDim Preserve Connections(1 To Count + 1)
    ConnectionEvents.Index = Count
    Add = Count
    Set Connections(Count) = ConnectionEvents
End Function

Sub Remove(Index As Long)
    'On Error Resume Next
    
    Set Connections(Index) = Nothing
    CleanUp
End Sub

Sub CleanUp()
    'On Error Resume Next
    
    Dim Index As Long
    Dim Lowest As Long
    
    For Index = Count To 1 Step -1
        If Connections(Index) Is Nothing Then
            Lowest = Index
        Else
            If Connections(Index).Connection Is Nothing Then
                Set Connections(Index) = Nothing
                Lowest = Index
            Else
                If Connections(Index).Connection.IsOpen Then
                    Exit For
                Else
                    Set Connections(Index) = Nothing
                    Lowest = Index
                End If
            End If
        End If
    Next Index
    
    If Lowest <> 0 Then
        If Lowest = 1 Then
            Erase Connections
        Else
            ReDim Preserve Connections(1 To Lowest)
        End If
    End If
End Sub
