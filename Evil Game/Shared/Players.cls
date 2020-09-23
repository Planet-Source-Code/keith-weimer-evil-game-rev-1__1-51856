VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Players"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Players() As Player

'--------------------------------------------------------------------------------
'PROPERTIES
'--------------------------------------------------------------------------------

Property Get FromCoord(X As Integer, Y As Integer) As Long
    'On Error Resume Next
    
    Dim Index As Long
    
    For Index = 1 To Count
        If Players(Index).X = X And Players(Index).Y = Y Then
            FromCoord = Index
            Exit For
        End If
    Next Index
End Property

Property Get Item(Index As Long) As Player
Attribute Item.VB_UserMemId = 0
    'On Error Resume Next
    
    Set Item = Players(Index)
End Property

Property Set Item(Index As Long, Item As Player)
    'On Error Resume Next
    
    Set Players(Index) = Item
End Property

Property Get Count() As Long
    On Error Resume Next
    
    Count = UBound(Players) - LBound(Players) + 1
End Property

Property Get OpenCount() As Long
    'On Error Resume Next
    
    Dim Index As Long
    
    For Index = 1 To Count
        If Not Players(Index) Is Nothing Then OpenCount = OpenCount + 1
    Next Index
End Property

'--------------------------------------------------------------------------------
'METHODS
'--------------------------------------------------------------------------------

Function Add(Player As Player) As Long
    'On Error Resume Next
    
    Dim Index As Long
        
    For Index = 1 To Count
        If Players(Index) Is Nothing Then
            Add = Index
            Set Players(Index) = Player
            Exit Function
        End If
    Next Index
    
    ReDim Preserve Players(1 To Count + 1)
    Add = Count
    Set Players(Count) = Player
End Function

Sub Remove(Index As Long)
    'On Error Resume Next
    
    Set Players(Index) = Nothing
    CleanUp
End Sub

Sub CleanUp()
    'On Error Resume Next
    
    Dim Index As Long
    Dim Lowest As Long
    
    For Index = Count To 1 Step -1
        If Players(Index) Is Nothing Then
            Lowest = Index
        Else
            If Players(Index) Is Nothing Then
                Lowest = Index
            Else
                If Players(Index).ConnectionHandle = 0 Then
                    Set Players(Index) = Nothing
                    Lowest = Index
                Else
                    Exit For
                End If
            End If
        End If
    Next Index
    
    If Lowest <> 0 Then
        If Lowest = 1 Then
            Erase Players
        Else
            ReDim Preserve Players(1 To Lowest)
        End If
    End If
End Sub
