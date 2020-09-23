Attribute VB_Name = "modMain"
Option Explicit

Public Locked_Brush As Boolean
Public Locked_Info As Boolean
Public Locked_TilePicker As Boolean

Public Map As Map 'Active map
Public Region As Region 'Brush

Sub LockAll()
    'On Error Resume Next
    
    Locked_Brush = True
    Locked_Info = True
    Locked_TilePicker = True
End Sub

Sub UnlockAll()
    'On Error Resume Next
    
    Locked_Brush = False
    Locked_Info = False
    Locked_TilePicker = False
End Sub
