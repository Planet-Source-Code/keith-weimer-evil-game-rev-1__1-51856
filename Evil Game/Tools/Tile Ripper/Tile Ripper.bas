Attribute VB_Name = "modTileRipper"
Option Explicit

Public LOCKED_INPUT As Boolean

Sub LockAll()
    LOCKED_INPUT = True
End Sub

Sub UnlockAll()
    LOCKED_INPUT = False
End Sub
