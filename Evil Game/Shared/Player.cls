VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum MovementConstants
    moveLeft = 1
    moveUp = 2
    moveRight = 4
    moveDown = 8
End Enum

Public ConnectionHandle As Long

Public Name As String
Public X As Integer
Public Y As Integer
Public TilesetIndex As Integer
Public TileIndex As Integer
Public MaskIndex As Integer

Public Movement As MovementConstants

Friend Sub Render(Surface As Surface, Tilesets As Tilesets)
    'On Error Resume Next
    
    If TilesetIndex <> 0 And TileIndex <> 0 Then
        With Tilesets.Tileset(TilesetIndex)
            If MaskIndex = 0 Then
                StretchBlt Surface.hDC, X * 32, Y * 32, 32, 32, .Surface.hDC, .Tile(TileIndex).X, .Tile(TileIndex).Y, .Tile(TileIndex).Width, .Tile(TileIndex).Height, srccopy
            Else
                StretchBlt Surface.hDC, X * 32, Y * 32, 32, 32, .Surface.hDC, .Tile(MaskIndex).X, .Tile(MaskIndex).Y, .Tile(MaskIndex).Width, .Tile(MaskIndex).Height, srcand
                StretchBlt Surface.hDC, X * 32, Y * 32, 32, 32, .Surface.hDC, .Tile(TileIndex).X, .Tile(TileIndex).Y, .Tile(TileIndex).Width, .Tile(TileIndex).Height, srcpaint
            End If
        End With
    End If
End Sub
