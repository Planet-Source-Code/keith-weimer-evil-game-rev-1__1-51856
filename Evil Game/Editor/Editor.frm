VERSION 5.00
Begin VB.Form frmEditor 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Editor.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "Editor.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   Begin VB.Timer tmrDisplay 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Modified As Boolean
Public Saved As Boolean

Dim Draw As Boolean

Public MouseX As Integer
Public MouseY As Integer
Public LocationX As Integer
Public LocationY As Integer

Dim BackBuffer As Surface

Dim Map As Map
Dim Brush As Region

Sub SetModified(ByVal Modified As Boolean)
    'On Error Resume Next
    
    Dim Text As String
    
    Me.Modified = Modified
    
    If Map.FileName = Empty Then
        Text = "New Map"
    Else
        Text = GetFileName(Map.FileName)
    End If
    
    Caption = Text & IIf(Modified, "*", Empty)
End Sub

Friend Sub NewMap(FileName As String, Width As Integer, Height As Integer)
    'On Error Resume Next
    
    With Map
        .Tilesets.FileName = "Tileset.egt"
        .Animations.FileName = "Animation.ega"
        .Tilesets = GetTilesets(.Tilesets.FileName)
        .Animations = GetAnimations(.Animations.FileName)
    End With
    
    ResizeMap Map, Width, Height
    FillMap Map, CreateRegion(CreateTile(1, 29), 0, 0)
    
    SetModified False
    Saved = False
    Display
End Sub

Sub OpenMap(FileName As String)
    'On Error Resume Next
    
    ClearMap Map
    
    Map = GetMap(FileName)
    Map.Tilesets = GetTilesets(Map.Tilesets.FileName)
    Map.Animations = GetAnimations(Map.Animations.FileName)
    
    SetModified False
    Saved = True
    Display
End Sub

Sub SaveMap(Optional ByVal FileName As String)
    'On Error Resume Next
    
    If FileName <> Empty Then Map.FileName = FileName
    
    modData.SaveMap Map
    SetModified False
    Saved = True
End Sub

Sub DrawBrush(X As Integer, Y As Integer)
    'On Error Resume Next
    
    Map.Region(X, Y) = Region
    
    SetModified True
    If Not tmrDisplay.Enabled Then Display
End Sub

Sub Display()
    'On Error Resume Next
    
    Dim ActualX As Integer
    Dim ActualY As Integer
    
    If Me Is frmMain.ActiveEditor Then
        ActualX = MouseX + LocationX - 7
        ActualY = MouseY + LocationY - 7
    
        frmInfo.lblDisplay.Caption = "Display Coordinates: (" & MouseX & "," & MouseY & ")"
        If InMapBounds(Map, ActualX, ActualY) Then
            frmInfo.lblMap.Caption = "Map Coordinates: (" & ActualX & "," & ActualY & ")"
            frmInfo.lblSolid.Caption = "Solid: " & GetFlag(Map.Region(ActualX, ActualY).Flags, rgnSolid)
            frmInfo.lblTileset.Caption = "Tileset: " & Map.Region(ActualX, ActualY).Tile.TilesetIndex
            frmInfo.lblTile.Caption = "Tile: " & Map.Region(ActualX, ActualY).Tile.Index
            frmInfo.lblAnimation.Caption = "Animation: " & Map.Region(ActualX, ActualY).AnimationIndex
            
            If Draw Then DrawBrush ActualX, ActualY
        Else
            frmInfo.lblMap.Caption = "Map Coordinates: N/A"
            frmInfo.lblSolid.Caption = "Solid: True"
            frmInfo.lblTileset.Caption = "Tileset: " & Map.OuterRegion.Tile.TilesetIndex
            frmInfo.lblTile.Caption = "Tile: " & Map.OuterRegion.Tile.Index
            frmInfo.lblAnimation.Caption = "Animation: " & Map.OuterRegion.AnimationIndex
        End If
    End If
    
    ClearSurface BackBuffer
    RenderMap BackBuffer, Map, LocationX, LocationY
    
    If frmMain.mnuViewHighlightSolids.Checked Then
        Dim DisplayX As Integer
        Dim DisplayY As Integer
        Dim X As Integer
        Dim Y As Integer
        
        For DisplayX = 0 To 14
            X = DisplayX + LocationX - 7
            For DisplayY = 0 To 14
                Y = DisplayY + LocationY - 7
                
                If InMapBounds(Map, X, Y) Then
                    If Map.Region(X, Y).Flags And rgnSolid Then
                        DrawRectangle BackBuffer, DisplayX * 32, DisplayY * 32, DisplayX * 32 + 32, DisplayY * 32 + 32
                        DrawLine BackBuffer, DisplayX * 32, DisplayY * 32, DisplayX * 32 + 32, DisplayY * 32 + 32
                    End If
                Else
                    DrawRectangle BackBuffer, DisplayX * 32, DisplayY * 32, DisplayX * 32 + 32, DisplayY * 32 + 32
                    DrawLine BackBuffer, DisplayX * 32, DisplayY * 32, DisplayX * 32 + 32, DisplayY * 32 + 32
                End If
            Next DisplayY
        Next DisplayX
    End If
    
    DrawRectangle BackBuffer, MouseX * 32, MouseY * 32, MouseX * 32 + 32, MouseY * 32 + 32
    
    RenderSurface hDC, BackBuffer
End Sub

Private Sub Form_Activate()
    'On Error Resume Next
    
    Set frmMain.ActiveEditor = Me
    modMain.Map = RemoveRegions(Map)
    Region = Brush
    
    frmTilePicker.Update
End Sub

Private Sub Form_Deactivate()
    'On Error Resume Next
    
    Brush = Region
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyLeft: If InMapBounds(Map, LocationX - 1, 0) Then LocationX = LocationX - 1
        Case vbKeyRight: If InMapBounds(Map, LocationX + 1, 0) Then LocationX = LocationX + 1
        Case vbKeyUp: If InMapBounds(Map, 0, LocationY - 1) Then LocationY = LocationY - 1
        Case vbKeyDown: If InMapBounds(Map, LocationY + 1, 0) Then LocationY = LocationY + 1
    End Select
    
    Display
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    
    Dim OffsetWidth As Single
    Dim OffsetHeight As Single
    
    ScaleMode = vbTwips
    OffsetWidth = Width - ScaleWidth
    OffsetHeight = Height - ScaleHeight
    ScaleMode = vbPixels
    
    Width = OffsetWidth + 7200
    Height = OffsetHeight + 7200
    
    BackBuffer = CreateSurface(480, 480)
    SetPen BackBuffer, vbSolid, 0, vbWhite
    SetBrush BackBuffer, BS_INVISIBLE, 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error Resume Next
    
    Dim ActualX As Integer
    Dim ActualY As Integer
    
    ActualX = (X \ 32) + LocationX - 7
    ActualY = (Y \ 32) + LocationY - 7
    
    If InMapBounds(Map, ActualX, ActualY) Then
        DrawBrush ActualX, ActualY
        
        Draw = True
    Else
        Map.OuterRegion = Region
    End If
    
    SetModified True
    Display
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error Resume Next
    
    MouseX = X \ 32
    MouseY = Y \ 32
    
    Display
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error Resume Next
    
    Draw = False
End Sub

Private Sub Form_Paint()
    'On Error Resume Next
    
    RenderSurface hDC, BackBuffer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'On Error Resume Next
    
    If (Not Saved Or Modified) And Map.FileName <> Empty Then
        Dim Prompt As String
        
        Prompt = "'" & GetFileName(Map.FileName) & "'"
        If Not Saved Then
            Prompt = Prompt & " has not been saved.  Save it now?"
        ElseIf Modified Then
            Prompt = Prompt & " has been modified.  Would you like to save the changes?"
        End If
        
        Select Case MsgBox(Prompt, vbQuestion Or vbYesNoCancel)
            Case vbYes
                Set frmMain.ActiveEditor = Me
                frmMain.mnuSave_Click
            Case vbCancel: Cancel = True
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next
    
    Dim EmptyRegion As Region
    
    modMain.Region = EmptyRegion
    
    ClearMap modMain.Map
    DeleteSurface BackBuffer
    
    frmTilePicker.Update
End Sub

Private Sub tmrDisplay_Timer()
    'On Error Resume Next
    
    Animate Map.Animations
    
    Display
End Sub
