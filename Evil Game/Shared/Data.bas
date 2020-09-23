Attribute VB_Name = "modData"
Option Explicit
Option Compare Text

'Data handling module for maps, tilesets, animations, and players
'DAMN THIS THING IS BLOATED!

'Include: GDI.bas
'Include: FileSystem.bas

Enum RegionConstants
    rgnSolid = 1
End Enum

Type Bounds
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type

Type Tileset
    FileName As String
    Surface As Surface
    Tile() As Bounds
End Type

Type Tile
    TilesetIndex As Integer
    Index As Integer
End Type

Type Animation
    FrameIndex As Integer
    Frame() As Tile
End Type

Type Region
    Tile As Tile
    AnimationIndex As Integer
    Flags As Long
End Type

Type Tilesets
    FileName As String
    Tileset() As Tileset
End Type

Type Animations
    FileName As String
    Animation() As Animation
End Type

Type Map
    FileName As String
    Name As String
    Tilesets As Tilesets
    Animations As Animations
    Width As Integer
    Height As Integer
    Region() As Region
    OuterRegion As Region
End Type

Function RemoveRegions(Map As Map) As Map
    'On Error Resume Next
    
    With RemoveRegions
        .Name = Map.Name
        .Tilesets = Map.Tilesets
        .Animations = Map.Animations
        .Width = .Width
        .Height = .Height
    End With
End Function

Function CreateTile(TilesetIndex As Integer, Index As Integer) As Tile
    'On Error Resume Next
    
    With CreateTile
        .TilesetIndex = TilesetIndex
        .Index = Index
    End With
End Function

Function CreateRegion(Tile As Tile, AnimationIndex As Integer, Flags As Long) As Region
    'On Error Resume Next
    
    With CreateRegion
        .Tile = Tile
        .AnimationIndex = AnimationIndex
        .Flags = Flags
    End With
End Function

Sub RenderTile(Surface As Surface, X As Integer, Y As Integer, Tilesets As Tilesets, Tile As Tile)
    'On Error Resume Next
    
    If Tile.TilesetIndex <> 0 And Tile.Index <> 0 Then
        With Tilesets.Tileset(Tile.TilesetIndex)
            StretchBlt Surface.hDC, X * 32, Y * 32, 32, 32, .Surface.hDC, .Tile(Tile.Index).X, .Tile(Tile.Index).Y, .Tile(Tile.Index).Width, .Tile(Tile.Index).Height, srccopy
        End With
    End If
End Sub

Sub RenderRegion(Surface As Surface, X As Integer, Y As Integer, Map As Map, Region As Region)
    'On Error Resume Next
    
    If Region.AnimationIndex <> 0 Then Region.Tile = GetAnimationFrame(Map.Animations, Region.AnimationIndex)
    RenderTile Surface, X, Y, Map.Tilesets, Region.Tile
End Sub

Sub RenderMap(Surface As Surface, Map As Map, LocationX As Integer, LocationY As Integer)
    'On Error Resume Next
    
    Dim DisplayX As Integer
    Dim DisplayY As Integer
    Dim X As Integer
    Dim Y As Integer
    
    With Map
        For DisplayX = 0 To 14
            X = DisplayX + LocationX - 7
            For DisplayY = 0 To 14
                Y = DisplayY + LocationY - 7
                
                If InMapBounds(Map, X, Y) Then
                    RenderRegion Surface, DisplayX, DisplayY, Map, Map.Region(X, Y)
                Else
                    RenderRegion Surface, DisplayX, DisplayY, Map, Map.OuterRegion
                End If
            Next DisplayY
        Next DisplayX
    End With
End Sub

Function RenderTileset(Surface As Surface, Tilesets As Tilesets, TilesetIndex As Integer, Page As Integer, Width As Integer, Height As Integer) As Integer
    'On Error Resume Next
    
    If TilesetIndex <> 0 Then
        Dim TileIndex As Integer
        Dim Offset As Long
        
        Offset = Page * Width * Height
        
        With Tilesets
            For TileIndex = 1 To Width * Height
                If TileIndex + Offset <= GetTileCount(.Tileset(TilesetIndex)) Then
                    RenderTile Surface, ((TileIndex - 1) \ Height), ((TileIndex - 1) Mod Height), Tilesets, CreateTile(TilesetIndex, Offset + TileIndex)
                Else
                    Exit For
                End If
            Next TileIndex
            
            RenderTileset = GetTileCount(.Tileset(TilesetIndex)) \ (Width * Height)
        End With
    End If
End Function

Function RenderAnimations(Surface As Surface, Map As Map, Page As Integer, Width As Integer, Height As Integer)
    'On Error Resume Next
    
    Dim AnimationIndex As Integer
    Dim Offset As Long
    
    Offset = Page * Width * Height
    
    With Map
        For AnimationIndex = 1 To Width * Height
            If AnimationIndex + Offset <= GetAnimationCount(.Animations) Then
                RenderRegion Surface, ((AnimationIndex - 1) \ Height), ((AnimationIndex - 1) Mod Height), Map, CreateRegion(CreateTile(0, 0), AnimationIndex, 0)
            Else
                Exit For
            End If
        Next AnimationIndex
        
        RenderAnimations = GetAnimationCount(.Animations) \ (Width * Height)
    End With
End Function

Function GetResourcePath() As String
    'On Error Resume Next
    
    GetResourcePath = GetFullPathName(AddSlash(App.Path) & "..\Resources\")
End Function

Function GetResourceFile(ByVal FileName As String) As String
    'On Error Resume Next
    
    If GetDriveName(FileName) = Empty Then FileName = GetResourcePath & FileName
    
    GetResourceFile = GetFullPathName(FileName)
End Function

Function GetMap(ByVal FileName As String) As Map
    'On Error Resume Next
    
    With GetMap
        .FileName = FileName
        
        FileName = GetResourceFile(FileName)
        If FileExists(FileName) Then
            Dim FileNum As Integer
            Dim Header As String
                    
            FileNum = FreeFile
            Open FileName For Binary Access Read As #FileNum
                Header = FixedLengthString(3)
                
                Get #FileNum, , Header
                If Header = "EGM" Then
                    Dim Length As Long
                    
                    Get #FileNum, , Length
                    .Name = FixedLengthString(Length)
                    Get #FileNum, , .Name
                    
                    Get #FileNum, , Length
                    .Tilesets.FileName = FixedLengthString(Length)
                    Get #FileNum, , .Tilesets.FileName
                    
                    Get #FileNum, , Length
                    .Animations.FileName = FixedLengthString(Length)
                    Get #FileNum, , .Animations.FileName
                    
                    Get #FileNum, , .Width
                    Get #FileNum, , .Height
                    ResizeMap GetMap
                    Get #FileNum, , .Region
                    Get #FileNum, , .OuterRegion
                Else
                    MsgBox "Invalid file format.", vbExclamation
                End If
            Close #FileNum
        End If
    End With
End Function

Sub SaveMap(Map As Map)
    'On Error Resume Next
    
    Dim FileName As String
    Dim FileNum As Integer
    Dim Header As String
    
    Header = "EGM"
    
    With Map
        FileName = GetResourceFile(.FileName)
        
        If FileExists(FileName) Then Kill FileName
        FileNum = FreeFile
        Open FileName For Binary Access Write As #FileNum
            Put #FileNum, , Header
            
            Put #FileNum, , Len(.Name)
            Put #FileNum, , .Name
            
            Put #FileNum, , Len(.Tilesets.FileName)
            Put #FileNum, , .Tilesets.FileName
            
            Put #FileNum, , Len(.Animations.FileName)
            Put #FileNum, , .Animations.FileName
            
            Put #FileNum, , .Width
            Put #FileNum, , .Height
            Put #FileNum, , .Region
            Put #FileNum, , .OuterRegion
        Close #FileNum
    End With
End Sub

Function InMapBounds(Map As Map, X As Integer, Y As Integer) As Boolean
    'On Error Resume Next
    
    InMapBounds = X >= 0 And Y >= 0 And X <= Map.Width - 1 And Y <= Map.Height - 1
End Function

Sub ResizeMap(Map As Map, Optional ByVal Width As Integer = -1, Optional ByVal Height As Integer = -1, Optional Save As Boolean = False)
    'On Error Resume Next
    
    If Width = -1 Then Width = Map.Width
    If Height = -1 Then Height = Map.Height
    
    If Width >= 0 And Height >= 0 Then
        With Map
            If Save Then
                ReDim Preserve .Region(Width - 1, Height - 1)
            Else
                ReDim .Region(Width - 1, Height - 1)
            End If
            
            .Width = Width
            .Height = Height
        End With
    End If
End Sub

Sub FillMap(Map As Map, Region As Region)
    'On Error Resume Next
    
    Dim X As Integer
    Dim Y As Integer
    
    For X = 0 To Map.Width - 1
        For Y = 0 To Map.Height - 1
            Map.Region(X, Y) = Region
        Next Y
    Next X
End Sub

Sub ClearMap(Map As Map)
    'On Error Resume Next
    
    Dim EmptyMap As Map
    
    ClearTilesets Map.Tilesets
    
    Map = EmptyMap
End Sub

Function GetTilesets(Optional ByVal FileName As String, Optional UseSurfaces As Boolean = True) As Tilesets
    'On Error Resume Next
    
    If FileName = Empty Then FileName = "Default.egt"
    
    With GetTilesets
        .FileName = FileName
        
        FileName = GetResourceFile(FileName)
        If FileExists(FileName) Then
            Dim FileNum As Integer
            Dim Header As String
                    
            Header = FixedLengthString(3)
            
            FileNum = FreeFile
            Open FileName For Binary Access Read As #FileNum
                Get #FileNum, , Header
                If Header = "EGT" Then
                    Dim TilesetCount As Integer
                                    
                    Dim Length As Long
                    
                    Get #FileNum, , TilesetCount
                    If TilesetCount > 0 Then
                        Dim TilesetIndex As Integer
                        Dim TileCount As Integer
                        
                        ReDim .Tileset(1 To TilesetCount)
                        For TilesetIndex = 1 To TilesetCount
                            With .Tileset(TilesetIndex)
                                Get #FileNum, , Length
                                .FileName = FixedLengthString(Length)
                                Get #FileNum, , .FileName
                                
                                If UseSurfaces Then .Surface = CreateSurfaceFromFile(GetResourceFile(.FileName), True)
                                
                                Get #FileNum, , TileCount
                                If TileCount > 0 Then
                                    ReDim .Tile(1 To TileCount)
                                    Get #FileNum, , .Tile
                                End If
                            End With
                        Next TilesetIndex
                    End If
                Else
                    MsgBox "Invalid file format.", vbExclamation
                End If
            Close #FileNum
        End If
    End With
End Function

Sub SaveTilesets(Tilesets As Tilesets)
    'On Error Resume Next
    
    Dim FileName As String
    Dim FileNum As Integer
    Dim Header As String
    
    Header = "EGT"
    
    With Tilesets
        FileName = GetResourceFile(.FileName)
        
        If FileExists(FileName) Then Kill FileName
        FileNum = FreeFile
        Open FileName For Binary Access Write As #FileNum
            Dim TilesetIndex As Integer
            
            Put #FileNum, , Header
            Put #FileNum, , GetTilesetCount(Tilesets)
            
            For TilesetIndex = 1 To GetTilesetCount(Tilesets)
                With .Tileset(TilesetIndex)
                    Put #FileNum, , Len(.FileName)
                    Put #FileNum, , .FileName
                    
                    Put #FileNum, , GetTileCount(Tilesets.Tileset(TilesetIndex))
                    Put #FileNum, , .Tile
                End With
            Next TilesetIndex
        Close #FileNum
    End With
End Sub

Function GetTilesetCount(Tilesets As Tilesets) As Integer
    On Error Resume Next
    
    With Tilesets
        GetTilesetCount = UBound(.Tileset) - LBound(.Tileset) + 1
    End With
End Function

Sub AddTileset(Tilesets As Tilesets, Tileset As Tileset, Optional ByVal TilesetIndex As Integer)
    'On Error Resume Next
    
    If TilesetIndex = 0 Then TilesetIndex = GetTilesetCount(Tilesets) + 1
    
    With Tilesets
        If GetTilesetCount(Tilesets) = 0 Then
            ReDim .Tileset(1 To 1)
            
            .Tileset(1) = Tileset
        ElseIf TilesetIndex >= 1 And TilesetIndex <= GetTilesetCount(Tilesets) + 1 Then
            Dim Index As Integer
            
            ReDim Preserve .Tileset(LBound(.Tileset) To UBound(.Tileset) + 1)
            For Index = UBound(.Tileset) To TilesetIndex + 1 Step -1
                    .Tileset(Index) = .Tileset(Index - 1)
            Next Index
            
            .Tileset(TilesetIndex) = Tileset
        End If
    End With
End Sub

Sub RemoveTileset(Tilesets As Tilesets, TilesetIndex As Integer)
    'On Error Resume Next
    
    With Tilesets
        Select Case GetTilesetCount(Tilesets)
            Case 1: Erase .Tileset
            Case Is > 1
                Dim Index As Integer
                
                For Index = TilesetIndex To UBound(.Tileset) - 1
                    .Tileset(Index) = .Tileset(Index + 1)
                Next Index
                
                ReDim Preserve .Tileset(LBound(.Tileset) To UBound(.Tileset) - 1)
        End Select
    End With
End Sub

Sub ClearTilesets(Tilesets As Tilesets)
    'On Error Resume Next
    
    Dim TilesetIndex As Integer
    
    With Tilesets
        For TilesetIndex = 1 To GetTilesetCount(Tilesets)
            DeleteSurface .Tileset(TilesetIndex).Surface
        Next TilesetIndex
        
        Erase .Tileset
    End With
End Sub

Function GetTileCount(Tileset As Tileset) As Integer
    On Error Resume Next
    
    With Tileset
        GetTileCount = UBound(.Tile) - LBound(.Tile) + 1
    End With
End Function

Sub AddTile(Tileset As Tileset, Tile As Bounds, Optional ByVal TileIndex As Integer)
    'On Error Resume Next
    
    If TileIndex = 0 Then TileIndex = GetTileCount(Tileset) + 1
    
    With Tileset
        If GetTileCount(Tileset) = 0 Then
            ReDim .Tile(1 To 1)
            
            .Tile(1) = Tile
        ElseIf TileIndex >= 1 And TileIndex <= GetTileCount(Tileset) + 1 Then
            Dim Index As Integer
            
            ReDim Preserve .Tile(LBound(.Tile) To UBound(.Tile) + 1)
            For Index = UBound(.Tile) To TileIndex + 1 Step -1
                    .Tile(Index) = .Tile(Index - 1)
            Next Index
            
            .Tile(TileIndex) = Tile
        End If
    End With
End Sub

Sub RemoveTile(Tileset As Tileset, TileIndex As Integer)
    'On Error Resume Next
    
    With Tileset
        Select Case GetTileCount(Tileset)
            Case 1: Erase .Tile
            Case Is > 1
                Dim Index As Integer
                
                For Index = TileIndex To UBound(.Tile) - 1
                    .Tile(Index) = .Tile(Index + 1)
                Next Index
                
                ReDim Preserve .Tile(LBound(.Tile) To UBound(.Tile) - 1)
        End Select
    End With
End Sub

Function GetAnimations(Optional ByVal FileName As String) As Animations
    'On Error Resume Next
    
    If FileName = Empty Then FileName = "Default.ega"
    
    With GetAnimations
        .FileName = FileName
        
        FileName = GetResourceFile(FileName)
        If FileExists(FileName) Then
            Dim FileNum As Integer
            Dim Header As String
            
            Header = FixedLengthString(3)
            
            FileNum = FreeFile
            Open FileName For Binary Access Read As #FileNum
                Get #FileNum, , Header
                If Header = "EGA" Then
                    Dim AnimationCount As Integer
                                    
                    Get #FileNum, , AnimationCount
                    If AnimationCount > 0 Then
                        Dim AnimationIndex As Integer
                        Dim FrameCount As Integer
                        
                        ReDim .Animation(1 To AnimationCount)
                        For AnimationIndex = 1 To AnimationCount
                            With .Animation(AnimationIndex)
                                Get #FileNum, , .FrameIndex
                                Get #FileNum, , FrameCount
                                If FrameCount > 0 Then
                                    ReDim .Frame(1 To FrameCount)
                                    Get #FileNum, , .Frame
                                End If
                            End With
                        Next AnimationIndex
                    End If
                Else
                    MsgBox "Invalid file format.", vbExclamation
                End If
            Close #FileNum
        End If
    End With
End Function

Sub SaveAnimations(Animations As Animations)
    'On Error Resume Next
    
    Dim FileName As String
    Dim FileNum As Integer
    Dim Header As String
        
    Header = "EGA"
    
    With Animations
        FileName = GetResourceFile(.FileName)
        
        If FileExists(FileName) Then Kill FileName
        FileNum = FreeFile
        Open FileName For Binary Access Write As #FileNum
            Dim AnimationIndex As Integer
            
            Put #FileNum, , Header
            Put #FileNum, , GetAnimationCount(Animations)
            
            For AnimationIndex = 1 To GetAnimationCount(Animations)
                With .Animation(AnimationIndex)
                    Put #FileNum, , .FrameIndex
                    Put #FileNum, , GetFrameCount(Animations.Animation(AnimationIndex))
                    Put #FileNum, , .Frame
                End With
            Next AnimationIndex
        Close #FileNum
    End With
End Sub

Function GetAnimationFrame(Animations As Animations, AnimationIndex As Integer) As Tile
    On Error Resume Next
    
    With Animations.Animation(AnimationIndex)
        GetAnimationFrame = .Frame(.FrameIndex)
    End With
End Function

Sub Animate(Animations As Animations)
    'On Error Resume Next
    
    Dim AnimationIndex As Integer
    
    With Animations
        For AnimationIndex = 1 To GetAnimationCount(Animations)
            With .Animation(AnimationIndex)
                If .FrameIndex >= UBound(.Frame) Then
                    .FrameIndex = LBound(.Frame)
                Else
                    .FrameIndex = .FrameIndex + 1
                End If
            End With
        Next AnimationIndex
    End With
End Sub

Function GetAnimationCount(Animations As Animations) As Integer
    On Error Resume Next
    
    With Animations
        GetAnimationCount = UBound(.Animation) - LBound(.Animation) + 1
    End With
End Function

Sub AddAnimation(Animations As Animations, Animation As Animation, Optional ByVal AnimationIndex As Integer)
    'On Error Resume Next
    
    If AnimationIndex = 0 Then AnimationIndex = GetAnimationCount(Animations) + 1
    
    With Animations
        If GetAnimationCount(Animations) = 0 Then
            ReDim .Animation(1 To 1)
            
            .Animation(1) = Animation
        ElseIf AnimationIndex >= 1 And AnimationIndex <= GetAnimationCount(Animations) + 1 Then
            Dim Index As Integer
            
            ReDim Preserve .Animation(LBound(.Animation) To UBound(.Animation) + 1)
            For Index = UBound(.Animation) To AnimationIndex + 1 Step -1
                .Animation(Index) = .Animation(Index - 1)
            Next Index
            
            .Animation(AnimationIndex) = Animation
        End If
    End With
End Sub

Sub RemoveAnimation(Animations As Animations, AnimationIndex As Integer)
    'On Error Resume Next
    
    With Animations
        Select Case GetAnimationCount(Animations)
            Case 1: Erase .Animation
            Case Is > 1
                Dim Index As Integer
            
                For Index = AnimationIndex To UBound(.Animation) - 1
                    .Animation(Index) = .Animation(Index + 1)
                Next Index
                
                ReDim Preserve .Animation(LBound(.Animation) To UBound(.Animation) - 1)
        End Select
    End With
End Sub

Sub ClearAnimations(Animations As Animations)
    'On Error Resume Next
    
    Erase Animations.Animation
End Sub

Function GetFrameCount(Animation As Animation) As Integer
    On Error Resume Next
    
    With Animation
        GetFrameCount = UBound(.Frame) - LBound(.Frame) + 1
    End With
End Function

Sub AddFrame(Animation As Animation, Frame As Tile, Optional ByVal FrameIndex As Integer)
    'On Error Resume Next
    
    If FrameIndex = 0 Then FrameIndex = GetFrameCount(Animation) + 1
    
    With Animation
        If GetFrameCount(Animation) = 0 Then
            ReDim .Frame(1 To 1)
            
            .Frame(1) = Frame
        ElseIf FrameIndex >= 1 And FrameIndex <= GetFrameCount(Animation) + 1 Then
            Dim Index As Integer
            
            ReDim Preserve .Frame(LBound(.Frame) To UBound(.Frame) + 1)
            For Index = UBound(.Frame) To FrameIndex + 1 Step -1
                .Frame(Index) = .Frame(Index - 1)
            Next Index
            
            .Frame(FrameIndex) = Frame
        End If
    End With
End Sub

Sub RemoveFrame(Animation As Animation, FrameIndex As Integer)
    'On Error Resume Next
    
    With Animation
        Select Case GetFrameCount(Animation)
            Case 1: Erase .Frame
            Case Is > 1
                Dim Index As Integer
                
                For Index = FrameIndex To UBound(.Frame) - 1
                    .Frame(Index) = .Frame(Index + 1)
                Next Index
                
                ReDim Preserve .Frame(LBound(.Frame) To UBound(.Frame) - 1)
        End Select
    End With
End Sub
