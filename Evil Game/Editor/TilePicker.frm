VERSION 5.00
Begin VB.Form frmTilePicker 
   Caption         =   "Tile Picker"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vsbTileOffset 
      Enabled         =   0   'False
      Height          =   7695
      Left            =   4080
      Max             =   1
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picTile 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      HasDC           =   0   'False
      Height          =   7680
      Left            =   1200
      MouseIcon       =   "TilePicker.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   1
      Top             =   0
      Width           =   2880
   End
   Begin VB.ListBox lstTileset 
      Height          =   1935
      IntegralHeight  =   0   'False
      ItemData        =   "TilePicker.frx":0152
      Left            =   0
      List            =   "TilePicker.frx":015D
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblTile 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   45
   End
End
Attribute VB_Name = "frmTilePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseX As Integer
Dim MouseY As Integer

Dim BackBuffer As Surface

Sub Update()
    'On Error Resume Next
    
    Dim TilesetIndex As Integer
    
    lstTileset.Clear
    lstTileset.AddItem "(None)"
    lstTileset.AddItem "(Animation)"
    lstTileset.ItemData(lstTileset.NewIndex) = -1
    
    For TilesetIndex = 1 To GetTilesetCount(Map.Tilesets)
        lstTileset.AddItem TilesetIndex
        lstTileset.ItemData(lstTileset.NewIndex) = TilesetIndex
    Next TilesetIndex
    
    Display
End Sub

Sub Display(Optional GetPages As Boolean = False)
    'On Error Resume Next
    
    Dim TilesetIndex As Integer
    
    If lstTileset.ListIndex <> -1 Then TilesetIndex = lstTileset.ItemData(lstTileset.ListIndex)
    
    ClearSurface BackBuffer

    If GetPages Then
        If TilesetIndex = 0 Then
            vsbTileOffset.Enabled = False
        Else
            Dim Pages As Integer
            
            If TilesetIndex = -1 Then
                Pages = RenderAnimations(BackBuffer, Map, 0, 6, 16)
            Else
                Pages = RenderTileset(BackBuffer, Map.Tilesets, TilesetIndex, 0, 6, 16)
            End If
            
            If Pages > 1 Then
                vsbTileOffset.Max = Pages - 1
                vsbTileOffset.Enabled = True
            Else
                vsbTileOffset.Enabled = False
            End If
        End If
    Else
        If TilesetIndex <> 0 Then
            If TilesetIndex = -1 Then
                RenderAnimations BackBuffer, Map, 0, 6, 16
            Else
                RenderTileset BackBuffer, Map.Tilesets, TilesetIndex, vsbTileOffset.Value, 6, 16
            End If
        End If
    End If
    DrawRectangle BackBuffer, MouseX * 32, MouseY * 32, MouseX * 32 + 32, MouseY * 32 + 32
    
    RenderSurface picTile.hDC, BackBuffer
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    
    BackBuffer = CreateSurface(192, 512)
    SetPen BackBuffer, vbSolid, 0, vbWhite
    SetBrush BackBuffer, BS_INVISIBLE
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'On Error Resume Next
    
    If Locked_TilePicker Then
        Hide
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next
    
    DeleteSurface BackBuffer
End Sub

Private Sub lstTileset_Click()
    'On Error Resume Next
    
    Region.Tile.TilesetIndex = 0
    Region.Tile.Index = 0
    Region.AnimationIndex = 0
    
    vsbTileOffset.Value = 0
    Display True
End Sub

Private Sub picTile_Click()
    'On Error Resume Next
    
    Dim TilesetIndex As Integer
    
    TilesetIndex = lstTileset.ItemData(lstTileset.ListIndex)
    
    If TilesetIndex = -1 Then
        Region.Tile.TilesetIndex = 0
        Region.Tile.Index = 0
        Region.AnimationIndex = vsbTileOffset.Value * 96 + MouseX * 16 + MouseY + 1
    Else
        Region.Tile.TilesetIndex = TilesetIndex
        Region.Tile.Index = vsbTileOffset.Value * 96 + MouseX * 16 + MouseY + 1
        Region.AnimationIndex = 0
    End If
End Sub

Private Sub picTile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error Resume Next
    
    MouseX = X \ 32
    MouseY = Y \ 32
    
    lblTile.Caption = vsbTileOffset.Value * 96 + MouseX * 16 + MouseY + 1
    
    Display
End Sub

Private Sub picTile_Paint()
    'On Error Resume Next
    
    RenderSurface picTile.hDC, BackBuffer
End Sub

Private Sub vsbTileOffset_Change()
    'On Error Resume Next
    
    Display
End Sub

Private Sub vsbTileOffset_Scroll()
    'On Error Resume Next
    
    Display
End Sub
