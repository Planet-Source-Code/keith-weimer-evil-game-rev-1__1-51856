VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEnable 
      Interval        =   1
      Left            =   4680
      Top             =   4200
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   4680
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source File"
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   285
         Left            =   4320
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         Caption         =   "Filename:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "M"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   19
      ToolTipText     =   "Modify"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "Wizard"
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      ToolTipText     =   "Wizard"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "R"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   12
      ToolTipText     =   "Remove"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "A"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Add"
      Top             =   3960
      Width           =   375
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Frame Info"
      Height          =   975
      Index           =   1
      Left            =   1440
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtFrame 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtFrame 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTile 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tile:"
         Height          =   195
         Left            =   1515
         TabIndex        =   16
         Top             =   600
         Width           =   300
      End
      Begin VB.Label lblTileset 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tileset:"
         Height          =   195
         Left            =   1305
         TabIndex        =   14
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Tile Info"
      Height          =   975
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtTile 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtTile 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtTile 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtTile 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTileHeight 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Height:"
         Height          =   195
         Left            =   1305
         TabIndex        =   8
         Top             =   600
         Width           =   510
      End
      Begin VB.Label lblTileWidth 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   195
         Left            =   1350
         TabIndex        =   7
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblTileY 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   600
         Width           =   150
      End
      Begin VB.Label lblTileX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.ListBox lstMain 
      Columns         =   3
      Height          =   2535
      IntegralHeight  =   0   'False
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblList 
      AutoSize        =   -1  'True
      Caption         =   "Item(s):"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EditIndex As Integer

Sub Activate(EditIndex As Long)
    'On Error Resume Next
    
    Dim Index As Long
        
    Me.EditIndex = EditIndex
    
    cmdWizard.Visible = frmMain.Mode = emTileset
    fraSource.Visible = frmMain.Mode = emTileset
    
    For Index = fraEdit.LBound To fraEdit.UBound
        fraEdit(Index).Left = fraEdit(0).Left
        fraEdit(Index).Top = fraEdit(0).Top
        fraEdit(Index).Visible = frmMain.Mode = Index + 1
    Next Index
    
    Rebuild
    
    Show vbModal, frmMain
End Sub

Sub Rebuild()
    'On Error Resume Next
    
    Select Case frmMain.Mode
        Case emTileset
            Dim TileIndex As Long
            
            With modMain.Tilesets.Tileset(EditIndex)
                lstMain.Clear
                For TileIndex = 1 To GetTileCount(modMain.Tilesets.Tileset(EditIndex))
                    With .Tile(TileIndex)
                        lstMain.AddItem TileIndex & " (" & .X & ", " & .Y & ")->(" & .Width & ", " & .Height & ")"
                    End With
                Next TileIndex
                
                lstMain.ListIndex = lstMain.ListCount - 1
                
                txtFileName.Text = .FileName
            End With
        Case emAnimation
            Dim FrameIndex As Long
            
            With modMain.Animations.Animation(EditIndex)
                lstMain.Clear
                For FrameIndex = 1 To GetFrameCount(modMain.Animations.Animation(EditIndex))
                    With .Frame(FrameIndex)
                        lstMain.AddItem FrameIndex & " (" & .TilesetIndex & ", " & .Index & ")"
                    End With
                    lstMain.ListIndex = .FrameIndex - 1
                Next FrameIndex
            End With
    End Select
End Sub

Private Sub cmdAdd_Click()
    'On Error Resume Next
    
    Select Case frmMain.Mode
        Case emTileset
            Dim Tile As Bounds
            
            With Tile
                .X = CLng(txtTile(0).Text)
                .Y = CLng(txtTile(1).Text)
                .Width = CLng(txtTile(2).Text)
                .Height = CLng(txtTile(3).Text)
            End With
            
            AddTile modMain.Tilesets.Tileset(EditIndex), Tile
            Rebuild
        Case emAnimation
            Dim Frame As Tile
            
            With Frame
                .TilesetIndex = CInt(txtFrame(0).Text)
                .Index = CInt(txtFrame(1).Text)
            End With
            
            AddFrame modMain.Animations.Animation(EditIndex), Frame
            Rebuild
    End Select
End Sub

Private Sub cmdBrowse_Click()
    On Error Resume Next
    
    With cdlMain
        .Filter = PictureFilter
        .FilterIndex = 0
        .Flags = OpenFlags
        .InitDir = GetParentFolderName(.FileName)
        .FileName = .FileName
        .ShowOpen
        
        If Err.Number <> cdlCancel Then
            modMain.Tilesets.Tileset(EditIndex).FileName = Replace(cdlMain.FileName, GetResourcePath, Empty)
            Rebuild
        End If
    End With
End Sub

Private Sub cmdModify_Click()
    'On Error Resume Next
    
    Select Case frmMain.Mode
        Case emTileset
            Dim Tile As Bounds
            
            With Tile
                .X = CLng(txtTile(0).Text)
                .Y = CLng(txtTile(1).Text)
                .Width = CLng(txtTile(2).Text)
                .Height = CLng(txtTile(3).Text)
            End With
            
            modMain.Tilesets.Tileset(EditIndex).Tile(lstMain.ListIndex + 1) = Tile
        Case emAnimation
            Dim Frame As Tile
            
            With Frame
                .TilesetIndex = CInt(txtFrame(0).Text)
                .Index = CInt(txtFrame(1).Text)
            End With
            
            modMain.Animations.Animation(EditIndex).Frame(lstMain.ListIndex + 1) = Frame
    End Select
    
    Rebuild
End Sub

Private Sub cmdRemove_Click()
    'On Error Resume Next
    
    Dim Index As Integer
    
    Select Case frmMain.Mode
        Case emTileset
            For Index = lstMain.ListCount - 1 To 0 Step -1
                If lstMain.Selected(Index) Then RemoveTile modMain.Tilesets.Tileset(EditIndex), Index + 1
            Next Index
        Case emAnimation
            For Index = lstMain.ListCount - 1 To 0 Step -1
                If lstMain.Selected(Index) Then RemoveFrame modMain.Animations.Animation(EditIndex), Index + 1
            Next Index
    End Select

    Rebuild
End Sub

Private Sub cmdWizard_Click()
    'On Error Resume Next
    
    frmWizard.Activate EditIndex
    Rebuild
End Sub

Private Sub lstMain_Click()
    'On Error Resume Next
    
    If lstMain.ListIndex <> -1 Then
        Select Case frmMain.Mode
            Case emTileset
                With modMain.Tilesets.Tileset(EditIndex).Tile(lstMain.ListIndex + 1)
                    txtTile(0).Text = .X
                    txtTile(1).Text = .Y
                    txtTile(2).Text = .Width
                    txtTile(3).Text = .Height
                End With
            Case emAnimation
                With modMain.Animations.Animation(EditIndex)
                    .FrameIndex = lstMain.ListIndex + 1
                    With .Frame(.FrameIndex)
                        txtFrame(0).Text = .TilesetIndex
                        txtFrame(1).Text = .Index
                    End With
                End With
        End Select
    End If
End Sub

Private Sub tmrEnable_Timer()
    'On Error Resume Next
    
    Dim Enabled As Boolean
    Dim Index As Long
    
    Select Case frmMain.Mode
        Case emTileset
            Enabled = True
            For Index = txtTile.LBound To txtTile.UBound
                Enabled = Enabled And IsNumeric(txtTile(Index).Text)
            Next Index
        Case emAnimation
            Enabled = True
            For Index = txtFrame.LBound To txtFrame.UBound
                Enabled = Enabled And IsNumeric(txtFrame(Index).Text)
            Next Index
    End Select
    
    cmdAdd.Enabled = Enabled
    cmdRemove.Enabled = lstMain.ListIndex <> -1
    cmdModify.Enabled = Enabled And lstMain.ListIndex <> -1
End Sub

Private Sub txtFileName_Change()
    'On Error Resume Next
    
    modMain.Tilesets.Tileset(EditIndex).FileName = txtFileName.Text
End Sub
