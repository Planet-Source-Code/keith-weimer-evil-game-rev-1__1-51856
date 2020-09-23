VERSION 5.00
Begin VB.Form frmWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wizard"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOffset 
      Caption         =   "Offset"
      Height          =   975
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtOffsetY 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtOffsetX 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblOffsetY 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   150
      End
      Begin VB.Label lblOffsetX 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.Timer tmrEnable 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Grid"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
      Begin VB.TextBox txtGridWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtGridHeight 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblGridWidth 
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblGridHeight 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame fraTile 
      Caption         =   "Tile"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtTileHeight 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtTileWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTileHeight 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTileWidth 
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EditIndex As Integer

Sub Activate(EditIndex As Integer)
    'On Error Resume Next
    
    Me.EditIndex = EditIndex
    
    Show vbModal, frmEdit
End Sub

Private Sub cmdCancel_Click()
    'On Error Resume Next
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'On Error Resume Next
    
    Dim OffsetX As Long
    Dim OffsetY As Long
    Dim GridX As Long
    Dim GridY As Long
    Dim SizeX As Long
    Dim SizeY As Long
    
    Dim X As Long
    Dim Y As Long
    
    Dim Tile As Bounds
    
    OffsetX = CLng(txtOffsetX.Text)
    OffsetY = CLng(txtOffsetY.Text)
    GridX = CLng(txtGridWidth.Text)
    GridY = CLng(txtGridHeight.Text)
    SizeX = CLng(txtTileWidth.Text)
    SizeY = CLng(txtTileHeight.Text)
    
    For X = 0 To GridX - 1
        For Y = 0 To GridY - 1
            With Tile
                .X = OffsetX + X * SizeX
                .Y = OffsetY + Y * SizeY
                .Width = SizeX
                .Height = SizeY
            End With
            
            AddTile modMain.Tilesets.Tileset(EditIndex), Tile
        Next Y
    Next X
    
    Unload Me
End Sub
