VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Info"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblTileset 
      AutoSize        =   -1  'True
      Caption         =   "Tileset: N/A"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblAttributes 
      AutoSize        =   -1  'True
      Caption         =   "Region Attributes"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblSolid 
      AutoSize        =   -1  'True
      Caption         =   "Solid: N/A"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblAnimation 
      AutoSize        =   -1  'True
      Caption         =   "Animation: N/A"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label lblTile 
      AutoSize        =   -1  'True
      Caption         =   "Tile: N/A"
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label lblMap 
      AutoSize        =   -1  'True
      Caption         =   "Map Coordinates: N/A"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1590
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Display Coordinates: N/A"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'On Error Resume Next
    
    If Locked_Info Then
        Hide
        Cancel = True
    End If
End Sub
