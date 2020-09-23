VERSION 5.00
Begin VB.Form frmBrush 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Brush"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBrush 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1200
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.CheckBox chkSolid 
      Caption         =   "Solid"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BackBuffer As Surface

Sub Display()
    'On Error Resume Next
    
    chkSolid.Value = Abs(GetFlag(Region.Flags, rgnSolid))
    
    ClearSurface BackBuffer
    RenderRegion BackBuffer, 0, 0, Map, Region
    RenderSurface picBrush.hDC, BackBuffer
End Sub

Private Sub chkSolid_Click()
    'On Error Resume Next
    
    SetFlag Region.Flags, rgnSolid, chkSolid.Value
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    
    BackBuffer = CreateSurface(32, 32)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'On Error Resume Next
    
    If Locked_Brush Then
        Hide
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next
    
    DeleteSurface BackBuffer
End Sub

Private Sub picBrush_Paint()
    'On Error Resume Next
    
    RenderSurface picBrush.hDC, BackBuffer
End Sub
