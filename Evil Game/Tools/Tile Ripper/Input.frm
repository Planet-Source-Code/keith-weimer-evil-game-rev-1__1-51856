VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInput 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    If LOCKED_INPUT Then
        Cancel = True
        Hide
    End If
End Sub

Private Sub picInput_Change()
    Dim OffsetX As Single
    Dim OffsetY As Single
    
    OffsetX = Width - ScaleWidth
    OffsetY = Height - ScaleHeight
    
    Width = OffsetX + picInput.Width
    Height = OffsetY + picInput.Height
End Sub
