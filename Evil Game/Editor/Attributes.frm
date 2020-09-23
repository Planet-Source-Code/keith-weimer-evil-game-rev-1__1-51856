VERSION 5.00
Begin VB.Form frmAttributes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map Attributes"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   Icon            =   "Attributes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   510
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Editor As frmEditor

Private Sub cmdCancel_Click()
    'On Error Resume Next
    
    Unload frmAttributes
End Sub

Private Sub cmdOK_Click()
    'On Error Resume Next
    
    If Editor Is Nothing Then
        frmAttributes.Hide
    Else
        Unload frmAttributes
    End If
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    
    If Editor Is Nothing Then
        txtWidth.Text = 100
        txtHeight.Text = 100
    Else
        txtWidth.Text = Editor.Width
        txtHeight.Text = Editor.Height
    End If
End Sub
