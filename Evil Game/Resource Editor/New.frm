VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   1950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox lstNew 
      Height          =   1035
      ItemData        =   "New.frx":0000
      Left            =   120
      List            =   "New.frx":000A
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'On Error Resume Next
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'On Error Resume Next
    
    frmMain.NewFile lstNew.ListIndex + 1
    Unload Me
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    
    lstNew.ListIndex = frmMain.Mode
End Sub
