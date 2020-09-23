VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tile Ripper"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPicture 
      Caption         =   "Picture"
      Height          =   1095
      Left            =   2400
      TabIndex        =   30
      Top             =   2040
      Width           =   2175
      Begin VB.TextBox txtActualPictureSizeY 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtActualPictureSizeX 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtGridPictureSizeY 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtGridPictureSizeX 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblPictureGrid 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Grid"
         Height          =   195
         Left            =   870
         TabIndex        =   38
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblPictureActual 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Actual"
         Height          =   195
         Left            =   1515
         TabIndex        =   37
         Top             =   120
         Width           =   465
      End
      Begin VB.Label lblPictureSizeY 
         AutoSize        =   -1  'True
         Caption         =   "Size Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblPictureSizeX 
         AutoSize        =   -1  'True
         Caption         =   "Size X:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraOffset 
      Caption         =   "Offset"
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   2175
      Begin VB.TextBox txtInputOffsetX 
         Height          =   285
         Left            =   720
         TabIndex        =   27
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtInputOffsetY 
         Height          =   285
         Left            =   720
         TabIndex        =   26
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtOutputOffsetY 
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtOutputOffsetX 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblOffsetX 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   150
      End
      Begin VB.Label lblOffsetY 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   150
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Grid"
      Height          =   1455
      Left            =   2400
      TabIndex        =   17
      Top             =   480
      Width           =   1575
      Begin VB.TextBox txtGridSizeY 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtGridSizeX 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdCalculateGridXY 
         Caption         =   "Calculate"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblGridSizeX 
         AutoSize        =   -1  'True
         Caption         =   "Size X:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblGridSizeY 
         AutoSize        =   -1  'True
         Caption         =   "Size Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame fraTile 
      Caption         =   "Tile"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   2175
      Begin VB.TextBox txtInputSizeX 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Text            =   "32"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtInputSizeY 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "32"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtOutputSizeY 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Text            =   "32"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtOutputSizeX 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "32"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdCalculateXY 
         Caption         =   "Calculate"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblSizeX 
         AutoSize        =   -1  'True
         Caption         =   "Size X:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSizeY 
         AutoSize        =   -1  'True
         Caption         =   "Size Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblInput 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Input"
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblOutput 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Output"
         Height          =   195
         Left            =   1500
         TabIndex        =   13
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   2055
      Left            =   4920
      TabIndex        =   5
      Top             =   480
      Width           =   1935
      Begin VB.PictureBox picOutput 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   112
         TabIndex        =   6
         Top             =   240
         Width           =   1680
      End
   End
   Begin VB.PictureBox picInput 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   495
      Left            =   4200
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   5160
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label lblFilename 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Sub Update()
    On Error Resume Next
    
    Dim Exists As Boolean
    
    Dim InputSizeX As Long
    Dim InputSizeY As Long
    Dim InputOffsetX As Long
    Dim InputOffsetY As Long
    Dim OutputSizeX As Long
    Dim OutputSizeY As Long
    Dim OutputOffsetX As Long
    Dim OutputOffsetY As Long
    Dim GridSizeX As Long
    Dim GridSizeY As Long
    
    Dim X As Long
    Dim Y As Long
    
    InputSizeX = Val(txtInputSizeX.Text)
    InputSizeY = Val(txtInputSizeY.Text)
    InputOffsetX = Val(txtInputOffsetX.Text)
    InputOffsetY = Val(txtInputOffsetY.Text)
    OutputSizeX = Val(txtOutputSizeX.Text)
    OutputSizeY = Val(txtOutputSizeY.Text)
    OutputOffsetX = Val(txtOutputOffsetX.Text)
    OutputOffsetY = Val(txtOutputOffsetY.Text)
    GridSizeX = Val(txtGridSizeX.Text)
    GridSizeY = Val(txtGridSizeY.Text)
    
    Exists = FileExists(txtFileName.Text)
    
    frmInput.Visible = Exists
    
    cmdCalculateXY.Enabled = Exists And GridSizeX > 0 And GridSizeY > 0
    cmdCalculateGridXY.Enabled = Exists And InputSizeX > 0 And InputSizeY > 0
    
    txtGridPictureSizeX.Text = InputSizeX * GridSizeX
    txtGridPictureSizeY.Text = InputSizeY * GridSizeY
    txtActualPictureSizeX.Text = picInput.ScaleWidth
    txtActualPictureSizeY.Text = picInput.ScaleHeight
    
    cmdExecute.Enabled = cmdCalculateXY.Enabled And cmdCalculateGridXY.Enabled And OutputSizeX > 0 And OutputSizeY > 0
    
    picOutput.Width = OutputSizeX * Screen.TwipsPerPixelX
    picOutput.Height = OutputSizeY * Screen.TwipsPerPixelY
    
    frmInput.picInput.Cls
    For X = 0 To GridSizeX
        frmInput.picInput.Line (InputOffsetX + (X * InputSizeX), InputOffsetY)-Step(0, GridSizeY * InputSizeY), vbBlack
    Next X

    For Y = 0 To GridSizeY
        frmInput.picInput.Line (InputOffsetX, InputOffsetY + (Y * InputSizeY))-Step(GridSizeX * InputSizeX, 0), vbBlack
    Next Y
End Sub

Private Sub cmdBrowse_Click()
    On Error Resume Next
    
    With cdlMain
        .FileName = txtFileName.Text
        .Filter = PictureFilter
        .InitDir = GetParentFolderName(txtFileName.Text)
        .Flags = OpenFlags
        .ShowOpen
        
        If Err.Number <> cdlCancel Then
            txtFileName.Text = .FileName
            picInput.Picture = LoadPicture(txtFileName.Text)
            frmInput.picInput.Picture = picInput.Picture
            Update
        End If
    End With
End Sub

Private Sub cmdCalculateGridXY_Click()
    On Error Resume Next
    
    Dim InputSizeX As Long
    Dim InputSizeY As Long
    Dim InputOffsetX As Long
    Dim InputOffsetY As Long
    
    InputSizeX = Val(txtInputSizeX.Text)
    InputSizeY = Val(txtInputSizeY.Text)
    InputOffsetX = Val(txtInputOffsetX.Text)
    InputOffsetY = Val(txtInputOffsetY.Text)

    txtGridSizeX.Text = (picInput.ScaleWidth - InputOffsetX) \ InputSizeX
    txtGridSizeY.Text = (picInput.ScaleHeight - InputOffsetY) \ InputSizeY
End Sub

Private Sub cmdCalculateXY_Click()
    On Error Resume Next
    
    Dim InputOffsetX As Long
    Dim InputOffsetY As Long
    Dim GridSizeX As Long
    Dim GridSizeY As Long
    
    Dim SizeX As Long
    Dim SizeY As Long
    
    InputOffsetX = Val(txtInputOffsetX.Text)
    InputOffsetY = Val(txtInputOffsetY.Text)
    GridSizeX = Val(txtGridSizeX.Text)
    GridSizeY = Val(txtGridSizeY.Text)
    
    SizeX = (picInput.ScaleWidth - InputOffsetX) \ GridSizeX
    SizeY = (picInput.ScaleHeight - InputOffsetY) \ GridSizeY
    
    txtInputSizeX.Text = SizeX
    txtInputSizeY.Text = SizeY
    txtOutputSizeX.Text = SizeX
    txtOutputSizeY.Text = SizeY
End Sub

Private Sub cmdExecute_Click()
    On Error Resume Next
    
    Dim InputSizeX As Long
    Dim InputSizeY As Long
    Dim InputOffsetX As Long
    Dim InputOffsetY As Long
    Dim OutputSizeX As Long
    Dim OutputSizeY As Long
    Dim OutputOffsetX As Long
    Dim OutputOffsetY As Long
    Dim GridSizeX As Long
    Dim GridSizeY As Long
    
    Dim X As Long
    Dim Y As Long
    
    InputSizeX = Val(txtInputSizeX.Text)
    InputSizeY = Val(txtInputSizeY.Text)
    InputOffsetX = Val(txtInputOffsetX.Text)
    InputOffsetY = Val(txtInputOffsetY.Text)
    OutputSizeX = Val(txtOutputSizeX.Text)
    OutputSizeY = Val(txtOutputSizeY.Text)
    OutputOffsetX = Val(txtOutputOffsetX.Text)
    OutputOffsetY = Val(txtOutputOffsetY.Text)
    GridSizeX = Val(txtGridSizeX.Text) - 1
    GridSizeX = Val(txtGridSizeY.Text) - 1

    picInput.Picture = LoadPicture(txtFileName.Text)
    picOutput.Width = OutputSizeX * Screen.TwipsPerPixelX
    picOutput.Height = OutputSizeY * Screen.TwipsPerPixelY
    
    For X = 0 To GridSizeX
        For Y = 0 To GridSizeY
            picOutput.Cls
            picOutput.PaintPicture picInput.Picture, OutputOffsetX, OutputOffsetY, OutputSizeX, OutputSizeY, (X * InputSizeX) + InputOffsetX, (Y * InputSizeY) + InputOffsetY, InputSizeX, InputSizeY
            SavePicture picOutput.Image, GetParentFolderName(txtFileName.Text) & "\tile" & Format$(X, "00") & Format$(Y, "00") & ".bmp"
        Next Y
    Next X
End Sub

Private Sub Form_Load()
    LockAll
    Load frmInput
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnlockAll
    Unload frmInput
End Sub

Private Sub picOutput_Click()
    On Error Resume Next
    
    With cdlMain
        .Flags = cdlCCRGBInit
        .Color = picOutput.BackColor
        .ShowColor
        
        If Err.Number <> cdlCancel Then picOutput.BackColor = .Color
    End With
End Sub

Private Sub txtGridSizeX_Change()
    Update
End Sub

Private Sub txtGridSizeY_Change()
    Update
End Sub

Private Sub txtInputOffsetX_Change()
    Update
End Sub

Private Sub txtInputOffsetY_Change()
    Update
End Sub

Private Sub txtInputSizeX_Change()
    Update
End Sub

Private Sub txtInputSizeY_Change()
    Update
End Sub

Private Sub txtOutputSizeX_Change()
    Update
End Sub

Private Sub txtOutputSizeY_Change()
    Update
End Sub
