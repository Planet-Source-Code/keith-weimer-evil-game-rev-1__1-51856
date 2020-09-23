VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Evil Game Editor"
   ClientHeight    =   5955
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   8385
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrDisplay 
      Interval        =   100
      Left            =   1920
      Top             =   600
   End
   Begin VB.Timer tmrMenu 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   600
   End
   Begin MSComctlLib.ImageList ilsToolbar 
      Left            =   600
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "egm"
      MaxFileSize     =   32767
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ilsToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Map"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Map"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Map"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Map..."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Map..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Map"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save Map As..."
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenTileset 
         Caption         =   "Open Tileset..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpenAnimations 
         Caption         =   "Open Animations..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Select Mode"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Edit Mode"
         Checked         =   -1  'True
         Index           =   1
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewInfo 
         Caption         =   "Info"
      End
      Begin VB.Menu mnuViewBrush 
         Caption         =   "Brush"
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHighlightSolids 
         Caption         =   "Highlight Solids"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuResourceEditor 
         Caption         =   "Resource Editor"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTileRipper 
         Caption         =   "Tile Ripper"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Cascade"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Tile Horizontally"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Tile Vertically"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Arrange Icons"
         Enabled         =   0   'False
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ActiveEditor As frmEditor

Function GetEditorCount() As Integer
    'On Error Resume Next
    
    Dim Form As Form
    
    For Each Form In Forms
        If TypeOf Form Is frmEditor Then GetEditorCount = GetEditorCount + 1
    Next Form
End Function

Sub DisplayAll()
    'On Error Resume Next
    
    Dim Form As Form
    
    For Each Form In Forms
        If TypeOf Form Is frmEditor Then Form.Display
    Next Form
End Sub

Private Sub MDIForm_Load()
    'On Error Resume Next
    
    LockAll
    
    frmInfo.Show , frmMain
    frmBrush.Show , frmMain
    frmTilePicker.Show , frmMain
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'On Error Resume Next
    
    UnlockAll
End Sub

Private Sub mnuClose_Click()
    'On Error Resume Next
    
    Unload ActiveEditor
End Sub

Private Sub mnuExit_Click()
    'On Error Resume Next
    
    Unload frmMain
End Sub

Private Sub mnuNew_Click()
    'On Error Resume Next
    
    frmAttributes.Show vbModal, frmMain
    
    If IsLoaded(frmAttributes) Then
        Static Count As Integer
        
        Dim Width As Integer
        Dim Height As Integer
        
        With frmAttributes
            Width = Val(.txtWidth.Text)
            Height = Val(.txtHeight.Text)
        End With
        Unload frmAttributes
        
        Count = Count + 1
        
        Dim Editor As New frmEditor
        Editor.NewMap "Map" & Format$(Count, "00000") & ".egm", Width, Height
        Editor.Show
    End If
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    
    With cdlMain
        .Filter = MapFilter
        .Flags = OpenFlags
        .FileName = Empty
        .ShowOpen
        
        If Err.Number <> cdlCancel Then
            Dim Editor As New frmEditor
            Editor.OpenMap .FileName
            Editor.Show
        End If
    End With
End Sub

Sub mnuSave_Click()
    'On Error Resume Next
    
    If ActiveEditor.Saved Then
        ActiveEditor.SaveMap
    Else
        mnuSaveAs_Click
    End If
End Sub

Private Sub mnuSaveAs_Click()
    On Error Resume Next
    
    With cdlMain
        .Filter = MapFilter
        .Flags = SaveFlags
        .FileName = Empty
        .ShowSave
        
        If Err.Number <> cdlCancel Then ActiveEditor.SaveMap .FileName
    End With
End Sub

Private Sub mnuViewBrush_Click()
    'On Error Resume Next
    
    frmBrush.Visible = Not frmBrush.Visible
    If frmBrush.Visible Then frmBrush.Show , frmMain
End Sub

Private Sub mnuViewHighlightSolids_Click()
    'On Error Resume Next
    
    mnuViewHighlightSolids.Checked = Not mnuViewHighlightSolids.Checked
    
    DisplayAll
End Sub

Private Sub mnuViewInfo_Click()
    'On Error Resume Next
    
    frmInfo.Visible = Not frmInfo.Visible
    If frmInfo.Visible Then frmInfo.Show , frmMain
End Sub

Private Sub mnuWindowArrange_Click(Index As Integer)
    'On Error Resume Next
    
    Arrange Index
End Sub

Private Sub tmrDisplay_Timer()
    'On Error Resume Next
    
    Animate Map.Animations
    
    frmBrush.Display
    frmTilePicker.Display
End Sub

Private Sub tmrMenu_Timer()
    'On Error Resume Next
    
    Dim EditorLoaded As Boolean
    Dim Index As Integer
    
    EditorLoaded = GetEditorCount > 0
    
    mnuSave.Enabled = EditorLoaded
    mnuSaveAs.Enabled = EditorLoaded
    mnuClose.Enabled = EditorLoaded
    'mnuOpenTileset.Enabled = EditorLoaded
    'mnuOpenAnimations.Enabled = EditorLoaded
    
    'For Index = mnuMode.LBound To mnuMode.UBound
        mnuMode(1).Enabled = EditorLoaded
    'Next Index
    
    For Index = mnuWindowArrange.LBound To mnuWindowArrange.UBound
        mnuWindowArrange(Index).Enabled = EditorLoaded
    Next Index
    
    mnuViewInfo.Checked = frmInfo.Visible
    mnuViewBrush.Checked = frmBrush.Visible
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    'On Error Resume Next
    
    Select Case Button.Key
        Case "New": mnuNew_Click
        Case "Open": mnuOpen_Click
        Case "Save": mnuSave_Click
    End Select
End Sub
