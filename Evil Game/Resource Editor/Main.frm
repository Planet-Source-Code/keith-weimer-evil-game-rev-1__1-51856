VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Resource Editor"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMenu 
      Interval        =   1
      Left            =   3360
      Top             =   480
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      MaxFileSize     =   8192
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFile_1 
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
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add..."
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
         Begin VB.Menu mnuRemoveSelection 
            Caption         =   "Selection"
         End
         Begin VB.Menu mnuRemoveInverseSelection 
            Caption         =   "Inverse Selection"
         End
         Begin VB.Menu mnuRemoveAll 
            Caption         =   "All"
         End
      End
      Begin VB.Menu mnuModify 
         Caption         =   "Modify..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Saved As Boolean
Public Modified As Boolean
Public Mode As EditorModeConstants

Dim LastClick As ListItem

Property Get FileName() As String
    'On Error Resume Next
    
    Select Case Mode
        Case emTileset: FileName = modMain.Tilesets.FileName
        Case emAnimation: FileName = modMain.Animations.FileName
    End Select
End Property

Sub ChangeMode(Mode As EditorModeConstants)
    'On Error Resume Next
    
    lvwMain.ListItems.Clear
    lvwMain.ColumnHeaders.Clear
    Select Case Mode
        Case emTileset
            lvwMain.ColumnHeaders.Add , , "#", 500
            lvwMain.ColumnHeaders.Add , , "Picture", 3000
            lvwMain.ColumnHeaders.Add , , "Tiles", 750
        Case emAnimation
            lvwMain.ColumnHeaders.Add , , "#", 500
            lvwMain.ColumnHeaders.Add , , "Frame", 750
            lvwMain.ColumnHeaders.Add , , "Frame(s)", 800
    End Select
    
    Me.Mode = Mode
End Sub

Sub NewFile(Mode As EditorModeConstants)
    'On Error Resume Next
    
    ChangeMode Mode
    
    Select Case Mode
        Case emTileset
            Dim Tilesets As Tilesets
            modMain.Tilesets = Tilesets
        Case emAnimation
            Dim Animations As Animations
            modMain.Animations = Animations
    End Select
    
    Saved = False
    Modified = False
End Sub

Sub OpenFile(ByVal FileName As String)
    'On Error Resume Next
    
    Select Case GetExtensionName(FileName)
        Case "egt"
            ChangeMode emTileset
            modMain.Tilesets = GetTilesets(FileName, False)
        Case "ega"
            ChangeMode emAnimation
            modMain.Animations = GetAnimations(FileName)
        Case Else
            Exit Sub
    End Select
    
    Rebuild
    Saved = True
    Modified = False
End Sub

Sub SaveFile(Optional ByVal FileName As String)
    On Error Resume Next
    
    Select Case Mode
        Case emTileset
            If FileName <> Empty Then modMain.Tilesets.FileName = FileName
            SaveTilesets modMain.Tilesets
        Case emAnimation
            If FileName <> Empty Then modMain.Animations.FileName = FileName
            SaveAnimations modMain.Animations
    End Select
    
    Saved = True
    Modified = False
End Sub

Function CloseFile() As VbMsgBoxResult
    'On Error Resume Next
    
    If (Not Saved Or Modified) And FileName <> Empty Then
        Dim Prompt As String
        
        Prompt = "'" & GetFileName(FileName) & "'"
        If Not Saved Then
            Prompt = Prompt & " has not been saved.  Save it now?"
        ElseIf Modified Then
            Prompt = Prompt & " has been modified.  Would you like to save the changes?"
        End If
        
        CloseFile = MsgBox(Prompt, vbQuestion Or vbYesNoCancel)
        If CloseFile = vbYes Then SaveFile
    Else
        CloseFile = vbIgnore
    End If
End Function

Sub Rebuild()
    'On Error Resume Next
    
    Dim ListItem As ListItem
    
    lvwMain.ListItems.Clear
    
    Select Case Mode
        Case emTileset
            Dim TilesetIndex As Long
            
            For TilesetIndex = 1 To GetTilesetCount(modMain.Tilesets)
                Set ListItem = lvwMain.ListItems.Add(, , TilesetIndex)
                ListItem.ListSubItems.Add , , modMain.Tilesets.Tileset(TilesetIndex).FileName
                ListItem.ListSubItems.Add , , GetTileCount(modMain.Tilesets.Tileset(TilesetIndex))
            Next TilesetIndex
        Case emAnimation
            Dim AnimationIndex As Long
            
            For AnimationIndex = 1 To GetAnimationCount(modMain.Animations)
                Set ListItem = lvwMain.ListItems.Add(, , AnimationIndex)
                ListItem.ListSubItems.Add , , modMain.Animations.Animation(AnimationIndex).FrameIndex
                ListItem.ListSubItems.Add , , GetFrameCount(modMain.Animations.Animation(AnimationIndex))
            Next AnimationIndex
    End Select
End Sub

Sub Clear()
    'On Error Resume Next
    
    Select Case Mode
        Case emTileset: ClearTilesets modMain.Tilesets
        Case emAnimation: ClearAnimations modMain.Animations
    End Select
    
    lvwMain.ListItems.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'On Error Resume Next
    
    Cancel = CloseFile = vbCancel
End Sub

Private Sub Form_Resize()
    'On Error Resume Next
    
    lvwMain.Width = ScaleWidth
    lvwMain.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next
    
    ClearTilesets modMain.Tilesets
End Sub

Private Sub lvwMain_DblClick()
    'On Error Resume Next
    
    mnuModify_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'On Error Resume Next
    
    Set LastClick = Item
End Sub

Private Sub mnuAdd_Click()
    On Error Resume Next
    
    Select Case Mode
        Case emTileset
            With cdlMain
                .Filter = PictureFilter
                .FilterIndex = 0
                .Flags = MultiOpenFlags
                .FileName = Empty
                .ShowOpen
                
                If Err.Number <> cdlCancel Then
                    Dim Tileset As Tileset
                    Dim File() As String
                    
                    File = Split(.FileName, Chr$(0))
                    If UBound(File) > 0 Then
                        Dim Path As String
                        Dim Index As Integer
                        
                        Path = AddSlash(File(0))
                        
                        For Index = 1 To UBound(File)
                            Tileset.FileName = Replace(Path & File(Index), GetResourcePath, Empty)
                            Tileset.Surface = CreateSurfaceFromFile(.FileName)
                            AddTileset modMain.Tilesets, Tileset
                        Next Index
                    Else
                        Tileset.FileName = Replace(.FileName, GetResourcePath, Empty)
                        Tileset.Surface = CreateSurfaceFromFile(.FileName)
                        AddTileset modMain.Tilesets, Tileset
                    End If
                    
                    Rebuild
                    Modified = True
                End If
            End With
        Case emAnimation
            Dim Count As Long
            Dim Animation As Animation
            Dim AnimationIndex As Long
            
            Count = Val(InputBox$("Enter number of animations:", , "1"))
            
            If Count > 0 Then
                For AnimationIndex = 1 To Count
                    AddAnimation modMain.Animations, Animation
                Next AnimationIndex
                
                Rebuild
                Modified = True
            End If
    End Select
End Sub

Private Sub mnuExit_Click()
    'On Error Resume Next
    
    Unload Me
End Sub

Private Sub mnuModify_Click()
    On Error Resume Next
    
    frmEdit.Activate lvwMain.SelectedItem.Index
    
    Rebuild
    Modified = True
End Sub

Private Sub mnuNew_Click()
    'On Error Resume Next
    
    If CloseFile <> vbCancel Then frmNew.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    
    If CloseFile <> vbCancel Then
        With cdlMain
            .Filter = ResourceFilter
            .FilterIndex = 0
            .Flags = OpenFlags
            .FileName = Empty
            .ShowOpen
            
            If Err.Number <> cdlCancel Then OpenFile .FileName
        End With
    End If
End Sub

Private Sub mnuRemoveAll_Click()
    'On Error Resume Next
    
    Clear
    Modified = True
End Sub

Private Sub mnuRemoveInverseSelection_Click()
    'On Error Resume Next
    
    Dim Index As Integer
    
    For Index = lvwMain.ListItems.Count To 1 Step -1
        If Not lvwMain.ListItems(Index).Selected Then
            Select Case Mode
                Case emTileset: RemoveTileset modMain.Tilesets, Index
                Case emAnimation: RemoveAnimation modMain.Animations, Index
            End Select
            
            Modified = True
        End If
    Next Index
    
    If Modified Then Rebuild
End Sub

Private Sub mnuRemoveSelection_Click()
    'On Error Resume Next
    
    Dim Index As Integer
    
    For Index = lvwMain.ListItems.Count To 1 Step -1
        If lvwMain.ListItems(Index).Selected Then
            Select Case Mode
                Case emTileset: RemoveTileset modMain.Tilesets, Index
                Case emAnimation: RemoveAnimation modMain.Animations, Index
            End Select
            
            Modified = True
        End If
    Next Index
    
    If Modified Then Rebuild
End Sub

Private Sub mnuSave_Click()
    'On Error Resume Next
    
    If Saved Then
        SaveFile
    Else
        mnuSaveAs_Click
    End If
End Sub

Private Sub mnuSaveAs_Click()
    On Error Resume Next
    
    With cdlMain
        Select Case Mode
            Case emTileset: .Filter = TilesetsFilter
            Case emAnimation: .Filter = AnimationsFilter
        End Select
        
        .Flags = SaveFlags
        .FileName = Empty
        .ShowSave
        
        If Err.Number <> cdlCancel Then SaveFile .FileName
    End With
End Sub

Private Sub tmrMenu_Timer()
    'On Error Resume Next
    
    mnuSave.Enabled = Mode <> emNone
    mnuSaveAs.Enabled = Mode <> emNone
    mnuAdd.Enabled = Mode <> emNone
    mnuRemove.Enabled = lvwMain.ListItems.Count > 0
    mnuModify.Enabled = lvwMain.ListItems.Count > 0
End Sub
