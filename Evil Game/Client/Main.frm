VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evil Game"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDisplay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7320
      Top             =   2160
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Login"
      Height          =   1455
      Left            =   1920
      TabIndex        =   10
      Top             =   2880
      Width           =   2655
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Text            =   "Player"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblUsername 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   9
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtCommand 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   8280
      Width           =   6495
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   7200
      Width           =   7215
   End
   Begin VB.Frame fraConnect 
      Caption         =   "Connect"
      Height          =   1455
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Text            =   "2468"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Text            =   "localhost"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   330
      End
      Begin VB.Label lblHost 
         AutoSize        =   -1  'True
         Caption         =   "Host:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer tmrInfo 
      Interval        =   1
      Left            =   7320
      Top             =   1680
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents Client As Connection
Attribute Client.VB_VarHelpID = -1

Dim Map As Map
Dim Player As Player
Dim Players() As Player

Dim BackBuffer As Surface

Dim Buffer As String
Dim DataLength As Long

Dim LoggedIn As Boolean

Sub AddLogText(ByVal Text As String, Optional NewLine As Boolean = True)
    'On Error Resume Next
    
    If NewLine Then Text = Text & vbNewLine
    
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelText = Text
End Sub

Function GetPlayerCount() As Integer
    On Error Resume Next
    
    GetPlayerCount = UBound(Players) - LBound(Players) + 1
End Function

Sub Display()
    'On Error Resume Next
    
    If GetPlayerCount > 0 Then
        Dim Index As Integer
        
        ClearSurface BackBuffer
        RenderMap BackBuffer, Map, 7, 7
        
        For Index = LBound(Players) To UBound(Players)
            Players(Index).Render BackBuffer, Map.Tilesets
        Next Index
        
        RenderSurface picDisplay.hDC, BackBuffer
    End If
End Sub

Sub Clear()
    'On Error Resume Next
    
    tmrDisplay.Enabled = False
    
    ClearMap Map
    ClearSurface BackBuffer
    RenderSurface picDisplay.hDC, BackBuffer
End Sub

Sub OpenClient(Optional Host As String, Optional Port As Integer)
    'On Error Resume Next
    
    Client.Connect Host, Port
End Sub

Private Sub Client_Connect()
    'On Error Resume Next
    
    AddLogText "*** Connected to server."
End Sub

Private Sub Client_Disconnect()
    'On Error Resume Next
    
    Clear
    LoggedIn = False
    
    AddLogText "*** Connection lost."
End Sub

Private Sub Client_Error(Number As Integer, Description As String)
    'On Error Resume Next
    
    AddLogText "!!! Error: " & Description
End Sub

Private Sub Client_ReceiveData(Data As String)
    'On Error Resume Next
    
    Dim Count As Integer
    Dim Index As Integer
    Dim ByteOffset As Long
    Dim Command As String
    
    Dim X As Long
    Dim Y As Long
    
    Command = Left$(Data, 4)
    Data = Right$(Data, Len(Data) - 4)
    Select Case Command
        Case "MESG" 'Message
            Dim Leader As String
            
            Select Case Asc(Mid$(Data, 1, 1)) 'Priority
                Case 0 'General
                    Leader = "***"
                Case 1 'Important
                    Leader = "%%%"
                Case 255 'Error
                    Leader = "!!!"
            End Select
            
            AddLogText Leader & " " & Mid$(Data, 2)
        Case "AUTH" 'Login Authorization
            If Len(Data) = 0 Then
                LoggedIn = True
                AddLogText "*** Logged into server"
            ElseIf Len(Data) = 1 Then
                Select Case Asc(Data)
                    Case 0 'Private
                        AddLogText "*** Login failed.  Private users only."
                    Case 1 'Username incorrect
                        AddLogText "*** Username incorrect."
                    Case 2 'Password incorrect
                        AddLogText "*** Password incorrect."
                    Case 3 'Banned
                        AddLogText "%%% Username is banned."
                End Select
            Else
                AddLogText "**** " & Data
            End If
        Case "CHAT" 'Chat
            Dim Username As String
            
            Username = RTrimNull(Mid$(Data, 1, 16))
            Data = Mid$(Data, 17)
            
            AddLogText Username & ": " & Data
        Case "MAPI" 'Map info
            If Len(Data) = 0 Then
                Clear
            Else
                Dim Arguments() As String
                                
                Arguments = Split(Data, vbNullChar)
                
                Map.Name = Arguments(0)
                Map.Tilesets.FileName = Arguments(1)
                Map.Animations.FileName = Arguments(2)
                
                Map.Tilesets = GetTilesets(Map.Tilesets.FileName)
                Map.Animations = GetAnimations(Map.Animations.FileName)
                
                ResizeMap Map, 15, 15
                
                tmrDisplay.Enabled = True
            End If
        Case "SRGN" 'Set regions
            Dim Region As Region
            
            Count = Len(Data) \ 10
            
            For Index = 0 To Count - 1
                ByteOffset = Index * 10
                
                X = StringToWord(Mid$(Data, ByteOffset + 1, 2))
                Y = StringToWord(Mid$(Data, ByteOffset + 3, 2))
                
                With Map.Region(X, Y)
                    .Tile.TilesetIndex = StringToWord(Mid$(Data, ByteOffset + 5, 2))
                    .Tile.Index = StringToWord(Mid$(Data, ByteOffset + 7, 2))
                    .AnimationIndex = StringToWord(Mid$(Data, ByteOffset + 9, 2))
                End With
            Next Index
        Case "PLAY" 'Set players
            Count = Len(Data) \ 26
            
            Erase Players
            If Count > 0 Then
                ReDim Players(Count - 1)
                
                For Index = 0 To Count - 1
                    ByteOffset = Index * 26
                    
                    Set Players(Index) = New Player
                    
                    Players(Index).Name = RTrimNull(Mid$(Data, ByteOffset + 1, 16))
                    Players(Index).X = StringToWord(Mid$(Data, ByteOffset + 17, 2))
                    Players(Index).Y = StringToWord(Mid$(Data, ByteOffset + 19, 2))
                    Players(Index).TilesetIndex = StringToWord(Mid$(Data, ByteOffset + 21, 2))
                    Players(Index).TileIndex = StringToWord(Mid$(Data, ByteOffset + 23, 2))
                    Players(Index).MaskIndex = StringToWord(Mid$(Data, ByteOffset + 25, 2))
                Next Index
            End If
        Case "DISP": Display
    End Select
End Sub

Private Sub cmdConnect_Click()
    'On Error Resume Next
    
    OpenClient txtHost.Text, Val(txtPort.Text)
End Sub

Private Sub cmdDisconnect_Click()
    'On Error Resume Next
    
    Client.Disconnect
End Sub

Private Sub cmdLogin_Click()
    'On Error Resume Next
    
    Client.SendData "AUTH" & txtUsername.Text & vbNullChar & txtPassword.Text & vbNullChar
End Sub

Private Sub cmdSend_Click()
    'On Error Resume Next
    
    Client.SendData "CHAT" & txtCommand.Text
    
    txtCommand.Text = Empty
    txtCommand.SetFocus
End Sub

Private Sub Form_Activate()
    'On Error Resume Next
    
    If txtHost.Visible Then txtHost.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    
    If KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown Then
        With Player
            Select Case KeyCode
                Case vbKeyLeft: .Movement = .Movement Or moveLeft
                Case vbKeyUp: .Movement = .Movement Or moveUp
                Case vbKeyRight: .Movement = .Movement Or moveRight
                Case vbKeyDown: .Movement = .Movement Or moveDown
            End Select
            
            Client.SendData "MOVE" & Chr$(.Movement)
        End With
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    
    If KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown Then
        With Player
            Select Case KeyCode
                Case vbKeyLeft: .Movement = .Movement And Not moveLeft
                Case vbKeyUp: .Movement = .Movement And Not moveUp
                Case vbKeyRight: .Movement = .Movement And Not moveRight
                Case vbKeyDown: .Movement = .Movement And Not moveDown
            End Select
            
            Client.SendData "MOVE" & Chr$(.Movement)
        End With
    End If
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    
    BackBuffer = CreateSurface(480, 480)
    
    AddLogText "Evil Game RPG Engine v" & App.Major & "." & App.Minor & " Rev " & App.Revision
    AddLogText String$(75, "-")
    
    Set Player = New Player
    Set Client = New Connection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next
    
    Set Client = Nothing
    DeleteSurface BackBuffer
End Sub

Private Sub picDisplay_Paint()
    'On Error Resume Next
    
    RenderSurface picDisplay.hDC, BackBuffer
End Sub

Private Sub tmrDisplay_Timer()
    'On Error Resume Next
    
    Animate Map.Animations
    Display
End Sub

Private Sub tmrInfo_Timer()
    'On Error Resume Next
    
    fraConnect.Visible = Not Client.IsOpen
    fraLogin.Visible = Client.IsOpen And Not LoggedIn
    cmdSend.Enabled = LoggedIn
End Sub

Private Sub txtLog_Change()
    'On Error Resume Next
    
    txtLog.SelStart = Len(txtLog.Text)
End Sub
