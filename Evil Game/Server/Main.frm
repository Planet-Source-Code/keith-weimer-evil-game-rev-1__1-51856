VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Evil Game Server"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   HasDC           =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFullClose 
      BackColor       =   &H000000FF&
      Caption         =   "Full Close"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtConsoleInput 
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
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   7575
   End
   Begin VB.Timer tmrInfo 
      Interval        =   1
      Left            =   7800
      Top             =   2160
   End
   Begin VB.Timer tmrDisplay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7800
      Top             =   2640
   End
   Begin VB.TextBox txtConsoleOutput 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   7575
   End
   Begin VB.Label lblConsole 
      AutoSize        =   -1  'True
      Caption         =   "Console:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblPlayers 
      AutoSize        =   -1  'True
      Caption         =   "Player(s): 0"
      Height          =   195
      Left            =   7920
      TabIndex        =   7
      Top             =   1920
      Width           =   780
   End
   Begin VB.Label lblSockets 
      AutoSize        =   -1  'True
      Caption         =   "Socket(s): 0"
      Height          =   195
      Left            =   7920
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim Map As Map

Dim WithEvents Server As Connection
Attribute Server.VB_VarHelpID = -1
Dim WithEvents Connections As Connections
Attribute Connections.VB_VarHelpID = -1
Dim Players As Players

Sub AddConsoleText(ByVal Text As String, Optional NewLine As Boolean = True)
    'On Error Resume Next
    
    If NewLine Then Text = Text & vbNewLine
    
    txtConsoleOutput.SelStart = Len(txtConsoleOutput.Text)
    txtConsoleOutput.SelText = Text
End Sub

Sub RunCommand(ByVal Command As String)
    'On Error Resume Next
    
    Dim Argument() As String
    
    AddConsoleText "> " & Command
    
    Argument = Split(Command, " ")
    
    Select Case Argument(0)
        Case "OPEN": OpenServer
        Case "CLOSE": CloseServer
        Case "PORT"
            If UBound(Argument) = 1 Then
                Server.Disconnect
                Server.LocalPort = Val(Argument(1))
            End If
        Case "MAP"
            If UBound(Argument) >= 1 Then
                OpenMap Argument(1)
            Else
                ClearMap Map
                AddConsoleText "Map cleared."
            End If
        Case "LIST"
            If UBound(Argument) >= 1 Then
                Select Case Argument(1)
                    Case "PLAYERS"
                        Dim PlayerHandle As Long
                        Dim Player As Player
                        Dim Temp As String
                        
                        If Players.Count = 0 Then
                            AddConsoleText "No players in system."
                        Else
                            For PlayerHandle = 1 To Players.Count
                                Set Player = Players(PlayerHandle)
                                
                                Temp = Temp & PlayerHandle & "-" & Player.Name & " (" & Player.X & ", " & Player.Y & ")" & vbNewLine
                            Next PlayerHandle
                            
                            AddConsoleText Temp, False
                        End If
                End Select
            End If
        Case "SWAP"
            If UBound(Argument) >= 1 Then
                Select Case Argument(1)
                    Case "PLAYERS"
                        If UBound(Argument) = 3 Then
                            If IsNumeric(Argument(2)) And IsNumeric(Argument(3)) Then
                                Dim ConnectionHandle As Long
                                
                                ConnectionHandle = Players(Val(Argument(2))).ConnectionHandle
                                Players(Val(Argument(2))).ConnectionHandle = Players(Val(Argument(3))).ConnectionHandle
                                Players(Val(Argument(3))).ConnectionHandle = ConnectionHandle
                            Else
                                AddConsoleText "2 numeric arugments required."
                            End If
                        Else
                            AddConsoleText "2 numeric arguments required."
                        End If
                End Select
            End If
    End Select
End Sub

Sub RunBatch(ByVal FileName As String)
    'On Error Resume Next
        
    If FileExists(FileName) Then
        Dim FileNum As Integer
        Dim Command As String
        
        FileNum = FreeFile(1)
        Open FileName For Input As #FileNum
            Do Until EOF(FileNum)
                Input #FileNum, Command
                RunCommand Command
            Loop
        Close #FileNum
    End If
End Sub

Sub OpenMap(ByVal FileName As String)
    'On Error Resume Next
    
    Dim PlayerHandle As Long
    
    AddConsoleText "Open map: " & FileName
    
    AddConsoleText Space$(4) & "Map................", False
    Map = GetMap(FileName)
    AddConsoleText "Complete."
        
    AddConsoleText Space$(4) & "Updating players...", False
    For PlayerHandle = 1 To Players.Count
        SendMapInfo PlayerHandle
    Next PlayerHandle
    AddConsoleText "Complete."
    
    tmrDisplay.Enabled = True
End Sub

Sub SendMapInfo(ByVal PlayerHandle As Long)
    'On Error Resume Next
        
    SendData PlayerHandle, "MAPI" & Map.Name & vbNullChar & Map.Tilesets.FileName & vbNullChar & Map.Animations.FileName & vbNullChar
End Sub

Function PlayerInBounds(ByVal BasePlayerHandle As Long, ByVal PlayerHandle As Long) As Boolean
    'On Error Resume Next
    
    Dim BasePlayer As Player
    Dim Player As Player
    
    Set BasePlayer = Players(BasePlayerHandle)
    Set Player = Players(PlayerHandle)
    
    If Not Player Is Nothing Then PlayerInBounds = Player.X >= BasePlayer.X - 7 And Player.Y >= BasePlayer.Y - 7 And Player.X <= BasePlayer.X + 7 And Player.Y <= BasePlayer.Y + 7
End Function

Sub SendRegions(PlayerHandle As Long)
    'On Error Resume Next
    
    Dim Player As Player
    
    Dim DisplayX As Integer
    Dim DisplayY As Integer
    Dim X As Integer
    Dim Y As Integer
    
    Dim RegionData As String
    
    Set Player = Players(PlayerHandle)
    
    For DisplayX = 0 To 14
        X = DisplayX + Player.X - 7
        For DisplayY = 0 To 14
            Y = DisplayY + Player.Y - 7
            
            RegionData = RegionData & WordToString(DisplayX)
            RegionData = RegionData & WordToString(DisplayY)
            
            If InMapBounds(Map, X, Y) Then
                With Map.Region(X, Y)
                    RegionData = RegionData & WordToString(.Tile.TilesetIndex)
                    RegionData = RegionData & WordToString(.Tile.Index)
                    RegionData = RegionData & WordToString(.AnimationIndex)
                End With
            Else
                With Map.OuterRegion
                    RegionData = RegionData & WordToString(.Tile.TilesetIndex)
                    RegionData = RegionData & WordToString(.Tile.Index)
                    RegionData = RegionData & WordToString(.AnimationIndex)
                End With
            End If
        Next DisplayY
    Next DisplayX
    
    SendData PlayerHandle, "SRGN" & RegionData
End Sub

Sub SendPlayers(ByVal PlayerHandle As Long)
    'On Error Resume Next
    
    Dim Player As Player
    Dim Index As Long
    Dim Data As String
    Dim PlayersData As String
    
    Set Player = Players(PlayerHandle)
        
    With Player
        For Index = 1 To Players.Count
            If PlayerInBounds(PlayerHandle, Index) Then
                With Players(Index)
                    Data = FixedLengthString(16, .Name)
                    Data = Data & WordToString(.X - Player.X + 7)
                    Data = Data & WordToString(.Y - Player.Y + 7)
                    Data = Data & WordToString(.TilesetIndex)
                    Data = Data & WordToString(.TileIndex)
                    Data = Data & WordToString(.MaskIndex)
                End With
                
                If Index = PlayerHandle Then
                    PlayersData = Data & PlayersData
                Else
                    PlayersData = PlayersData & Data
                End If
            End If
        Next Index
    End With
    
    SendData PlayerHandle, "PLAY" & PlayersData
End Sub

Sub Display(PlayerHandle As Long)
    'On Error Resume Next
    
    SendData PlayerHandle, "DISP"
End Sub

Sub OpenServer()
    'On Error Resume Next
    
    Server.Listen
End Sub

Sub CloseServer(Optional ByVal Full As Boolean = False)
    'On Error Resume Next
    
    Server.Disconnect
    
    If Full Then
        Dim ConnectionHandle As Long
        
        For ConnectionHandle = 1 To Connections.Count
            If Not Connections(ConnectionHandle) Is Nothing Then
                Connections(ConnectionHandle).Disconnect
            End If
        Next ConnectionHandle
    End If
End Sub

Sub SendData(PlayerHandle As Long, Data As String)
    'On Error Resume Next
    
    If PlayerHandle <> 0 Then
        If Not Players(PlayerHandle) Is Nothing Then
            Dim ConnectionHandle As Long
            
            ConnectionHandle = Players(PlayerHandle).ConnectionHandle
            If ConnectionHandle <> 0 Then
                If Not Connections(ConnectionHandle) Is Nothing Then
                    Connections(ConnectionHandle).SendData Data
                End If
            End If
        End If
    End If
End Sub

Sub SendDataAll(Data As String)
    'On Error Resume Next
    
    Dim PlayerHandle As Long
    
    For PlayerHandle = 1 To Players.Count
        SendData PlayerHandle, Data
    Next PlayerHandle
End Sub

Private Sub cmdClose_Click()
    'On Error Resume Next
    
    CloseServer
End Sub

Private Sub cmdFullClose_Click()
    'On Error Resume Next
    
    CloseServer True
End Sub

Private Sub cmdOpen_Click()
    'On Error Resume Next
    
    OpenServer
End Sub

Private Sub Connections_Disconnect(Index As Long)
    'On Error Resume Next
    
    If Connections(Index).ExtraHandle <> 0 Then
        Players.Remove Connections(Index).ExtraHandle
    End If
End Sub

Private Sub Connections_ReceiveData(Index As Long, Data As String)
    'On Error Resume Next
    
    Dim PlayerHandle As Long
    Dim Player As Player
    
    If Connections(Index).ExtraHandle <> 0 Then
        PlayerHandle = Connections(Index).ExtraHandle
        Set Player = Players(PlayerHandle)
    End If
    
    Dim Command As String
    
    Command = Left$(Data, 4)
    Data = Right$(Data, Len(Data) - 4)
    Select Case Command
        Case "MOVE"
            Player.Movement = Asc(Data)
        Case "CHAT"
            SendDataAll "CHAT" & FixedLengthString(16, Player.Name) & Data
        Case "AUTH"
            Dim Username As String
            Dim Password As String
            
            If Len(Data) > 0 Then
                Dim Arg() As String
                
                Arg = Split(Data, vbNullChar)
                If UBound(Arg) > 0 Then
                    Username = Arg(0)
                    Password = Arg(1)
                Else
                    Username = Arg(0)
                End If
            End If
                        
            If Username = Empty Then
                Connections(Index).SendData "AUTH" & Chr$(1)
                DoEvents
                Connections(Index).Disconnect
            Else
                '*****************************
                'Put authorization system here
                '*****************************
                
                Set Player = New Player
                
                Player.ConnectionHandle = Index
                Player.Name = Username
                Player.X = 3
                Player.Y = 3
                Player.TilesetIndex = 2
                Player.TileIndex = 1
                Player.MaskIndex = 2
                
                PlayerHandle = Players.Add(Player)
                Connections(Index).ExtraHandle = PlayerHandle
                Connections(Index).SendData "AUTH"
                
                SendMapInfo PlayerHandle
            End If
    End Select
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    
    AddConsoleText App.Title & " v" & App.Major & "." & App.Minor
    AddConsoleText String$(80, "-")
    
    Show
    Refresh
    
    Set Server = New Connection
    Set Connections = New Connections
    Set Players = New Players
    
    Server.LocalPort = 2468
    
    RunBatch GetFullPathName(App.Path & "\Autoexec.egb")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next
    
    CloseServer True
End Sub

Private Sub Server_ConnectionRequest(Connection As Connection)
    'On Error Resume Next
    
    Set Connection = New Connection
    
    Connections.Add Connection
End Sub

Private Sub tmrDisplay_Timer()
    'On Error Resume Next
    
    Dim PlayerHandle As Long
    Dim Player As Player
    
    Dim X As Integer
    Dim Y As Integer
    
    For PlayerHandle = 1 To Players.Count
        Set Player = Players(PlayerHandle)
        
        If Not Player Is Nothing Then
            X = 0
            Y = 0
                
            If GetFlag(Player.Movement, moveLeft) Then X = X - 1
            If GetFlag(Player.Movement, moveUp) Then Y = Y - 1
            If GetFlag(Player.Movement, moveRight) Then X = X + 1
            If GetFlag(Player.Movement, moveDown) Then Y = Y + 1
            
            If X Then
                If Player.X + X >= 0 And Player.X + X < Map.Width Then
                    If Not Map.Region(Player.X + X, Player.Y).Flags And rgnSolid Then Player.X = Player.X + X
                End If
            End If
            
            If Y Then
                If Player.Y + Y >= 0 And Player.Y + Y < Map.Height Then
                    If Not Map.Region(Player.X, Player.Y + Y).Flags And rgnSolid Then Player.Y = Player.Y + Y
                End If
            End If
        
            SendRegions PlayerHandle
            SendPlayers PlayerHandle
            Display PlayerHandle
        End If
        
        DoEvents
    Next PlayerHandle
End Sub

Private Sub tmrInfo_Timer()
    'On Error Resume Next
    
    cmdOpen.Enabled = Server.State <> sckListening
    cmdClose.Enabled = Server.State = sckListening
    cmdFullClose.Enabled = Connections.Count > 0
    
    lblSockets.Caption = "Socket(s): " & Connections.OpenCount
    lblPlayers.Caption = "Player(s): " & Players.OpenCount
End Sub

Private Sub txtConsoleInput_KeyPress(KeyAscii As Integer)
    'On Error Resume Next
    
    If KeyAscii = 13 Then
        RunCommand txtConsoleInput.Text
        
        txtConsoleInput.Text = Empty
        KeyAscii = 0
    End If
End Sub

Private Sub txtConsoleOutput_Change()
    'On Error Resume Next
    
    txtConsoleOutput.SelStart = Len(txtConsoleOutput.Text)
End Sub
