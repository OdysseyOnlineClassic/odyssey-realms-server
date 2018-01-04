VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "The Odyssey Classic Server"
   ClientHeight    =   1605
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCloseScks 
      Interval        =   1000
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer PlayerTimer 
      Interval        =   2000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer MinuteTimer 
      Interval        =   60000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer MapTimer 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox lstLog 
      Height          =   1230
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu menuShutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu mnuRelayChat 
         Caption         =   "Relay Chat"
      End
      Begin VB.Menu mnuLogScripts 
         Caption         =   "Log Scripts"
      End
      Begin VB.Menu mnuServerOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu cmdGetPasswordUser 
         Caption         =   "Get Password (user)"
      End
      Begin VB.Menu cmdGetPasswordName 
         Caption         =   "Get Password (name)"
      End
      Begin VB.Menu cmdReportsScripts 
         Caption         =   "Scripts"
      End
      Begin VB.Menu cmdReportsObjectUsage 
         Caption         =   "Object Usage"
      End
      Begin VB.Menu cmdReportsWarpToMap 
         Caption         =   "Warps to Map"
      End
      Begin VB.Menu cmdReportsWarps 
         Caption         =   "All Warps"
      End
      Begin VB.Menu cmdReportsMapSpawns 
         Caption         =   "Map Object Spawns"
      End
      Begin VB.Menu mnuReportsUnlinked 
         Caption         =   "Unlinked Maps"
      End
      Begin VB.Menu mnuReportsMaps 
         Caption         =   "&Free Maps"
      End
      Begin VB.Menu cmdReportsNPCUsage 
         Caption         =   "NPC Usage"
      End
      Begin VB.Menu cmdReportsMonsterUsage 
         Caption         =   "Monster Usage"
      End
      Begin VB.Menu cmdReportsKeepTiles 
         Caption         =   "Keep Tile Objects"
      End
      Begin VB.Menu cmdReportsAboveLevel 
         Caption         =   "Above Level"
      End
      Begin VB.Menu mnuGold 
         Caption         =   "Gold"
      End
      Begin VB.Menu mnuReportsObjects 
         Caption         =   "Object List"
      End
      Begin VB.Menu mnuLastScript 
         Caption         =   "Last Script"
      End
      Begin VB.Menu cmdReportsHasObject 
         Caption         =   "Has Object"
      End
      Begin VB.Menu mnuReportsGods 
         Caption         =   "&Gods"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu cmdConvert 
         Caption         =   "Convert"
      End
      Begin VB.Menu mnuDatabaseResetItem 
         Caption         =   "Erase Object"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
         Begin VB.Menu mnuDatabaseEmptyNPCs 
            Caption         =   "Empty NPCs"
         End
         Begin VB.Menu mnuDatabaseEmptyObjects 
            Caption         =   "Empty Objects"
         End
         Begin VB.Menu mnuDatabaseEmptyMonsters 
            Caption         =   "Empty Monsters"
         End
         Begin VB.Menu mnuDatabaseDelMaps 
            Caption         =   "Empty Maps"
         End
         Begin VB.Menu cmdResetFlags 
            Caption         =   "Player Flags"
         End
         Begin VB.Menu mnuDatabaseResetAccounts 
            Caption         =   "Accounts"
         End
         Begin VB.Menu mnuDatabaseResetObjects 
            Caption         =   "Objects"
         End
         Begin VB.Menu mnuDatabaseResetMonsters 
            Caption         =   "Monsters"
         End
         Begin VB.Menu mnuDatabaseResetNPCs 
            Caption         =   "NPCs"
         End
         Begin VB.Menu mnuDatabaseResetInventory 
            Caption         =   "Inventories and Banks"
         End
         Begin VB.Menu mnuDatabaseResetKeep 
            Caption         =   "Keep Tiles"
         End
         Begin VB.Menu mnuDatabaseResetGuilds 
            Caption         =   "Guilds"
         End
         Begin VB.Menu mnuDatabaseResetGods 
            Caption         =   "Gods"
         End
         Begin VB.Menu mnuDatabaseResetMagic 
            Caption         =   "Magic"
         End
         Begin VB.Menu mnuDatabaseResetMoney 
            Caption         =   "Money (Gold)"
         End
         Begin VB.Menu mnuDatabaseResetBans 
            Caption         =   "Bans"
         End
         Begin VB.Menu mnuDatabaseResetBugReports 
            Caption         =   "Bug Reports"
         End
      End
   End
   Begin VB.Menu mnuAddGod 
      Caption         =   "Add InGame God"
   End
   Begin VB.Menu mnuConnections 
      Caption         =   "Connections"
      Begin VB.Menu cmdAcceptConnections 
         Caption         =   "Accept"
         Checked         =   -1  'True
      End
      Begin VB.Menu cmdCloseConnections 
         Caption         =   "Close All"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAcceptConnections_Click()
    If cmdAcceptConnections.Checked = True Then
        closesocket ListeningSocket
        ListeningSocket = INVALID_SOCKET
        cmdAcceptConnections.Checked = False
    Else
        cmdAcceptConnections.Checked = True

        ListeningSocket = ListenForConnect(World.ServerPort, gHW, 1025)
        If ListeningSocket = INVALID_SOCKET Then
            MsgBox "Unable to create listening socket!1", vbOKOnly + vbExclamation, TitleString
            EndWinsock
            Unhook
            End
        End If
        If SetSockLinger(ListeningSocket, 1, 0) = SOCKET_ERROR Then
            If ListeningSocket > 0 Then
                closesocket (ListeningSocket)
            End If
            MsgBox "Unable to create listening socket (linger)!", vbOKOnly + vbExclamation, TitleString
            EndWinsock
            Unhook
            End
        End If
        If setsockopt(ListeningSocket, IPPROTO_TCP, TCP_NODELAY, 1&, 4) <> 0 Then
            MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
            EndWinsock
            Unhook
            End
        End If
        'If setsockopt(ListeningSocket, SOL_SOCKET, SO_RCVBUF, 32768, 4) <> 0 Then
        '    MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        '    EndWinsock
        '    Unhook
        '    End
        'End If
        'If setsockopt(ListeningSocket, SOL_SOCKET, SO_SNDBUF, 32768, 4) <> 0 Then
        '    MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        '    EndWinsock
        '    Unhook
        '    End
        'End If
    End If
End Sub

Private Sub cmdCloseConnections_Click()
    Dim A As Long
    For A = 1 To MaxUsers
        If Player(A).InUse = True Then
            CloseClientSocket A
        End If
    Next A
End Sub

Private Sub cmdConvert_Click()
    Dim A As Long
    
    For A = 1 To MaxMaps
        ConvertMap A
    Next A
End Sub

Private Sub cmdGetPasswordName_Click()
    Open "report.txt" For Output As #1

    Dim A As String
    A = UCase$(InputBox$("Enter the name which you want to look up the password for ", "Enter Player Name"))

    Print #1, "*** Password Report (user) ***"
    Print #1, ""

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If Not IsNull(UserRS!Name) Then
                If StrCmp(UCase$(UserRS!Name), A) Then
                    Print #1, "User=" + UserRS!User + " Name=" + UserRS!Name + " Password=" + CStr(UserRS!Password)
                End If
            End If
            UserRS.MoveNext
        Wend
    End If

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdGetPasswordUser_Click()
    Open "report.txt" For Output As #1

    Dim A As String
    A = UCase$(InputBox$("Enter the user which you want to look up the password for ", "Enter Player Name"))

    Print #1, "*** Password Report (user) ***"
    Print #1, ""

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If StrCmp(UCase$(UserRS!User), A) Then
                Print #1, "User=" + UserRS!User + " Name=" + UserRS!Name + " Password=" + CStr(UserRS!Password)
            End If
            UserRS.MoveNext
        Wend
    End If
    
        If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If UserRS!Access > 0 Then
                Print #1, "User=" + UserRS!User + " Name=" + UserRS!Name + " Access=" + CStr(UserRS!Access)
            End If
            UserRS.MoveNext
        Wend
    End If

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsAboveLevel_Click()
    Open "report.txt" For Output As #1

    Dim A As Long, Count As Long
    A = Val(InputBox("What Level?"))

    Print #1, "*** Above Level Report ***"
    Print #1, ""

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If UserRS!Level > A Then
                Print #1, "User=" + UserRS!User + " Name=" + UserRS!Name + " Level=" + CStr(UserRS!Level)
                Count = Count + 1
            End If
            UserRS.MoveNext
        Wend
    End If

    Print #1, ""
    Print #1, "*** Count - " + CStr(Count) + " ***"

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsHasObject_Click()
    Open "report.txt" For Output As #1

    Dim GrandInvCount As Long, GrandBankCount As Long, GrandEquippedCount As Long, GrandCount As Long
    Dim TheObj As Long, A As Long
    TheObj = Val(InputBox("What object?"))
    If TheObj = 0 Then Exit Sub

    Print #1, "*** Has Object Report ***"
    Print #1, ""

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            Dim InvCount As Byte
            Dim EquippedCount As Byte
            Dim BankCount As Byte
            Dim Count As Long

            InvCount = 0
            EquippedCount = 0
            BankCount = 0
            Count = 0

            'Inventory Data
            For A = 1 To 20
                If TheObj = UserRS.Fields("InvObject" + CStr(A)) Then
                    InvCount = InvCount + 1
                End If
            Next A
            For A = 1 To 6
                If TheObj = UserRS.Fields("EquippedObject" + CStr(A)) Then
                    EquippedCount = EquippedCount + 1
                End If
            Next A

            'Item Bank
            For A = 0 To 29
                If TheObj = UserRS.Fields("BankObject" + CStr(A)) Then
                    BankCount = BankCount + 1
                End If
            Next A

            GrandInvCount = GrandInvCount + InvCount
            GrandBankCount = GrandBankCount + BankCount
            GrandEquippedCount = GrandEquippedCount + EquippedCount
            Count = InvCount + BankCount + EquippedCount
            GrandCount = GrandCount + Count

            If Count > 0 Then
                Print #1, "User=" + UserRS!User + " Name=" + UserRS!Name + " Level=" + CStr(UserRS!Level) + " Count=" + CStr(Count) + " InvCount=" + CStr(InvCount) + " BankCount=" + CStr(BankCount) + " EquippedCount=" + CStr(EquippedCount)
            End If

            UserRS.MoveNext
        Wend
    End If

    Dim KeepCount As Long, B As Long
    For A = 1 To MaxMaps
        With Map(A)
            If .Keep = True Then
                For B = 0 To MaxMapObjects
                    With .Object(B)
                        If .Object = TheObj Then
                            If Map(A).Tile(.X, .Y).Att = 5 Or Map(A).Tile(.X, .Y).Att2 = 5 Then
                                KeepCount = KeepCount + 1
                                Print #1, "Map " & A & " (" & Map(A).Name & ") " & Object(Map(A).Object(B).Object).Name
                            End If
                        End If
                    End With
                Next B
            End If
        End With
    Next A

    GrandCount = GrandCount + KeepCount

    Print #1, ""
    Print #1, "***TOTALS***"
    Print #1, "GrandTotalCount=" + CStr(GrandCount) + " GrandInvCount=" + CStr(GrandInvCount) + " GrandBankCount=" + CStr(GrandBankCount) + " GrandEquippedCount=" + CStr(GrandEquippedCount) + " GrandKeepCount=" + CStr(KeepCount)
    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsKeepTiles_Click()
    Open "report.txt" For Output As #1

    Print #1, "Keep Tiles"
    Print #1, ""

    Dim A As Long, B As Long

    For A = 1 To MaxMaps
        With Map(A)
            If .Keep = True Then
                For B = 0 To MaxMapObjects
                    With .Object(B)
                        If .Object > 0 Then
                            If Map(A).Tile(.X, .Y).Att = 5 Or Map(A).Tile(.X, .Y).Att2 = 5 Then
                                Print #1, "Map " & A & " (" & Map(A).Name & ") " & Object(Map(A).Object(B).Object).Name
                            End If
                        End If
                    End With
                Next B
            End If
        End With
    Next A

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsMapSpawns_Click()
    Open "report.txt" For Output As #1

    Print #1, "OBJECT SPAWNS"
    Print #1, ""

    Dim A As Long, X As Long, Y As Long
    For A = 1 To MaxMaps
        For Y = 0 To 11
            For X = 0 To 11
                With Map(A).Tile(X, Y)
                    If .Att = 7 Then
                        If Not Object(.AttData(0)).Name = "Flower" Then
                            Print #1, "Map " & A & " (" & Map(A).Name & ") " & Object(.AttData(0)).Name
                        End If
                    End If
                End With
            Next X
        Next Y
    Next A

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsMonsterUsage_Click()
    Dim TheMonster(1 To MaxTotalMonsters) As Boolean

    Dim A As Long, B As Long

    For A = 1 To MaxMaps
        For B = 0 To 9
            If Map(A).MonsterSpawn(B).Monster > 0 Then
                TheMonster(Map(A).MonsterSpawn(B).Monster) = True
            End If
        Next B
    Next A

    A = 0

    Open "report.txt" For Output As #1

    'Monsters
    Print #1, "MONSTERS NOT PLACED ON ANY MAP"
    Print #1, ""
    For A = 1 To MaxTotalMonsters
        If Monster(A).Name <> "" Then
            If ExamineBit(Monster(A).flags, 4) = True Then
                Print #1, Monster(A).Name & "  -  Monster #" & A
            End If
        End If
    Next A
    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsNPCUsage_Click()
    Dim TheNPC(1 To MaxNPCs) As Boolean

    Dim A As Long

    For A = 1 To MaxMaps
        If Map(A).NPC > 0 Then
            TheNPC(Map(A).NPC) = True
        End If
    Next A

    A = 0

    Open "report.txt" For Output As #1

    'NPCs
    Print #1, "NPC'S NOT PLACED ON ANY MAP"
    Print #1, ""
    For A = 1 To MaxNPCs
        If TheNPC(A) = False Then
            If NPC(A).Name <> "" Then
                Print #1, NPC(A).Name & "  -  NPC #" & A
            End If
        End If
    Next A
    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsObjectUsage_Click()
    Dim ItemUse(1 To MaxObjects) As Boolean

    Dim A As Long, B As Long, X As Long, Y As Long

    For A = 1 To MaxTotalMonsters
        If Monster(A).Name <> "" Then    'Monster in use
            For B = 0 To 2
                If Monster(A).Object(B) > 0 Then
                    ItemUse(Monster(A).Object(B)) = True
                End If
            Next B
        End If
    Next A

    A = 0
    B = 0

    For A = 1 To MaxNPCs
        If NPC(A).Name <> "" Then    'NPC in use
            For B = 0 To 9
                If NPC(A).SaleItem(B).GiveObject > 0 Then ItemUse(NPC(A).SaleItem(B).GiveObject) = True
                If NPC(A).SaleItem(B).TakeObject > 0 Then ItemUse(NPC(A).SaleItem(B).TakeObject) = True
            Next B
        End If
    Next A

    A = 0
    X = 0
    Y = 0
    For A = 1 To MaxMaps
        For Y = 0 To 11
            For X = 0 To 11
                With Map(A).Tile(X, Y)
                    If .Att = 7 Then
                        ItemUse(.AttData(1) * 256 + .AttData(0)) = True
                    End If
                End With
            Next X
        Next Y
    Next A

    A = 0
    B = 0

    Open "report.txt" For Output As #1

    A = 0
    Print #1, ""
    Print #1, "OBJECTS NOT IN USE"
    Print #1, ""
    For A = 1 To MaxObjects
        If ItemUse(A) = False Then
            If Object(A).Name <> "" Then
                Print #1, Object(A).Name & "  -  Object #" & A
            End If
        End If
    Next A
    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub
Private Sub cmdReportsScripts_Click()
    Open "report.txt" For Output As #1

    Dim Count As Long

    Print #1, "*** Scripts Report ***"
    Print #1, ""

    If ScriptRS.BOF = False Then
        ScriptRS.MoveFirst
        While ScriptRS.EOF = False
            Print #1, ""
            Print #1, "*******************************************************"
            Print #1, "*******************************************************"
            Print #1, "Name=" + ScriptRS!Name
            Print #1, ""
            Print #1, "Script Source="
            Print #1, ScriptRS!Source
            Print #1, "*******************************************************"
            Print #1, "*******************************************************"
            Print #1, ""
            Count = Count + 1
            ScriptRS.MoveNext
        Wend
    End If

    Print #1, ""
    Print #1, "*** Count - " + CStr(Count) + " ***"

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsWarps_Click()
    Open "report.txt" For Output As #1

    Print #1, "WARPS"
    Print #1, ""

    Dim A As Long, X As Long, Y As Long, D As Long
    For A = 1 To MaxMaps
        For Y = 0 To 11
            For X = 0 To 11
                With Map(A).Tile(X, Y)
                    If .Att = 2 Then
                        D = CLng(Map(A).Tile(X, Y).AttData(0)) * 256 + CLng(Map(A).Tile(X, Y).AttData(1))
                        Print #1, "Map " & A & " (" & Map(A).Name & ") X:" & X & " Y:" & Y & " warps to " & D & " X:" & Map(A).Tile(X, Y).AttData(2) & " Y:" & Map(A).Tile(X, Y).AttData(3) & " (" & Map(D).Name & ")"
                    End If
                End With
            Next X
        Next Y
    Next A

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdReportsWarpToMap_Click()
    Dim D As Long
    D = Val(InputBox("Which map?"))
    If D > 0 And D <= MaxMaps Then

        Open "report.txt" For Output As #1

        Print #1, "WARPS TO MAP " & D
        Print #1, ""

        Dim A As Long, X As Long, Y As Long
        For A = 1 To MaxMaps
            For Y = 0 To 11
                For X = 0 To 11
                    With Map(A).Tile(X, Y)
                        If .Att = 2 Then
                            If (CLng(Map(A).Tile(X, Y).AttData(0)) * 256 + CLng(Map(A).Tile(X, Y).AttData(1))) = D Then
                                Print #1, "Map " & A & " (" & Map(A).Name & ") X:" & X & " Y:" & Y & " warps to " & D & " (" & Map(D).Name & ")"
                            End If
                        End If
                    End With
                Next X
            Next Y
            If Map(A).BootLocation.Map = D Then
                Print #1, "Map " & A & " (" & Map(A).Name & ") X:" & X & " Y:" & Y & " warps to " & D & " (" & Map(D).Name & ")" & " - Boot Location"
            End If
        Next A

        Close #1

    End If

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub cmdResetFlags_Click()
    If MsgBox("Are you *sure* you wish to reset all the player's flags?", vbYesNo) = vbYes Then
        If UserRS.BOF = False Then
            UserRS.MoveFirst
            While UserRS.EOF = False
                UserRS.Edit
                UserRS!flags = vbNullString
                UserRS.Update
                UserRS.MoveNext
            Wend
        End If
    End If
End Sub

Private Sub Form_Load()
    gHW = Me.hWnd
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = 1 Then
        lstLog.Width = Me.ScaleWidth
        lstLog.Height = Me.ScaleHeight - txtMessage.Height
        txtMessage.Top = lstLog.Height
        txtMessage.Width = Me.ScaleWidth
    End If
End Sub

Private Sub MapTimer_Timer()
    On Error GoTo LogDatShit

    Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long, G As Long, H As Long
    Dim MapNum As Long, Tick As Currency
    Dim St1 As String

    Randomize

    Tick = getTime()

    For H = 1 To GetMaxUsers
        MapNum = Player(H).Map
        If Player(H).InUse = True Then
            If MapNum > 0 Then
                With Map(MapNum)
                    If Not .LastUpdate = Tick Then
                        .LastUpdate = Tick
                        If .NumPlayers > 0 Then
                            St1 = ""
                            For A = 0 To 9
                                If .Door(A).Att > 0 Then
                                    If Tick - .Door(A).T > 10000 Then
                                        .Tile(.Door(A).X, .Door(A).Y).Att = .Door(A).Att
                                        .Door(A).Att = 0
                                        St1 = St1 + DoubleChar(2) + Chr$(37) + Chr$(A)
                                    End If
                                End If
                            Next A

                            For A = 0 To MaxMapObjects
                                With .Object(A)
                                    If .Object > 0 Then
                                        If Map(MapNum).Tile(.X, .Y).Att <> 5 And Map(MapNum).Tile(.X, .Y).Att2 <> 5 And .TimeStamp > 1 Then
                                            If Tick - .TimeStamp >= World.ObjResetTime Then
                                                .Object = 0
                                                .Value = 0
                                                .ItemPrefix = 0
                                                .ItemSuffix = 0
                                                St1 = St1 + DoubleChar(2) + Chr$(15) + Chr$(A)
                                            End If
                                        End If
                                    End If
                                End With
                            Next A

                            'Move monsters
                            St1 = St1 + MoveMonsters(MapNum, Tick)
                            
                            If St1 <> "" Then
                                SendToMapRaw MapNum, St1
                            End If
                        End If
                    End If
                End With
            End If
        End If
    Next H

    Exit Sub

LogDatShit:
    SendToGods Chr$(56) + Chr$(7) + "WARNING:  Server Crashed on MapTimer   " & Err.Description & "  Source: " & Err.Source & "  Number: " & Err.number
    PrintLog "WARNING:  Server Crashed on MapTimer   " & Err.Description & "  Source: " & Err.Source & "  Number: " & Err.number
    PrintDebug "WARNING:  Server Crashed on MapTimer   " & Err.Description & "  Source: " & Err.Source & "  Number: " & Err.number

End Sub
Private Sub menuShutdown_Click()
    ShutdownServer
    Unload frmOptions
    Unload Me
End Sub

Private Sub MinuteTimer_Timer()
    On Error GoTo LogDatShit
    Dim A As Long, B As Long, C As Long

    If World.BackupInterval > 0 Then
        BackupCounter = BackupCounter + 1
        If BackupCounter >= World.BackupInterval Then
            BackupCounter = 0
            'Backup Server Data
            For A = 1 To MaxUsers
                If Player(A).Mode = modePlaying Then
                    SavePlayerData A
                    If Player(A).SpeedStrikes > 0 Then
                        Player(A).SpeedStrikes = Player(A).SpeedStrikes - 1
                    End If
                    If Len(Player(A).Email) < 2 Then
                        SendSocket A, Chr$(16) + Chr$(47)
                    End If
                    If Player(A).Access > 0 Then
                        'Server is being backed up message for gods
                        SendSocket A, Chr$(16) + Chr$(48)
                    End If
                End If
            Next A
            For A = 1 To MaxGuilds
                If Guild(A).UpdateFlag = True Then
                    If Guild(A).Name = vbNullString Then

                    Else
                       GuildRS.Seek "=", A
                       If GuildRS.NoMatch = False Then
                           GuildRS.Edit
                           GuildRS!Kills = Guild(A).Kills
                           GuildRS!Deaths = Guild(A).Deaths
                           For B = 0 To 19
                               GuildRS("MemberKills" + CStr(B)) = Guild(A).Member(B).Kills
                               GuildRS("MemberDeaths" + CStr(B)) = Guild(A).Member(B).Deaths
                           Next B
                           For B = 0 To DeclarationCount
                               GuildRS("DeclarationKills" + CStr(B)) = Guild(A).Declaration(B).Kills
                               GuildRS("DeclarationDeaths" + CStr(B)) = Guild(A).Declaration(B).Deaths
                           Next B
                           GuildRS.Update
                       End If
                    End If
                    Guild(A).UpdateFlag = False
                End If
            Next A
            SaveFlags
            SaveObjects
        End If
    End If

    If World.LastUpdate <> CLng(Date) Then

        'Backup Server Data
        For A = 1 To MaxUsers
            If Player(A).Mode = modePlaying Then
                SavePlayerData A
            End If
        Next A
        For A = 1 To MaxGuilds
            If Guild(A).UpdateFlag = True Then
                If Not Guild(A).Name = vbNullString Then
                    GuildRS.Seek "=", A
                    If GuildRS.NoMatch = False Then
                        GuildRS.Edit
                        GuildRS!Kills = Guild(A).Kills
                        GuildRS!Deaths = Guild(A).Deaths
                        For B = 0 To 19
                            GuildRS("MemberKills" + CStr(B)) = Guild(A).Member(B).Kills
                            GuildRS("MemberDeaths" + CStr(B)) = Guild(A).Member(B).Deaths
                        Next B
                        For B = 0 To DeclarationCount
                            GuildRS("DeclarationKills" + CStr(B)) = Guild(A).Declaration(B).Kills
                            GuildRS("DeclarationDeaths" + CStr(B)) = Guild(A).Declaration(B).Deaths
                        Next B
                        GuildRS.Update
                    End If
                    Guild(A).UpdateFlag = False
                End If
            End If
        Next A
        SaveFlags
        SaveObjects

        'A new day has dawned in the land of Odyssey
        SendAll Chr$(16) + Chr$(42)
        RunScript "DAYTIMER"
        World.LastUpdate = CLng(Date)
        DataRS.Edit
        DataRS!LastUpdate = World.LastUpdate
        DataRS.Update

        'Update Guilds
        For A = 1 To MaxGuilds
            With Guild(A)
                If .Name <> vbNullString Then
                    If .Bank < 0 And World.LastUpdate >= .DueDate Then
                        'Debt not payed, delete guild
                        DeleteGuild A, 0
                    ElseIf CountGuildMembers(A) < 1 Then
                        'Not enough members, guild deleted
                        DeleteGuild A, 1
                    Else
                        If .Bank >= 0 Then
                            C = 0
                        Else
                            C = 1
                        End If

                        'Pay bill
                        .Bank = .Bank - GetGuildUpkeep(A)
                        If C = 0 And .Bank < 0 Then
                            .DueDate = CLng(Date) + 2
                        End If
                        If .Bank >= 0 Then
                            SendToGuild A, Chr$(152) + QuadChar(.Bank) + QuadChar$(GetGuildUpkeep(A))
                        Else
                            SendToGuild A, Chr$(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                        End If
                        GuildRS.Seek "=", A
                        If GuildRS.NoMatch = False Then
                            GuildRS.Edit
                            GuildRS!Bank = .Bank
                            GuildRS!DueDate = .DueDate
                            GuildRS.Update
                        End If
                    End If
                End If
            End With
        Next A
    End If

    RunScript "MINUTETIMER"

    Exit Sub

LogDatShit:
    SendToGods Chr$(56) + Chr$(7) + "WARNING:  Server Crashed on MinuteTimer   " & Err.Description & "  Source: " & Err.Source & "  Number: " & Err.number
    PrintLog "WARNING:  Server Crashed on MinuteTimer   " & Err.Description & "  Source: " & Err.Source & "  Number: " & Err.number
    PrintDebug "WARNING:  Server Crashed on MinuteTimer   " & Err.Description & "  Source: " & Err.Source & "  Number: " & Err.number
End Sub

Private Sub mnuDatabaseDelMaps_Click()
    Dim A As Long, B As Long
    For A = 1 To MaxMaps
        MapRS.Seek "=", A
        If MapRS.NoMatch = False Then
            If Map(A).Name = "" Then
                MapRS.Delete
                B = B + 1
            End If
        End If
    Next A
    MsgBox B & " Maps Deleted."
End Sub

Private Sub mnuAddGod_Click()
    Dim A As Long, B As Long
    A = FindPlayer(InputBox$("Enter the name of a person ingame in which you would like to change their access: ", "Enter Player Name"))
    B = Val(InputBox$("Enter a numerical number between 0 and 3 in which the selected player's new access shall be (1 = Crowd Control (Boot, Ban, Warp), 2 = God (Scripting, Mapping, Editing), 3 = Super Admin (access to locked features): ", "Enter Access"))
    If A >= 1 And A <= MaxUsers And B >= 0 And B <= 3 Then
        With Player(A)
            .Access = B
            SendSocket A, Chr$(65) + Chr$(B)
            If .Access > 0 Then
                SendToMap .Map, Chr$(91) + Chr$(A) + Chr$(3)
                .Status = 3
                MsgBox "Added God Successfully!", vbExclamation + vbOKOnly, "Added Successfully"
            Else
                SendToMap .Map, Chr$(91) + Chr$(A) + Chr$(0)
                Player(A).Status = 0
                MsgBox "Successfully removed god account!", vbOKOnly, "Removed God"
            End If
        End With
    Else
        MsgBox "You have entered an invalid selection. Please choose 'Add InGame God' again supplying valid information!", vbExclamation + vbOKOnly, "Invalid Information"
    End If
End Sub

Private Sub mnuDatabaseEmptyMonsters_Click()
    Dim B As Long
    If Not MonsterRS.BOF Then
        MonsterRS.MoveFirst
        While Not MonsterRS.EOF
            If MonsterRS!Name = "" Then
                MonsterRS.Delete
                B = B + 1
            End If
            MonsterRS.MoveNext
        Wend
    End If

    MsgBox B & " Monsters Erased."
End Sub

Private Sub mnuDatabaseEmptyNPCs_Click()
    Dim B As Long
    If Not NPCRS.BOF Then
        NPCRS.MoveFirst
        While Not NPCRS.EOF
            If NPCRS!Name = "" Then
                NPCRS.Delete
                B = B + 1
            End If
            NPCRS.MoveNext
        Wend
    End If

    MsgBox B & " NPCs Erased."
End Sub

Private Sub mnuDatabaseEmptyObjects_Click()
    Dim B As Long
    If Not ObjectRS.BOF Then
        ObjectRS.MoveFirst
        While Not ObjectRS.EOF
            If ObjectRS!ObjName = "" Then
                ObjectRS.Delete
                B = B + 1
            End If
            ObjectRS.MoveNext
        Wend
    End If

    MsgBox B & " Objects Erased."
End Sub

Private Sub mnuDatabaseResetAccounts_Click()
    If MsgBox("Are you *sure* you wish to delete every account?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every account -- continue?", vbYesNo) = vbYes Then
            If Not UserRS.BOF Then
                UserRS.MoveFirst
                While Not UserRS.EOF
                    DeleteAccount
                    UserRS.MoveNext
                Wend
            End If
            UserRS.Close
            Set UserRS = Nothing
            DB.TableDefs.Delete "Accounts"
            CreateAccountsTable
            Set UserRS = DB.TableDefs("Accounts").OpenRecordset(dbOpenTable)
            UserRS.Index = "User"
        End If
    End If
End Sub

Private Sub mnuDatabaseResetBans_Click()
    Dim A As Long

    If MsgBox("Are you *sure* you wish to delete every Ban?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every Ban -- continue?", vbYesNo) = vbYes Then
            BanRS.Close
            Set BanRS = Nothing
            DB.TableDefs.Delete "Bans"
            CreateBansTable
            Set BanRS = DB.TableDefs("Bans").OpenRecordset(dbOpenTable)
            BanRS.Index = "Number"
            For A = 1 To 50
                With Ban(A)
                    If .Banner = "ReMoTech" Then

                    Else
                        .Banner = ""
                        .ComputerID = ""
                        .IPAddress = ""
                        .Reason = ""
                        .Name = ""
                        .UnbanDate = 0
                        .InUse = False
                    End If
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetBugReports_Click()
Dim A As Long
    If MsgBox("Are you *sure* you wish to delete every Bug Report?", vbYesNo) = vbYes Then
        If MsgBox("About to reset all Bug Reports -- continue?", vbYesNo) = vbYes Then
            BugsRS.Close
            Set BugsRS = Nothing
            DB.TableDefs.Delete "Bugs"
            CreateBugsTable
            Set BugsRS = DB.TableDefs("Bugs").OpenRecordset(dbOpenTable)
            BugsRS.Index = "ID"
            For A = 1 To 500
                With Bug(A)
                    .PlayerUser = ""
                    .PlayerName = ""
                    .PlayerIP = ""
                    .Title = ""
                    .Description = ""
                    .Status = 0
                    .ResolverUser = ""
                    .ResolverName = ""
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetGods_Click()
    If MsgBox("Are you *sure* you wish to delete every God?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every God -- continue?", vbYesNo) = vbYes Then
            If UserRS.BOF = False Then
                UserRS.MoveFirst
                While UserRS.EOF = False
                    If UserRS!Access > 0 Then
                        UserRS.Edit
                        UserRS!Access = 0
                        UserRS.Update
                    End If
                    UserRS.MoveNext
                Wend
            End If
        End If
    End If
End Sub

Private Sub mnuDatabaseResetGuilds_Click()
    Dim A As Long, B As Long

    If MsgBox("Are you *sure* you wish to delete every guild?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every guild -- continue?", vbYesNo) = vbYes Then
            GuildRS.Close
            Set GuildRS = Nothing
            DB.TableDefs.Delete "Guilds"
            CreateGuildsTable
            Set GuildRS = DB.TableDefs("Guilds").OpenRecordset(dbOpenTable)
            GuildRS.Index = "Number"
            For A = 1 To MaxGuilds
                With Guild(A)
                    .Bank = 0
                    .Bookmark = 0
                    .Name = ""
                    .Hall = 0
                    .DueDate = 0
                    .Sprite = 0

                    For B = 0 To DeclarationCount
                        With .Declaration(B)
                            .Guild = 0
                            .Type = 0
                        End With
                    Next B
                    For B = 0 To 19
                        With .Member(B)
                            .Name = ""
                            .Rank = 0
                        End With
                    Next B
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetInventory_Click()
    Dim A As Long
    If MsgBox("Are you *sure* you wish to delete peoples inventories?", vbYesNo) = vbYes Then
        If MsgBox("About to reset everyones inventory -- continue?", vbYesNo) = vbYes Then
            If Not UserRS.BOF Then
                UserRS.MoveFirst
                While Not UserRS.EOF
                    UserRS.Edit
                    For A = 1 To 20
                        UserRS.Fields("InvObject" + CStr(A)) = 0
                        UserRS.Fields("InvValue" + CStr(A)) = 0
                    Next A
                    For A = 1 To 6
                        UserRS.Fields("EquippedObject" + CStr(A)) = 0
                        UserRS.Fields("EquippedVal" + CStr(A)) = 0
                    Next A
                    For A = 0 To 29
                        UserRS.Fields("BankObject" + CStr(A)) = 0
                        UserRS.Fields("BankValue" + CStr(A)) = 0
                        UserRS.Fields("BankPrefix" + CStr(A)) = 0
                        UserRS.Fields("BankSuffix" + CStr(A)) = 0
                    Next A
                    UserRS!Bank = 0
                    UserRS.Update
                    UserRS.MoveNext
                Wend
            End If
            PrintLog "Inventories Reset"
        End If
    End If
End Sub

Private Sub mnuDatabaseResetItem_Click()
    Dim TheObj As Long, A As Long
    TheObj = Val(InputBox("What object?"))
    If TheObj = 0 Then Exit Sub

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            UserRS.Edit
            'Inventory Data
            For A = 1 To 20
                If TheObj = UserRS.Fields("InvObject" + CStr(A)) Then
                    UserRS.Fields("InvObject" + CStr(A)) = 0
                    UserRS.Fields("InvValue" + CStr(A)) = 0
                    UserRS.Fields("InvPrefix" + CStr(A)) = 0
                    UserRS.Fields("InvSuffix" + CStr(A)) = 0
                End If
            Next A
            For A = 1 To 6
                If TheObj = UserRS.Fields("EquippedObject" + CStr(A)) Then
                    UserRS.Fields("EquippedObject" + CStr(A)) = 0
                    UserRS.Fields("EquippedValue" + CStr(A)) = 0
                    UserRS.Fields("EquippedPrefix" + CStr(A)) = 0
                    UserRS.Fields("EquippedSuffix" + CStr(A)) = 0
                End If
            Next A
            For A = 0 To 29
                If TheObj = UserRS.Fields("BankObject" + CStr(A)) Then
                    UserRS.Fields("BankObject" + CStr(A)) = 0
                    UserRS.Fields("BankValue" + CStr(A)) = 0
                    UserRS.Fields("BankPrefix" + CStr(A)) = 0
                    UserRS.Fields("BankSuffix" + CStr(A)) = 0
                End If
            Next A
            UserRS.Update
            UserRS.MoveNext
        Wend
    End If

    Dim B As Long
    For A = 1 To MaxMaps
        With Map(A)
            If .Keep = True Then
                For B = 0 To MaxMapObjects
                    With .Object(B)
                        If Map(A).Object(B).Object = TheObj Then
                            Map(A).Object(B).Object = 0
                            Map(A).Object(B).Value = 0
                            Map(A).Object(B).ItemPrefix = 0
                            Map(A).Object(B).ItemSuffix = 0
                        End If
                    End With
                Next B
            End If
        End With
    Next A

    If MonsterRS.BOF = False Then
        MonsterRS.MoveFirst
        While MonsterRS.EOF = False
            MonsterRS.Edit
            If MonsterRS!Object0 = TheObj Then
                MonsterRS!Object0 = 0
                MonsterRS!Value0 = 0
            End If
            If MonsterRS!Object1 = TheObj Then
                MonsterRS!Object1 = 0
                MonsterRS!Value1 = 0
            End If
            If MonsterRS!Object2 = TheObj Then
                MonsterRS!Object2 = 0
                MonsterRS!Value2 = 0
            End If
            MonsterRS.Update
            MonsterRS.MoveNext
        Wend
    End If

    If NPCRS.BOF = False Then
        NPCRS.MoveFirst
        While NPCRS.EOF = False
            NPCRS.Edit
            For B = 0 To 9
                If NPCRS("GiveObject" + CStr(B)) = TheObj Then
                    NPCRS("GiveObject" + CStr(B)) = 0
                    NPCRS("GiveValue" + CStr(B)) = 0
                    NPCRS("TakeObject" + CStr(B)) = 0
                    NPCRS("TakeValue" + CStr(B)) = 0
                End If
                If NPCRS("TakeObject" + CStr(B)) = TheObj Then
                    NPCRS("GiveObject" + CStr(B)) = 0
                    NPCRS("GiveValue" + CStr(B)) = 0
                    NPCRS("TakeObject" + CStr(B)) = 0
                    NPCRS("TakeValue" + CStr(B)) = 0
                End If
            Next B
            NPCRS.Update
            NPCRS.MoveNext
        Wend
    End If

    ObjectRS.Seek "=", TheObj
    If ObjectRS.NoMatch Then
        ObjectRS.AddNew
        ObjectRS!number = A
    Else
        ObjectRS.Edit
    End If
    ObjectRS!ObjName = ""
    ObjectRS!Version = 0
    ObjectRS.Update
End Sub

Private Sub mnuDatabaseResetKeep_Click()
    If MsgBox("Reset Keep Tiles?", vbYesNo) = vbYes Then
        Dim A As Long, B As Long
        For A = 1 To MaxMaps
            With Map(A)
                If .Keep = True Then
                    For B = 0 To MaxMapObjects
                        With .Object(B)
                            If .Object > 0 Then
                                If Map(A).Tile(.X, .Y).Att = 5 Or Map(A).Tile(.X, .Y).Att2 = 5 Then
                                    Map(A).Object(B).Object = 0
                                    Map(A).Object(B).Value = 0
                                    Map(A).Object(B).ItemPrefix = 0
                                    Map(A).Object(B).ItemSuffix = 0
                                End If
                            End If
                        End With
                    Next B
                End If
            End With
        Next A
        DataRS.Edit
        DataRS!ObjectData = ""
        DataRS.Update
        PrintLog "Keep Tiles Reset"
    End If
End Sub

Private Sub mnuDatabaseResetMagic_Click()
    Dim A As Long

    If MsgBox("Are you *sure* you wish to delete all magic?", vbYesNo) = vbYes Then
        If MsgBox("About to reset all magic -- continue?", vbYesNo) = vbYes Then
            MagicRS.Close
            Set MagicRS = Nothing
            DB.TableDefs.Delete "Magic"
            CreateMagicTable
            Set MagicRS = DB.TableDefs("Magic").OpenRecordset(dbOpenTable)
            MagicRS.Index = "Number"
            For A = 1 To MaxMagic
                With Magic(A)
                    .Class = 0
                    .Description = vbNullString
                    .Level = 0
                    .Name = vbNullString
                    .Icon = 0
                    .IconType = 0
                    .CastTimer = 0
                    .Version = 0
                End With
            Next A
        End If
    End If
    PrintLog "Magic Reset"
End Sub

Private Sub mnuDatabaseResetMoney_Click()
    Dim TheObj As Long, A As Long, B As Long

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            UserRS.Edit
                'Inventory Data
                For A = 1 To 20
                    If Not IsNull(UserRS.Fields("InvObject" + CStr(A))) Then
                        TheObj = UserRS.Fields("InvObject" + CStr(A))
                        If TheObj = 6 Then
                            UserRS.Fields("InvObject" + CStr(A)) = 0
                        End If
                    End If
                Next A
    
                UserRS!Bank = 0
            UserRS.Update
            UserRS.MoveNext
        Wend
    End If
    
    For A = 1 To MaxMaps
        With Map(A)
            If .Keep = True Then
                For B = 0 To MaxMapObjects
                    With .Object(B)
                        If .Object = 6 Then .Object = 0
                    End With
                Next B
            End If
        End With
    Next A
End Sub

Private Sub mnuDatabaseResetMonsters_Click()
    Dim A As Long

    If MsgBox("Are you *sure* you wish to delete every monster?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every monster -- continue?", vbYesNo) = vbYes Then
            MonsterRS.Close
            Set MonsterRS = Nothing
            DB.TableDefs.Delete "Monsters"
            CreateMonstersTable
            Set MonsterRS = DB.TableDefs("Monsters").OpenRecordset(dbOpenTable)
            MonsterRS.Index = "Number"
            For A = 1 To MaxTotalMonsters
                With Monster(A)
                    .Armor = 0
                    .Agility = 0
                    .Description = ""
                    .flags = 0
                    .HP = 0
                    .Name = ""
                    .Sight = 0
                    .Speed = 0
                    .Sprite = 0
                    .Strength = 0
                    .Object(0) = 0
                    .Object(1) = 0
                    .Object(2) = 0
                    .Value(0) = 0
                    .Value(1) = 0
                    .Value(2) = 0
                    .Experience = 0
                End With
            Next A
        End If
    End If
    PrintLog "Monsters Reset"
End Sub

Private Sub mnuDatabaseResetNPCs_Click()
    Dim A As Long, B As Long

    If MsgBox("Are you *sure* you wish to delete every NPC?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every NPC -- continue?", vbYesNo) = vbYes Then
            NPCRS.Close
            Set NPCRS = Nothing
            DB.TableDefs.Delete "NPCS"
            CreateNPCsTable
            Set NPCRS = DB.TableDefs("NPCS").OpenRecordset(dbOpenTable)
            NPCRS.Index = "Number"
            For A = 1 To MaxNPCs
                With NPC(A)
                    .Name = ""
                    .JoinText = ""
                    .LeaveText = ""
                    .flags = 0
                    For B = 0 To 9
                        With .SaleItem(B)
                            .GiveObject = 0
                            .GiveValue = 0
                            .TakeObject = 0
                            .TakeValue = 0
                        End With
                    Next B
                End With
            Next A
        End If
    End If
End Sub
Private Sub mnuDatabaseResetObjects_Click()
    Dim A As Long

    If MsgBox("Are you *sure* you wish to delete every object?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every object -- continue?", vbYesNo) = vbYes Then
            ObjectRS.Close
            Set ObjectRS = Nothing
            DB.TableDefs.Delete "Objects"
            CreateObjectsTable
            Set ObjectRS = DB.TableDefs("Objects").OpenRecordset(dbOpenTable)
            ObjectRS.Index = "Number"

            For A = 1 To MaxObjects
                With Object(A)
                    .Data(0) = 0
                    .Data(1) = 0
                    .Data(2) = 0
                    .Data(3) = 0
                    .flags = 0
                    .Name = ""
                    .Picture = 0
                    .Type = 0
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuGold_Click()
    Open "report.txt" For Output As #1

    Print #1, "***Odyssey God Report ***"
    Print #1, ""

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If UserRS!Bank > 100000000 Then
                Print #1, "User=" + UserRS!User + " Name=" + UserRS!Name + " Bank=" + CStr(UserRS!Bank)
            End If
            UserRS.MoveNext
        Wend
    End If
    
    Dim A As Long, B As Long
    For A = 1 To MaxMaps
        With Map(A)
            If .Keep = True Then
                For B = 0 To MaxMapObjects
                    With .Object(B)
                        If .Object = 6 Then
                            If .Value > 10000000 Then
                                Print #1, "Map " & A & " (" & Map(A).Name & ") " & Object(Map(A).Object(B).Object).Name + "  -  " + CStr(Map(A).Object(B).Value)
                            End If
                        End If
                    End With
                Next B
            End If
        End With
    Next A

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub mnuLastScript_Click()
    Open "report.txt" For Output As #1

    Print #1, ""
    Print #1, "Last Script Report"
    Print #1, ""

    Print #1, "LastScript:  " + LastScript
    Print #1, "LastScript2:  " + LastScript2
    Print #1, "LastScript3:  " + LastScript3

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub mnuLogScripts_Click()
    If mnuLogScripts.Checked = False Then
        mnuLogScripts.Checked = True
    Else
        mnuLogScripts.Checked = False
    End If
End Sub

Private Sub mnuRelayChat_Click()
    If mnuRelayChat.Checked = False Then
        mnuRelayChat.Checked = True
    Else
        mnuRelayChat.Checked = False
    End If
End Sub

Private Sub mnuReportsGods_Click()
    
    Open "report.txt" For Output As #1

    Print #1, "***Odyssey God Report ***"
    Print #1, ""

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If UserRS!Access > 0 Then
                Print #1, "User=" + UserRS!User + " Name=" + UserRS!Name + " Access=" + CStr(UserRS!Access)
            End If
            UserRS.MoveNext
        Wend
    End If

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub
Private Sub mnuReportsMaps_Click()
    Dim A As Long, StartFree As Long, IsFree As Boolean
    Open "report.txt" For Output As #1

    Print #1, "***Odyssey Free Map Report ***"
    Print #1, ""

    IsFree = False

    For A = 1 To MaxMaps
        MapRS.Seek "=", A
        If MapRS.NoMatch = False And Not Map(A).Name = "" Then
            If IsFree = True Then
                If StartFree < A - 1 Then
                    Print #1, CStr(StartFree) + "-" + CStr(A - 1)
                Else
                    Print #1, CStr(A - 1)
                End If
                IsFree = False
            End If
        Else
            If IsFree = False Then
                StartFree = A
                IsFree = True
            End If
        End If
    Next A

    If IsFree = True Then
        If StartFree < MaxMaps Then
            Print #1, CStr(StartFree) + "-" + CStr(MaxMaps)
        Else
            Print #1, CStr(MaxMaps)
        End If
    End If
    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub mnuReportsObjects_Click()
    Open "report.txt" For Output As #1

    Print #1, ""
    Print #1, "Object Report"
    Print #1, ""
    Dim A As Long
    For A = 1 To MaxObjects
        If Object(A).Name <> "" Then
            Print #1, Object(A).Name & "  -  Object #" & A
        End If
    Next A
    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub mnuReportsUnlinked_Click()
    Dim TheMap(1 To MaxMaps) As Boolean, A As Long, X As Long, Y As Long, D As Long

    For A = 1 To MaxMaps
        If Not Map(A).Name = "" Then
            If Map(A).ExitLeft > 0 Then TheMap(Map(A).ExitLeft) = True
            If Map(A).ExitUp > 0 Then TheMap(Map(A).ExitUp) = True
            If Map(A).ExitRight > 0 Then TheMap(Map(A).ExitRight) = True
            If Map(A).ExitDown > 0 Then TheMap(Map(A).ExitDown) = True
            For Y = 0 To 11
                For X = 0 To 11
                    With Map(A).Tile(X, Y)
                        If .Att = 2 Then
                            D = CLng(Map(A).Tile(X, Y).AttData(0)) * 256 + CLng(Map(A).Tile(X, Y).AttData(1))
                            TheMap(D) = True
                        End If
                    End With
                Next X
            Next Y
        Else
            TheMap(A) = True
        End If
    Next A

    Open "report.txt" For Output As #1

    Print #1, "*** Unlinked Map Report ***"
    Print #1, ""

    For A = 1 To MaxMaps
        If TheMap(A) = False Then Print #1, "Map " & A & " (" & Map(A).Name & ")"
    Next A

    Close #1

    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub mnuServerOptions_Click()
    frmOptions.Show 1
End Sub

Private Sub PlayerTimer_Timer()
    On Error GoTo LogDatShit
    Dim St1 As String, A As Long, B As Long, C As Long
    Dim PlayerNum As Long
    
    Dim Tick As Currency
    Tick = getTime

    For PlayerNum = 1 To MaxUsers
        With Player(PlayerNum)
            If .InUse = True Then
                St1 = ""
                If .IsDead = True Then
                    If Tick > .DeadTick Then
                        Partmap PlayerNum

                        .IsDead = False
                        SendAllBut PlayerNum, Chr$(120) + Chr$(PlayerNum)
                        St1 = St1 + DoubleChar$(2) + Chr$(120) + Chr$(PlayerNum)

                        If .Guild > 0 Then
                            If Guild(.Guild).Hall >= 1 Then
                                A = 1
                            Else
                                A = 0
                            End If
                        Else
                            A = 0
                        End If

                        If A = 0 Then
                            'Random Start Location
                            Randomize
                            A = Int(Rnd * 4)
                            If .Level <= 15 Then A = 2

                            .Map = World.StartLocation(A).Map
                            .X = World.StartLocation(A).X
                            .Y = World.StartLocation(A).Y

                            If World.StartLocation(A).Message <> "" Then
                                St1 = St1 + DoubleChar(2 + Len(World.StartLocation(A).Message)) + Chr$(56) + Chr$(15) + World.StartLocation(A).Message
                            End If
                        Else
                            A = Guild(.Guild).Hall

                            .Map = Hall(A).StartLocation.Map
                            .X = Hall(A).StartLocation.X
                            .Y = Hall(A).StartLocation.Y
                        End If

                        If .Map < 1 Then .Map = 1
                        If .Map > MaxMaps Then .Map = MaxMaps
                        If .Y > 11 Then .Y = 11
                        If .X > 11 Then .X = 11

                        If St1 <> "" Then
                            SendRaw PlayerNum, St1
                            St1 = ""
                        End If

                        JoinMap PlayerNum

                        .HP = .MaxHP
                        .Mana = .MaxMana
                        .Energy = .MaxEnergy

                        Parameter(0) = PlayerNum
                        RunScript ("PLAYERRESURRECT")
                    End If
                End If

                If Tick - .LastMsg >= 30000 And Tick - .LastMsg <= 35000 Then
                    If .Mode <> modeNotConnected Then
                        'Send ping
                        St1 = St1 + DoubleChar(1) + Chr$(58)
                        .LastMsg = Tick - 35000
                    Else
                        AddSocketQue PlayerNum
                    End If
                End If
                If Tick - .LastMsg >= 90000 Then
                    'Lag time out
                    If .LastSkillUse = 69 Then

                    Else
                        AddSocketQue PlayerNum
                    End If
                End If
                If .Mode = modePlaying Then
                    If .IsDead = False Then
                        C = 0
                        If .HP < .MaxHP Then
                            C = 1
                            A = .HP
                            B = .HP

                            A = A + Int(World.StatEndurance)
                            If A > .MaxHP Then A = .MaxHP

                            .HP = A
                            If Not A = B Then
                                St1 = St1 + DoubleChar(2) + Chr$(46) + Chr$(.HP)
                                SendHPUpdate PlayerNum
                            End If
                        End If
                        If .Energy < .MaxEnergy Then
                            C = 1
                            A = .Energy
                            B = .Energy

                            If .HP > 0 Then
                                A = A + Int((CSng(.HP) / CSng(.MaxHP)) * 4)
                                If A > .MaxEnergy Then A = .MaxEnergy
                            End If

                            .Energy = A
                        End If
                        If .Mana < .MaxMana Then
                            C = 1
                            A = .Mana
                            B = .Mana

                            A = A + Int(World.StatIntelligence)
                            If A > .MaxMana Then A = .MaxMana

                            .Mana = A
                            If Not A = B Then
                                St1 = St1 + DoubleChar(2) + Chr$(48) + Chr$(.Mana)
                            End If
                        End If
                        
                        If C = 1 Then
                            Parameter(0) = PlayerNum
                            RunScript "PLAYERREGEN"
                        End If

                        If St1 <> "" Then
                            SendRaw PlayerNum, St1
                        End If
                    End If
                End If
                For A = 1 To MaxPlayerTimers
                    If .ScriptTimer(A) > 0 Then
                        If Tick >= .ScriptTimer(A) Then
                            Parameter(0) = PlayerNum
                            .ScriptTimer(A) = 0
                            RunScript .Script(A)
                        End If
                    End If
                Next A
            End If
        End With
    Next PlayerNum

    Exit Sub

LogDatShit:
    If PlayerNum > 0 Then
        SendToGods Chr$(56) + Chr$(7) + "WARNING:  Server Crashed on Player Timer from " & Player(PlayerNum).Name & "  DATA:  " & Err.Description
        PrintLog "WARNING:  Server Crashed on Player Timer from " & Player(PlayerNum).Name & "  DATA:  " & Err.Description
        PrintDebug "WARNING:  Server Crashed on Player Timer from " & Player(PlayerNum).Name & "  DATA:  " & Err.Description
        BanPlayer PlayerNum, 0, 1, "Crashed Server", "Server"
    Else
        SendToGods Chr$(56) + Chr$(7) + "WARNING:  Server Crashed on Player Timer   DATA:  " & Err.Description
        PrintLog "WARNING:  Server Crashed on Player Timer    DATA:  " & Err.Description
        PrintDebug "WARNING:  Server Crashed on Player Timer    DATA:  " & Err.Description
    End If
End Sub

Private Sub tmrCloseScks_Timer()
    On Error GoTo LogDatShit

    Dim A As Long
    'Wait Procedure for Sockets
    For A = 1 To MaxUsers
        If CloseSocketQue(A) > 0 Then
            If Player(CloseSocketQue(A)).Mode = modePlaying Then
                If Player(CloseSocketQue(A)).HP = Player(CloseSocketQue(A)).MaxHP And Player(CloseSocketQue(A)).Mana = Player(CloseSocketQue(A)).MaxMana And Player(CloseSocketQue(A)).Energy = Player(CloseSocketQue(A)).MaxEnergy Then
                    CloseClientSocket CloseSocketQue(A)
                    CloseSocketQue(A) = 0
                End If
            Else
                CloseClientSocket CloseSocketQue(A)
                CloseSocketQue(A) = 0
            End If
        End If
    Next A

    Exit Sub

LogDatShit:
    SendToGods Chr$(56) + Chr$(7) + "WARNING:  Server Crashed on Close Socket Timer: " + Err.Description
    PrintLog "WARNING:  Server Crashed on Close Socket Timer: " + Err.Description
    PrintDebug "WARNING:  Server Crashed on Close Socket Timer: " + Err.Description
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        SendAll Chr$(30) + txtMessage
        PrintLog "Server Message: " + txtMessage
        txtMessage = ""
    End If
End Sub
