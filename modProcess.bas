Attribute VB_Name = "modProcess"
Option Explicit

Sub ReadClientData(Index As Long)
    On Error GoTo LogDatShit
    Dim St As String, SocketData As String, PacketLength As Long, PacketOrder As Integer, PacketID As Long, PacketCheckSum As Long, CurrentCheckSum As Long
    Dim MapNum As Long
    Dim Tick As Currency
    Tick = getTime

    With Player(Index)
        MapNum = .Map
        SocketData = .SocketData + Receive(.Socket)
        .LastMsg = Tick
        If .Access = 0 Then
            If .LastMsg > .FloodTimer Then
                .FloodTimer = .LastMsg
            End If
            If .FloodTimer - .LastMsg > 5000 And .Flag(41) = 0 And .Mode = modePlaying Then
                .FloodTimer = 0
                SendAll Chr$(56) + Chr$(15) + .Name + " has been autosquelched!"
                .Flag(41) = 5
                Exit Sub
            End If
        End If
LoopRead:
        If Len(SocketData) >= 5 Then
            PacketLength = GetInt(Mid$(SocketData, 1, 2))
            PacketCheckSum = Asc(Mid$(SocketData, 3, 1))
            PacketOrder = Asc(Mid$(SocketData, 4, 1))
            If PacketLength >= 5072 And .Access = 0 Then
                Hacker Index, "C.1.2"
                Exit Sub
            End If
            If Len(SocketData) - 4 >= PacketLength Then
                St = Mid$(SocketData, 5, PacketLength)
                SocketData = Mid$(SocketData, PacketLength + 5)
                If PacketLength > 0 Then
                    PacketID = Asc(Mid$(St, 1, 1))
                    CurrentCheckSum = CheckSum(St) * 20 Mod 194
                    If Not CurrentCheckSum = PacketCheckSum Then
                        PrintDebug "Bad Checksum from " & .Name + " " + .IP
                        AddSocketQue Index
                        Exit Sub
                    End If
                    If Not PacketOrder = .PacketOrder Then
                        AddSocketQue Index
                        Exit Sub
                    Else
                        .PacketOrder = .PacketOrder + 1
                        If .PacketOrder > 250 Then .PacketOrder = 0
                    End If
                    If Len(St) > 1 Then
                        St = Mid$(St, 2)
                    Else
                        St = ""
                    End If
                    Select Case PacketID
                    Case 170:    'Raw Data
                        ProcessRawData Index, St
                    Case Else:
                        ProcessString Index, PacketID, St
                    End Select
                End If
                GoTo LoopRead
            End If
        End If
        .SocketData = SocketData
    End With

    Exit Sub

LogDatShit:
    SendToGods Chr$(16) & Chr$(0) & "WARNING:  Server Crashed on Packet # " + CStr(PacketID) & " from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    PrintLog "WARNING:  Server Crashed on Packet # " + CStr(PacketID) & " from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    PrintDebug "WARNING:  Server Crashed on Packet # " + CStr(PacketID) & " from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    BootPlayer Index, 0, "Crashed Server"
End Sub

Sub ProcessGodCommand(Index As Long, St As String)
    Dim A As Long, B As Long, C As Long, D As Long, St1 As String
    With Player(Index)
        Select Case Asc(Mid$(St, 1, 1))
        Case 0    'Server Message
            If Len(St) >= 2 Then
                SendAll Chr$(30) + "[" + .Name + "] " + Mid$(St, 2)
            Else
                Hacker Index, "A.28"
            End If

        Case 1    'Warp
            If Len(St) = 5 Then
                A = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                B = Asc(Mid$(St, 4, 1))
                C = Asc(Mid$(St, 5, 1))
                If A >= 1 And A <= MaxMaps And B <= 11 And C <= 11 Then
                    PrintGod Player(Index).User, " (Warp) Map: " + CStr(A) + " X: " + CStr(B) + " Y: " + CStr(C) + " Map Name: " + Map(A).Name
                    Partmap Index
                    .Map = A
                    .X = B
                    .Y = C
                    JoinMap Index
                End If
            Else
                Hacker Index, "A.29"
            End If

        Case 2    'WarpMe
            If Len(St) = 2 Then
                A = Asc(Mid$(St, 2, 1))
                If A >= 1 And A <= MaxUsers And A <> Index Then
                    If Player(A).Mode = modePlaying Then
                        PrintGod Player(Index).User, " (WarpMe) Warped to Player: " + Player(A).Name + " Map: " + CStr(Player(A).Map) + " Map Name: " + Map(Player(A).Map).Name
                        Partmap Index
                        .Map = Player(A).Map
                        .X = Player(A).X
                        .Y = Player(A).Y
                        JoinMap Index
                    End If
                End If
            Else
                Hacker Index, "A.30"
            End If

        Case 3    'WarpPlayer
            If Len(St) = 6 Then
                A = Asc(Mid$(St, 2, 1))
                B = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                C = Asc(Mid$(St, 5, 1))
                D = Asc(Mid$(St, 6, 1))
                If A >= 1 And A <= MaxUsers And B >= 1 And B <= MaxMaps And C <= 11 And D <= 11 Then
                    With Player(A)
                        If .Mode = modePlaying Then
                            PrintGod Player(Index).User, " (WarpToMe) Warped Player: " + Player(A).Name + " Map: " + CStr(Player(Index).Map) + " Map Name: " + Map(Player(Index).Map).Name
                            Partmap A
                            .Map = B
                            .X = C
                            .Y = D
                            JoinMap A
                        End If
                    End With
                End If
            Else
                Hacker Index, "A.31"
            End If

        Case 4    'Set MOTD
            If Len(St) > 1 And .Access >= 3 Then
                World.MOTD = Mid$(St, 2)
                PrintGod Player(Index).User, " (MOTD) " + World.MOTD
                DataRS.Edit
                DataRS!MOTD = World.MOTD
                DataRS.Update
            Else
                Hacker Index, "A.32"
            End If

        Case 5    'Disband Guild
            If Len(St) = 2 And .Access >= 2 Then
                A = Asc(Mid$(St, 2, 1))
                PrintGod Player(Index).User, " (Delete Guild) " + Guild(A).Name
                DeleteGuild A, 3
            Else
                Hacker Index, "A.33"
            End If

        Case 6    'Set sprite
            If Len(St) = 4 Then
                A = Asc(Mid$(St, 2, 1))
                B = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                If A >= 1 And A <= MaxUsers And B <= MaxSprite Then
                    With Player(A)
                        If .Mode = modePlaying Then
                            PrintGod Player(Index).User, " (SetSprite) " + Player(A).Name
                            If B = 0 Then
                                .Sprite = .Class * 2 + .Gender - 1
                            Else
                                .Sprite = B
                            End If
                            SendToMap .Map, Chr$(63) + Chr$(A) + DoubleChar$(CLng(.Sprite))
                        End If
                    End With
                End If
            Else
                Hacker Index, "A.34"
            End If

        Case 7    'Set name
            If Len(St) >= 3 And Len(St) <= 17 And .Access >= 2 Then
                A = Asc(Mid$(St, 2, 1))
                If A >= 1 And A <= MaxUsers Then
                    With Player(A)
                        If .Mode = modePlaying Then
                            UserRS.Index = "Name"
                            UserRS.Seek "=", Mid$(St, 3)
                            If UserRS.NoMatch = True Then
                                PrintGod Player(Index).User, " (SetName) " + Player(A).Name + " (to) " + Mid$(St, 3)
                                .Name = Mid$(St, 3)
                                SendAll Chr$(64) + Chr$(A) + .Name
                            Else
                                SendSocket Index, Chr$(16) + Chr$(16)
                            End If
                        End If
                    End With
                End If
            Else
                Hacker Index, "A.35"
            End If

        Case 8    'Resetmap
            If Len(St) = 1 And .Access >= 2 Then
                Dim PlayerString As String
                For A = 1 To MaxUsers
                    If Player(A).Map = .Map Then PlayerString = PlayerString + Player(A).Name + ", "
                Next A
                PrintGod Player(Index).User, " (ResetMap) Map: " + CStr(Player(Index).Map) + " - " + PlayerString
                ResetMap CLng(.Map)
            Else
                Hacker Index, "A.84"
            End If

        Case 9    'Boot
            If Len(St) >= 2 Then
                A = Asc(Mid$(St, 2, 1))
                If A >= 1 And A <= MaxUsers Then
                    If Not Player(A).Access = 4 Then
                        PrintGod Player(Index).User, " (Boot) Booted " + Player(A).Name + " for reason " + Mid$(St, 3)
                        BootPlayer A, Index, Mid$(St, 3)
                    End If
                End If
            Else
                Hacker Index, "A.36"
            End If

        Case 10    'Ban
            If Len(St) >= 3 Then
                A = Asc(Mid$(St, 2, 1))
                If A >= 1 And A <= MaxUsers Then
                    If Not Player(A).Access = 4 Then
                        If BanPlayer(A, Index, Asc(Mid$(St, 3, 1)), Mid$(St, 4), Player(Index).Name) = False Then
                            PrintGod Player(Index).User, " (Ban) Banned " + Player(A).Name + " for reason " + Mid$(St, 4)
                            SendSocket Index, Chr$(16) + Chr$(13)    'Ban list full
                        End If
                    End If
                End If
            Else
                Hacker Index, "A.37"
            End If

        Case 11    'Remove Ban
            If Len(St) = 2 Then
                A = Asc(Mid$(St, 2, 1))
                If A >= 1 And A <= 50 Then
                    If Not Ban(A).Banner = "SuperAdmin" Or .Access = 4 Then
                        PrintGod Player(Index).User, " (Remove Ban) Player: " + Ban(A).Name
                        SendAll Chr$(56) + Chr$(15) + Ban(A).Name + " has been unbanned by " + .Name + "."
                        With Ban(A)
                            .ComputerID = ""
                            .IPAddress = ""
                            .Name = ""
                            .Reason = ""
                            .UnbanDate = 0
                            .Banner = ""
                            .InUse = False
                        End With
                        BanRS.Seek "=", A
                        If BanRS.NoMatch = False Then
                            BanRS.Delete
                        End If
                    End If
                End If
            Else
                Hacker Index, "A.38"
            End If

        Case 12    'List Bans
            If Len(St) = 1 And .Access >= 1 Then
                St1 = ""
                For A = 1 To 50
                    With Ban(A)
                        If .InUse = True Then
                            St1 = St1 + DoubleChar(2 + Len(.Name)) + Chr$(69) + Chr$(A) + .Name
                        End If
                    End With
                Next A
                St1 = St1 + DoubleChar(1) + Chr$(69)
                SendRaw Index, St1
            Else
                Hacker Index, "A.39"
            End If

        Case 14    'Chat
            If Len(St) >= 2 Then
                SendToGodsAllBut Index, Chr$(90) + Chr$(Index) + Mid$(St, 2)
                PrintChat "God", Player(Index).User + ": " + Mid$(St, 2)
            Else
                Hacker Index, "Attempted God Chat Attempt"
            End If

        Case 15    'Set Guild Sprite
            If Len(St) = 3 And .Access >= 2 Then
                A = Asc(Mid$(St, 2, 1))
                B = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                If A >= 1 Then
                    UserRS.Index = "Name"
                    With Guild(A)
                        If .Name <> "" Then
                            .Sprite = B
                            GuildRS.Bookmark = .Bookmark
                            GuildRS.Edit
                            GuildRS!Sprite = B
                            GuildRS.Update

                            For C = 0 To 19
                                With .Member(C)
                                    If .Name <> "" Then
                                        D = FindPlayer(.Name)
                                        If D > 0 Then
                                            With Player(D)
                                                If B > 0 Then
                                                    .Sprite = B
                                                Else
                                                    .Sprite = .Class * 2 + .Gender - 1
                                                End If
                                                SendToMap .Map, Chr$(63) + Chr$(D) + DoubleChar$(CLng(.Sprite))
                                            End With
                                        Else
                                            UserRS.Seek "=", .Name
                                            If UserRS.NoMatch = False Then
                                                If B > 0 Then
                                                    D = B
                                                Else
                                                    D = UserRS!Class * 2 + UserRS!Gender - 1
                                                End If
                                                If D >= 1 And D <= MaxSprite Then
                                                    UserRS.Edit
                                                    UserRS!Sprite = D
                                                    UserRS.Update
                                                End If
                                            End If
                                        End If
                                    End If
                                End With
                            Next C
                        End If
                    End With
                End If
            Else
                Hacker Index, "A.41"
            End If
        Case 16    'Float
            If Len(St) > 4 Then
                A = Asc(Mid$(St, 2, 1))
                B = Asc(Mid$(St, 3, 1))
                C = Asc(Mid$(St, 4, 1))
                SendToMap .Map, Chr$(112) + Chr$(A) + Chr$(B) + Chr$(C) + Mid$(St, 5)
            End If
        Case 17    'Set Status
            If Len(St) = 3 Then
                A = Asc(Mid$(St, 2, 1))
                B = Asc(Mid$(St, 3, 1))
                If A >= 1 And A <= MaxUsers And B <= 255 Then
                    With Player(A)
                        If .Mode = modePlaying Then
                            PrintGod Player(Index).User, " (Set Status) Player: " + .Name + " - Status: " + CStr(B)
                            If B = 0 Then
                                .Status = 0
                            Else
                                .Status = B
                            End If
                            SendToMap .Map, Chr$(91) + Chr$(A) + Chr$(.Status)
                        End If
                    End With
                End If
            End If
        Case 18    'Scan
            If Len(St) >= 1 And .Access >= 2 Then
                A = Asc(Mid$(St, 2, 1))
                If A >= 1 And A <= MaxUsers Then
                    PrintGod Player(Index).User, " (Scan) Player: " + Player(A).Name
                    St1 = ""
                    With Player(A)
                        St1 = Chr$(104) & Chr$(A) & Chr$(.Class) & Chr$(.Level) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(.MaxHP) & Chr$(.MaxMana) & Chr$(.MaxEnergy)
                        For B = 1 To 20
                            St1 = St1 & DoubleChar$(CLng(.Inv(B).Object)) & QuadChar(.Inv(B).Value)
                        Next B
                        For B = 1 To 5
                            St1 = St1 & DoubleChar$(CLng(.EquippedObject(B).Object)) & QuadChar(.EquippedObject(B).Value)
                        Next B
                        For B = 0 To 29
                            St1 = St1 & DoubleChar$(CLng(.ItemBank(B).Object)) & QuadChar(.ItemBank(B).Value)
                        Next B
                        SendSocket Index, St1
                        SendSocket Index, Chr$(16) & Chr$(0) & .User & " - " & .IP & " - Bank: " & CStr(.Bank)
                        SendSocket A, Chr$(154) + Chr$(Index)
                    End With
                End If
            End If
        Case 25    'Set Class
            If Len(St) = 3 Then
                A = Asc(Mid$(St, 2, 1))
                B = Asc(Mid$(St, 3, 1))
                If A >= 1 And A <= MaxUsers And B <= 255 Then
                    With Player(A)
                        If .Mode = modePlaying Then
                            If B > 0 & B <= NumClasses Then
                                .Class = B
                                SetPlayerClass A, B
                                SetPlayerHP A, .MaxHP
                                SetPlayerEnergy A, .MaxEnergy
                                SetPlayerMana A, .MaxMana
                                SetPlayerSprite A, 0
                                CalculateStats A
                            End If
                        End If
                    End With
                End If
            Else
                Hacker Index, "A.34"
            End If
        End Select
    End With
End Sub

Sub ProcessScriptProjectile(Index As Long, St As String)
    Dim Damage As Long
    Dim Target As Long
    Target = Asc(Mid$(St, 2, 1))
    Damage = Asc(Mid$(St, 3, 1))
    
    If Damage > 20 Then
        Hacker Index, "P.1"
    End If
    
    If Player(Index).ProjectileDamage(Damage).Live = True Then
        Player(Index).ProjectileDamage(Damage).Live = False
        Damage = Player(Index).ProjectileDamage(Damage).Damage
    Else
        Damage = 0
    End If
    
    Select Case Asc(Mid$(St, 1, 1))
        Case 1    'Magic Monster Attack
            MagicAttackMonster Index, Target, Damage
        Case 2    'Non Magic Monster Attack
            AttackMonster Index, Target, Damage
        Case 3    'Magic Player Attack
            MagicAttackPlayer Index, Target, Damage
        Case 4    'Non Magic Player Attack
            AttackPlayer Index, Target, Damage
    End Select
End Sub

Sub ProcessString(Index As Long, PacketID As Long, St As String)
    On Error GoTo LogDatShit
    Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long, G As Long, H As Long, I As Long
    Dim St1 As String, St2 As String
    Dim MapNum As Long
    Dim Tick As Currency
    
    Dim GetObjResult As Long
    Dim DropObjResult As Long
    
    Tick = getTime()

    With Player(Index)
        MapNum = .Map
        Select Case .Mode
        Case modeNotConnected
            Select Case PacketID
            Case 0    'New Account
                If .ClientVer = CurrentClientVer Then
                    A = InStr(1, St, Chr$(0))
                    If A > 1 And A < Len(St) Then
                        St1 = Trim$(Mid$(St, 1, A - 1))
                        B = Len(St1)
                        If B >= 3 And B <= 15 And ValidName(St1) Then
                            UserRS.Index = "User"
                            UserRS.Seek "=", St1

                            E = 0
                            St1 = UCase$(St1)
                            For F = 1 To MaxUsers
                                If F <> Index Then
                                    If St1 = UCase$(Player(F).User) Then
                                        E = 1
                                        Exit For
                                    End If
                                End If
                            Next F

                            If UserRS.NoMatch = True And GuildNum(St1) = 0 And E = 0 Then
                                UserRS.AddNew
                                UserRS!User = St1
                                .User = St1
                                St1 = Trim$(UCase$(Mid$(St, A + 1)))
                                If Len(St1) > 15 Then
                                    UserRS!Password = Left$(St1, 15)
                                Else
                                    UserRS!Password = St1
                                End If
                                UserRS.Update
                                UserRS.Seek "=", .User
                                .Bookmark = UserRS.Bookmark
                                .Access = 0
                                .Class = 0
                                SavePlayerData Index
                                SendSocket Index, Chr$(2)    'New account created!
                                AddSocketQue Index
                            Else
                                SendSocket Index, Chr$(1) + Chr$(1)    'User Already Exists
                                AddSocketQue Index
                            End If
                        Else
                            Hacker Index, "A.79"
                        End If
                    Else
                        AddSocketQue Index
                    End If
                Else
                    SendSocket Index, Chr$(116)
                    AddSocketQue Index
                End If

            Case 1    'Log on
                If .ClientVer = CurrentClientVer Then
                    A = InStr(1, St, Chr$(0))
                    If A > 1 And A < Len(St) Then
                        .User = Mid$(St, 1, A - 1)
                        UserRS.Index = "User"
                        UserRS.Seek "=", .User
                        If UserRS.NoMatch = False Then
                            If UCase$(Mid$(St, A + 1)) = UCase$(UserRS!Password) Then
                                B = 0
                                St1 = UCase$(.User)
                                For A = 1 To MaxUsers
                                    If A <> Index Then
                                        If St1 = UCase$(Player(A).User) Then
                                            AddSocketQue A
                                            B = 1
                                            Exit For
                                        End If
                                        'If .ComputerID = Player(A).ComputerID Then
                                        '    AddSocketQue A
                                        '    B = 2
                                        '    Exit For
                                        'End If
                                    End If
                                Next A
                                If B = 0 Then
                                    'Account Data
                                    .Access = UserRS!Access
                                    .Bookmark = UserRS.Bookmark

                                    'Character Data
                                    .Name = UserRS!Name
                                    CheckBan Index, .Name, .ComputerID, .IP
                                    PrintLog "Login accepted from " + .Name
                                    .Class = UserRS!Class
                                    .Gender = UserRS!Gender
                                    .Sprite = UserRS!Sprite
                                    .Map = UserRS!Map
                                    If Not IsNull(UserRS!Email) Then
                                        .Email = UserRS!Email
                                    Else
                                        .Email = ""
                                    End If
                                    If .Map < 1 Then .Map = 1
                                    If .Map > MaxMaps Then .Map = MaxMaps
                                    .X = UserRS!X
                                    If .X > 11 Then .X = 11
                                    .Y = UserRS!Y
                                    If .Y > 11 Then .Y = 11
                                    .D = UserRS!D
                                    .desc = UserRS!desc
                                    .IsDead = False
                                    .CastTimer = 0
                                    .SpeedTick = 0
                                    .LastSkillUse = 0

                                    'Character Physical Stats
                                    .Level = UserRS!Level
                                    If .Level > World.MaxLevel Then .Level = World.MaxLevel
                                    .Experience = UserRS!Experience

                                    'Inventory Data
                                    For A = 1 To 20
                                        .Inv(A).Object = UserRS.Fields("InvObject" + CStr(A))
                                        .Inv(A).Value = UserRS.Fields("InvValue" + CStr(A))
                                        .Inv(A).ItemPrefix = UserRS.Fields("InvPrefix" + CStr(A))
                                        .Inv(A).ItemSuffix = UserRS.Fields("InvSuffix" + CStr(A))
                                        
                                        If .Inv(A).ItemPrefix > 0 Then
                                            If ItemPrefix(.Inv(A).ItemPrefix).Name = "" Then
                                                .Inv(A).ItemPrefix = 0
                                            End If
                                        End If
                                        If .Inv(A).ItemSuffix > 0 Then
                                            If ItemSuffix(.Inv(A).ItemSuffix).Name = "" Then
                                                .Inv(A).ItemSuffix = 0
                                            End If
                                        End If
                                    Next A
                                    For A = 1 To 6
                                        .EquippedObject(A).Object = UserRS.Fields("EquippedObject" + CStr(A))
                                        .EquippedObject(A).Value = UserRS.Fields("EquippedVal" + CStr(A))
                                        .EquippedObject(A).ItemPrefix = UserRS.Fields("EquippedPrefix" + CStr(A))
                                        .EquippedObject(A).ItemSuffix = UserRS.Fields("EquippedSuffix" + CStr(A))
                                    Next A
                                    
                                    If .Class > 0 Then CalculateStats Index

                                    'Item Bank
                                    For B = 0 To 29
                                        .ItemBank(B).Object = UserRS.Fields("BankObject" + CStr(B))
                                        .ItemBank(B).Value = UserRS.Fields("BankValue" + CStr(B))
                                        .ItemBank(B).ItemPrefix = UserRS.Fields("BankPrefix" + CStr(B))
                                        .ItemBank(B).ItemSuffix = UserRS.Fields("BankSuffix" + CStr(B))
                                    Next B

                                    .SpeedStrikes = 0

                                    'Character Vital Stats
                                    .HP = .MaxHP
                                    .Energy = .MaxEnergy
                                    .Mana = .MaxMana

                                    'Flags
                                    For B = 0 To MaxPlayerFlags
                                        .Flag(B) = 0
                                    Next B
                                    Dim Position As Long
                                    St1 = ""
                                    If Not IsNull(UserRS!flags) Then
                                        St1 = UserRS!flags

                                        For B = 0 To MaxPlayerFlags
                                            If Position < Len(St1) Then
                                                A = Asc(Mid$(St1, Position + 1, 1)) * 256& + Asc(Mid$(St1, Position + 2, 1))
                                                .Flag(A) = Asc(Mid$(St1, Position + 3, 1)) * 16777216 + Asc(Mid$(St1, Position + 4, 1)) * 65536 + Asc(Mid$(St1, Position + 5, 1)) * 256& + Asc(Mid$(St1, Position + 6, 1))
                                                Position = Position + 6
                                            End If
                                        Next B
                                    End If

                                    'Load Skills
                                    St1 = ""
                                    If Not IsNull(UserRS!Skills) Then
                                        St1 = UserRS!Skills
                                        For A = 0 To 9
                                            With .Skill(A + 1)
                                                If Len(St1) >= A * 5 + 5 Then
                                                    .Level = Asc(Mid$(St1, A * 5 + 1, 1))
                                                    .Experience = Asc(Mid$(St1, A * 5 + 2, 1)) * 16777216 + Asc(Mid$(St1, A * 5 + 3, 1)) * 65536 + Asc(Mid$(St1, A * 5 + 4, 1)) * 256& + Asc(Mid$(St1, A * 5 + 5, 1))
                                                Else
                                                    .Level = 0
                                                    .Experience = 0
                                                End If
                                            End With
                                        Next A
                                    End If
                                    
                                    'Load Magic
                                    St1 = ""
                                    If Not IsNull(UserRS!Magic) Then
                                        St1 = UserRS!Magic
                                        For A = 0 To 254
                                            With .MagicLevel(A + 1)
                                                If Len(St1) >= A * 5 + 6 Then
                                                    .Level = Asc(Mid$(St1, A * 5 + 1, 1))
                                                    .Experience = Asc(Mid$(St1, A * 5 + 2, 1)) * 16777216 + Asc(Mid$(St1, A * 5 + 3, 1)) * 65536 + Asc(Mid$(St1, A * 5 + 4, 1)) * 256& + Asc(Mid$(St1, A * 5 + 5, 1))
                                                Else
                                                    .Level = 0
                                                    .Experience = 0
                                                End If
                                            End With
                                        Next A
                                    End If

                                    'Misc Data
                                    .Bank = UserRS!Bank
                                    .Status = UserRS!Status

                                    .Guild = 0
                                    .GuildRank = 0
                                    .GuildSlot = 0

                                    'Find Guild
                                    St1 = .Name
                                    For A = 1 To MaxGuilds
                                        With Guild(A)
                                            If .Name <> "" Then
                                                For B = 0 To 19
                                                    If .Member(B).Name = St1 Then
                                                        Player(Index).Guild = A
                                                        Player(Index).GuildRank = .Member(B).Rank
                                                        Player(Index).GuildSlot = B
                                                        Exit For
                                                    End If
                                                Next B
                                            End If
                                        End With
                                    Next A

                                    For A = 1 To MaxPlayerTimers
                                        .JoinRequest = 0
                                        .ScriptTimer(A) = 0
                                    Next A

                                    If Not .Mode = modeBanned Then
                                        .Mode = modeConnected

                                        SendCharacterData Index
                                        If World.MOTD <> "" Then
                                            SendSocket Index, Chr$(4) + World.MOTD
                                        End If
                                    End If
                                Else
                                    If B = 1 Then
                                        SendSocket Index, Chr$(0) + Chr$(2)
                                    Else
                                        SendSocket Index, Chr$(0) + Chr$(5)
                                    End If
                                End If
                            Else
                                SendSocket Index, Chr$(0) + Chr$(1)    'Invalid User/Password
                            End If
                        Else
                            SendSocket Index, Chr$(0) + Chr$(1)    'Invalid User/Password
                        End If
                    Else
                        Hacker Index, "A.1"
                    End If
                End If

            Case 29    'Pong
                If Len(St) > 0 Then
                    Hacker Index, "A.2"
                End If
                
            Case 35     'Get player count
                SendRawReal Index, CStr(NumUsers - 1)

            Case 61    'Version & CompID
                If Len(St) > 1 Then
                    ProcessVersion Index, St
                Else
                    SendSocket Index, Chr$(116)
                End If
            End Select
        Case modeConnected
            Select Case PacketID
            Case 2    'Create New Character
                If Len(St) >= 4 Then
                    A = InStr(3, St, vbNullChar)
                    If A > 1 Then
                        St1 = Trim$(Mid$(St, 3, A - 3))
                        If Len(St1) >= 3 And Len(St1) <= 15 And ValidName(St1) Then
                            UserRS.Index = "Name"
                            UserRS.Seek "=", St1
                            If (UserRS.NoMatch Or UCase$(.Name) = UCase$(St1)) And GuildNum(St1) = 0 And NPCNum(St1) = 0 Then
                                If .Class > 0 Then
                                    UserRS.Bookmark = .Bookmark
                                    DeleteCharacter
                                End If
                                .Class = Asc(Mid$(St, 1, 1))
                                If .Class < 1 Then .Class = 1
                                If .Class > NumClasses Then .Class = NumClasses
                                .Gender = Asc(Mid$(St, 2, 1))
                                If .Gender > 1 Then .Gender = 1
                                .Sprite = .Class * 2 + .Gender - 1
                                .Name = St1
                                If A < Len(St) Then
                                    St1 = Mid$(St, A + 1)
                                    If Len(St1) > 255 Then
                                        .desc = Left$(St1, 255)
                                    Else
                                        .desc = St1
                                    End If
                                Else
                                    .desc = ""
                                End If
                                .Level = 1
                                .Bank = 0
                                .Status = 2
                                .MaxHP = Class(.Class).StartHP
                                .HP = .MaxHP
                                .MaxEnergy = Class(.Class).StartEnergy
                                .Energy = .MaxEnergy
                                .MaxMana = Class(.Class).StartMana
                                .Mana = .MaxMana
                                .Email = ""
                                .Experience = 0
                                .SpeedStrikes = 0
                                For A = 1 To 20
                                    .Inv(A).Object = 0
                                    .Inv(A).Value = 0
                                    .Inv(A).ItemPrefix = 0
                                    .Inv(A).ItemSuffix = 0
                                Next A
                                For A = 1 To 6
                                    .EquippedObject(A).Object = 0
                                    .EquippedObject(A).Value = 0
                                    .EquippedObject(A).ItemPrefix = 0
                                    .EquippedObject(A).ItemSuffix = 0
                                Next A
                                For A = 0 To 29
                                    .ItemBank(A).Object = 0
                                    .ItemBank(A).Value = 0
                                    .ItemBank(A).ItemPrefix = 0
                                    .ItemBank(A).ItemSuffix = 0
                                Next A
                                For A = 0 To 255
                                    .Flag(A) = 0
                                Next A
                                .Map = World.StartLocation(0).Map
                                If .Map < 1 Then .Map = 1
                                If .Map > MaxMaps Then .Map = MaxMaps
                                .X = World.StartLocation(0).X
                                If .X > 11 Then .X = 11
                                .Y = World.StartLocation(0).Y
                                If .Y > 11 Then .Y = 11
                                .Guild = 0
                                .GuildRank = 0
                                GiveStartingEQ Index
                                SavePlayerData Index
                                CalculateStats Index
                                SendCharacterData Index
                            Else
                                SendSocket Index, Chr$(13)    'Name already in use
                            End If
                        Else
                            AddSocketQue Index
                        End If
                    Else
                        Hacker Index, "A.4"
                    End If
                Else
                    Hacker Index, "A.5"
                End If

            Case 4    'Delete Account
                If Len(St) = 0 Then
                    .Class = 0
                    UserRS.Bookmark = .Bookmark
                    DeleteAccount
                    AddSocketQue Index
                Else
                    Hacker Index, "A.7"
                End If

            Case 5    'Play
                If .Class > 0 Then
                    If MapNum > 0 Then
                        Hacker Index, "Old"
                    Else
                        Hacker Index, "A.8"
                    End If
                Else
                    Hacker Index, "A.9"
                End If

            Case 6    'Done Processing New Data
                SendDataPacket Index, 1

            Case 7    'Request List
                If Len(St) = 1 Then
                    Select Case Asc(Mid$(St, 1, 1))
                    Case 1    'Objects
                        SendSocket Index, Chr$(122) + ObjectVersionList
                    Case 2    'NPCs
                        SendSocket Index, Chr$(123) + NPCVersionList
                    Case 3    'Halls
                        SendSocket Index, Chr$(124) + HallVersionList
                    Case 4    'Monsters
                        SendSocket Index, Chr$(125) + MonsterVersionList
                    Case 5    'Magic
                        SendSocket Index, Chr$(126) + MagicVersionList
                    Case 6    'Prefix
                        SendSocket Index, Chr$(131) + PrefixVersionList
                    Case 7    'Suffix
                        SendSocket Index, Chr$(132) + SuffixVersionList
                    Case 8    'Server Options
                        SendServerOptions Index
                        SendSocket Index, Chr$(140)
                    End Select
                Else
                    Hacker Index, "Request List"
                End If

            Case 23    'Done receiving Data
                If Len(St) = 0 Then
                    JoinGame Index
                Else
                    Hacker Index, "A.10"
                End If

            Case 24    'Send Next Packet
                If Len(St) = 1 Then
                    SendDataPacket Index, Asc(Mid$(St, 1, 1))
                Else
                    Hacker Index, "A.11"
                End If

            Case 29    'Pong
                If Len(St) > 0 Then
                    Hacker Index, "A.13"
                End If

            Case 79    'Request Object
                If Len(St) > 0 Then
                    A = GetInt(Mid$(St, 1, 2))
                    With Object(A)
                        SendSocket Index, Chr$(31) + DoubleChar$(A) + DoubleChar$(CLng(.Picture)) + Chr$(.Type) + Chr$(.Data(0)) + Chr$(.Data(1)) + Chr$(.Data(2)) + Chr$(.flags) + Chr$(.ClassReq) + Chr$(.LevelReq) + Chr$(.Version) + DoubleChar$(.SellPrice) + .Name
                    End With
                End If

            Case 80    'Request NPC
                If Len(St) > 0 Then
                    A = GetInt(Mid$(St, 1, 2))
                    St2 = vbNullString
                    With NPC(A)
                        For B = 0 To 9
                            St2 = St2 + DoubleChar$(.SaleItem(B).GiveObject) + QuadChar$(.SaleItem(B).GiveValue) + DoubleChar$(.SaleItem(B).TakeObject) + QuadChar$(.SaleItem(B).TakeValue)
                        Next B
                        SendSocket Index, Chr$(85) + DoubleChar$(A) + Chr$(.Version) + Chr$(.flags) + St2 + .Name
                    End With
                End If

            Case 81    'Request Hall
                If Len(St) > 0 Then
                    A = Asc(Mid$(St, 1, 1))
                    With Hall(A)
                        SendSocket Index, Chr$(82) + Chr$(A) + Chr$(.Version) + .Name
                    End With
                End If

            Case 82    'Request Monster
                If Len(St) > 0 Then
                    A = GetInt(Mid$(St, 1, 2))
                    With Monster(A)
                        SendSocket Index, Chr$(32) + DoubleChar$(A) + DoubleChar$(CLng(.Sprite)) + Chr$(.Version) + DoubleChar$(CLng(.HP)) + Chr$(.flags) + .Name
                    End With
                End If
                
            Case 83    'Request Magic
                If Len(St) > 0 Then
                    A = GetInt(Mid$(St, 1, 2))
                    With Magic(A)
                        SendSocket Index, Chr$(127) + DoubleChar$(A) + Chr$(.Class) + Chr$(.Level) + Chr$(.Version) + DoubleChar$(CLng(.Icon)) + Chr$(.IconType) + DoubleChar$(CLng(.CastTimer)) + .Name + vbNullChar + .Description
                    End With
                End If

            Case 84    'Request Prefix
                If Len(St) > 0 Then
                    A = Asc(Mid$(St, 1, 1))
                    With ItemPrefix(A)
                        SendSocket Index, Chr$(133) + Chr$(A) + Chr$(.ModificationType) + Chr$(.ModificationValue) + Chr$(.OccursNaturally) + Chr$(.Version) + .Name
                    End With
                End If

            Case 85    'Request Suffix
                If Len(St) > 0 Then
                    A = Asc(Mid$(St, 1, 1))
                    With ItemSuffix(A)
                        SendSocket Index, Chr$(134) + Chr$(A) + Chr$(.ModificationType) + Chr$(.ModificationValue) + Chr$(.OccursNaturally) + Chr$(.Version) + .Name
                    End With
                End If
                
            Case 98    'Ban
                If Len(St) > 0 Then
                    If Len(St) < 255 Then
                        BanPlayer Index, 0, 1, St, "Server"
                    Else
                        BanPlayer Index, 0, 1, "Cheating", "Server"
                    End If
                    PrintCheat .Name & "  " & .IP & "  " & St
                End If

            Case 99    'Cheat Report
                If Len(St) > 0 Then
                    PrintCheat .Name & "  " & .IP & "  " & St
                End If
                
            Case 100    'Debug Report
                If Len(St) > 0 Then
                    PrintDebug "(Player Submitted) - " + .Name + "  " + .IP + "  " & St
                End If
                
            Case Else
                Hacker Index, "B.2 - " + CStr(PacketID)
            End Select
        Case modePlaying
            Select Case PacketID
            Case 3    'Change Password
                If Len(St) > 0 Then
                    UserRS.Bookmark = .Bookmark
                    UserRS.Edit
                    PrintPassword .Name + " change password " + UserRS!Password + " to " + St
                    UserRS!Password = St
                    UserRS.Update
                Else
                    Hacker Index, "A.6"
                End If

            Case 4    'Request Map
                If Len(St) = 2 And .Access > 0 Then
                    .FloodTimer = 0
                    A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                    MapRS.Seek "=", A
                    If MapRS.NoMatch Then
                        SendSocket Index, Chr$(21) + String$(2677, vbNullChar)
                    Else
                        SendSocket Index, Chr$(21) + MapRS!Data
                    End If
                Else
                    Hacker Index, "A.12"
                End If

            Case 6    'Say
                If Not .Flag(40) > 0 And Not .Flag(41) > 0 And .IsDead = False Then
                    .FloodTimer = .FloodTimer + 1000
                    If Len(St) >= 1 And Len(St) <= 512 Then
                        A = SysAllocStringByteLen(St, Len(St))
                        Parameter(0) = Index
                        Parameter(1) = MapNum
                        Parameter(2) = A
                        B = RunScript("MAPSAY")
                        SysFreeString A

                        PrintChat "Say", .Name + " says, '" + St + "'"

                        If B = 0 Then
                            SendToMapAllBut MapNum, Index, Chr$(11) + Chr$(Index) + St

                            If Int(Rnd * 100) <= 8 Then
                                A = Map(MapNum).NPC
                                If A >= 1 Then
                                    With NPC(A)
                                        B = Int(Rnd * 5)
                                        If .SayText(B) <> "" Then
                                            SendToMap MapNum, Chr$(88) + DoubleChar$(A) + .SayText(B)
                                        End If
                                    End With
                                End If
                            End If
                        End If
                    Else
                        Hacker Index, "A.15"
                    End If
                Else
                    SendSocket Index, Chr$(16) + Chr$(41)
                End If
            Case 7    'Move
                If Len(St) = 4 Then
                    If .LastMsg > .WalkTimer Then
                        .WalkCount = 0
                        .WalkTimer = .LastMsg + 500
                    End If
                    'If .WalkCount > 5 And .Access = 0 Then

                    'Else
                        '.WalkCount = .WalkCount + 1
                        If .IsDead = False Then
                            ProcessMovement Index, St, MapNum
                        End If
                    'End If
                Else
                    Hacker Index, "A.16"
                End If

            Case 8    'Pick up map object
                If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))    'map object #
                    If A <= MaxMapObjects And .IsDead = False Then
                        C = Map(MapNum).Object(A).Object
                        If C > 0 Then
                            If Map(MapNum).Object(A).X = .X And Map(MapNum).Object(A).Y = .Y Then
                                Parameter(0) = Index
                                Parameter(1) = C
                                Parameter(2) = Map(MapNum).Object(A).Value
                                If RunScript("GETOBJ") = 0 Then
                                    If .Access > 0 Then PrintGod .User, " (Pick up) Object: " + Object(Map(MapNum).Object(A).Object).Name + "   Value: " + CStr(Map(MapNum).Object(A).Value)
                                    PrintItem .User + " - " + .Name + " (Pick up) " + Object(Map(MapNum).Object(A).Object).Name + " (" + CStr(Map(MapNum).Object(A).Value) + ") - Map: " + CStr(.Map)
                                    If Object(C).Type = 6 Or Object(C).Type = 11 Then
                                        'Money
                                        B = FindInvObject(Index, C)
                                        If B = 0 Then
                                            B = FreeInvNum(Index)
                                            E = 0
                                        Else
                                            E = 1
                                        End If
                                    Else
                                        B = FreeInvNum(Index)
                                        E = 0
                                    End If
                                    If B > 0 Then
                                        With .Inv(B)
                                            .Object = C
                                            .ItemPrefix = Map(MapNum).Object(A).ItemPrefix
                                            .ItemSuffix = Map(MapNum).Object(A).ItemSuffix
                                            If E = 1 Then
                                                D = .Value + Map(MapNum).Object(A).Value
                                                .Value = D
                                                G = 0
                                            Else
                                                D = Map(MapNum).Object(A).Value
                                                .Value = D
                                                G = 0
                                            End If
                                        End With
                                        If G = 0 Then
                                            Map(MapNum).Object(A).Object = 0
                                            Map(MapNum).Object(A).Value = 0
                                            Map(MapNum).Object(A).ItemPrefix = 0
                                            Map(MapNum).Object(A).ItemSuffix = 0
                                            SendToMap MapNum, Chr$(15) + Chr$(A)    'Erase Map Obj
                                        End If
                                        SendSocket Index, Chr$(17) + Chr$(B) + DoubleChar$(C) + QuadChar(D) + Chr$(.Inv(B).ItemPrefix) + Chr$(.Inv(B).ItemSuffix)    'New Inv Obj
                                    Else
                                        SendSocket Index, Chr$(16) + Chr$(1)    'Inv Full
                                    End If
                                End If
                            Else
                                SendSocket Index, Chr$(16) + Chr$(3)    'No such object
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(3)    'No such object
                        End If
                    End If
                Else
                    Hacker Index, "A.17"
                End If

            Case 9    'Drop Object
                ProcessDropObject Index, MapNum, St

            Case 10    'Use Object
                If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 And A <= 20 Then
                        If .Inv(A).Object > 0 Then
                            Parameter(0) = Index
                            Parameter(1) = .Inv(A).Object
                            If RunScript("USEOBJ") = 0 Then
                                If .Inv(A).Object > 0 Then
                                    If Not ExamineBit(Object(.Inv(A).Object).ClassReq, .Class - 1) = 255 Then
                                        Select Case Object(.Inv(A).Object).Type
                                        Case 1    'Weapon
                                            B = 1
                                            C = 0
                                            If .EquippedObject(6).Object > 0 Then
                                                SendSocket Index, Chr$(20) + Chr$(.EquippedObject(6).Object)    'Stop Using Object
                                                .EquippedObject(6).Object = 0
                                                .EquippedObject(6).ItemPrefix = 0
                                                .EquippedObject(6).ItemSuffix = 0
                                            End If
                                        Case 2    'Shield
                                            B = 2
                                            C = 0
                                        Case 3    'Armor
                                            B = 3
                                            C = 0
                                        Case 4    'Helmut
                                            B = 4
                                            C = 0
                                        Case 5    'Potion
                                            B = Object(.Inv(A).Object).Data(1)
                                            Select Case Object(.Inv(A).Object).Data(0)
                                            Case 0    'Gives HP
                                                If CLng(.HP) + B < .MaxHP Then
                                                    .HP = .HP + B
                                                Else
                                                    .HP = .MaxHP
                                                End If
                                                SendToMap .Map, Chr$(111) + Chr$(10) + Chr$(B) + Chr$(.X) + Chr$(.Y)
                                                SendSocket Index, Chr$(46) + Chr$(.HP)
                                            Case 1    'Takes HP
                                                If CLng(.HP) - B > 0 Then
                                                    .HP = .HP - B
                                                Else
                                                    .HP = 0
                                                End If
                                                SendSocket Index, Chr$(46) + Chr$(.HP)
                                            Case 2    'Gives Mana
                                                If CLng(.Mana) + B < .MaxMana Then
                                                    .Mana = .Mana + B
                                                Else
                                                    .Mana = .MaxMana
                                                End If
                                                SendSocket Index, Chr$(48) + Chr$(.Mana)
                                                SendToMap .Map, Chr$(111) + Chr$(9) + Chr$(B) + Chr$(.X) + Chr$(.Y)
                                            Case 3    'Takes Mana
                                                If CLng(.Mana) - B > 0 Then
                                                    .Mana = .Mana - B
                                                Else
                                                    .Mana = 0
                                                End If
                                                SendSocket Index, Chr$(48) + Chr$(.Mana)
                                            Case 4    'Gives Energy
                                                If CLng(.Energy) + B < .MaxEnergy Then
                                                    .Energy = .Energy + B
                                                Else
                                                    .Energy = .MaxEnergy
                                                End If
                                                SendSocket Index, Chr$(47) + Chr$(.Energy)
                                                SendToMap .Map, Chr$(111) + Chr$(14) + Chr$(B) + Chr$(.X) + Chr$(.Y)
                                            Case 5    'Takes Energy
                                                If CLng(.Energy) - B > 0 Then
                                                    .Energy = .Energy - B
                                                Else
                                                    .Energy = 0
                                                End If
                                                SendSocket Index, Chr$(47) + Chr$(.Energy)
                                            End Select
                                            B = 0
                                            C = 1
                                        Case 7    'Key
                                            Select Case .D
                                            Case 0    'Up
                                                C = .X
                                                D = CLng(.Y) - 1
                                            Case 1    'Down
                                                C = .X
                                                D = .Y + 1
                                            Case 2    'Left
                                                C = CLng(.X) - 1
                                                D = .Y
                                            Case 3    'Right
                                                C = .X + 1
                                                D = .Y
                                            End Select
                                            If C >= 0 And C <= 11 And D >= 0 And D <= 11 Then
                                                If Map(MapNum).Tile(C, D).Att = 3 Then
                                                    If (Map(MapNum).Tile(C, D).AttData(0) + (Map(MapNum).Tile(C, D).AttData(1) * 256)) = .Inv(A).Object Then
                                                        E = FreeMapDoorNum(MapNum)
                                                        If E >= 0 Then
                                                            With Map(MapNum).Door(E)
                                                                .Att = 3
                                                                .X = C
                                                                .Y = D
                                                                .T = Player(Index).LastMsg
                                                            End With
                                                            Map(MapNum).Tile(C, D).Att = 0
                                                            SendToMap MapNum, Chr$(36) + Chr$(E) + Chr$(C) + Chr$(D)
                                                            If Object(.Inv(A).Object).Data(0) = 0 Then
                                                                C = 1
                                                            Else
                                                                C = 0
                                                            End If
                                                        End If
                                                    Else
                                                        C = 0
                                                    End If
                                                Else
                                                    C = 0
                                                End If
                                            Else
                                                C = 0
                                            End If
                                            B = 0
    
                                        Case 8    'Ring
                                            B = 5
                                            C = 0
    
                                        Case 9    'Guild Deed
                                            If Player(Index).Guild = 0 Then
                                                B = 0
                                                C = 0
                                                SendSocket Index, Chr$(97) + Chr$(1)
                                            Else
                                                B = 0
                                                C = 0
                                                SendSocket Index, Chr$(97) + vbNullChar
                                            End If
    
                                        Case 10    'Projectile Weapon
                                            If .EquippedObject(6).Object > 0 Then
                                                SendSocket Index, Chr$(20) + Chr$(.EquippedObject(6).Object)    'Stop Using Object
                                                .EquippedObject(6).Object = 0
                                                .EquippedObject(6).ItemPrefix = 0
                                                .EquippedObject(6).ItemSuffix = 0
                                            End If
                                            B = 1
                                            C = 0
                                        Case 11    'Ammo!
                                            If .EquippedObject(1).Object > 0 Then
                                                E = .EquippedObject(1).Object
                                                If E > 0 Then
                                                    If Object(E).Type = 10 Then    'Using Projectile Weapon
                                                        If Object(E).Data(1) = .Inv(A).Object Or Object(E).Data(2) = .Inv(A).Object Or Object(E).Data(3) = .Inv(A).Object Then
                                                            B = 6
                                                            C = 0
                                                        Else
                                                            SendSocket Index, Chr$(16) & Chr$(37)
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(16) & Chr$(36)
                                                    End If
                                                End If
                                            Else
                                                SendSocket Index, Chr$(16) & Chr$(37)
                                            End If
                                        Case Else
                                            B = 0
                                            C = 0
                                            SendSocket Index, Chr$(16) + Chr$(8)    'You cannot use that
                                        End Select
                                        If B > 0 Then
                                            'Equip Item
                                            EquipObject Index, A
                                        End If
                                        If C > 0 Then
                                            'Destroy Item
                                            .Inv(A).Object = 0
                                            .Inv(A).Value = 0
                                            .Inv(A).ItemPrefix = 0
                                            .Inv(A).ItemSuffix = 0
                                            SendSocket Index, Chr$(18) + Chr$(A)    'Remove inv object
                                        End If
                                    Else
                                        SendSocket Index, Chr$(16) + Chr$(40)    'No such object
                                    End If
                                End If
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(3)    'No such object
                        End If
                    Else
                        Hacker Index, "A.19"
                    End If
                Else
                    Hacker Index, "A.20"
                End If

            Case 11    'Stop Using Object
                If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 And A <= 6 Then
                        UnEquipObject Index, A
                    Else
                        Hacker Index, "A.21"
                    End If
                End If

            Case 12    'Upload Map
                If Len(St) = 2677 And .Access >= 1 Then
                    PrintGod .User, " (Change Map) Map: " + CStr(.Map) + " - " + Map(.Map).Name + " - Objects: " + GetMapObjectList(.Map) + " - Warps: " + GetMapWarpList(.Map)
                    St2 = CompressString(St)
                    MapRS.Seek "=", MapNum
                    If MapRS.NoMatch Then
                        MapRS.AddNew
                        MapRS!number = MapNum
                    Else
                        MapRS.Edit
                    End If
                    MapRS!Data = St2
                    MapRS.Update
                    LoadMap MapNum, St
                    For A = 0 To 9
                        Map(MapNum).Door(A).Att = 0
                    Next A
                    For A = 1 To MaxUsers
                        With Player(A)
                            If .Mode = modePlaying And .Map = MapNum Then
                                Partmap A
                                .Map = MapNum
                                JoinMap A
                            End If
                        End With
                    Next A
                Else
                    Hacker Index, "A.22"
                End If

            Case 13    'Exit Map
                If Len(St) = 1 Then
                    Select Case Asc(Mid$(St, 1, 1))
                    Case 0
                        If Map(MapNum).ExitUp > 0 And Map(MapNum).ExitUp <= MaxMaps Then
                            If .Y = 0 Then
                                Partmap Index
                                .Map = Map(MapNum).ExitUp
                                .Y = 11
                                JoinMap Index
                            Else
                                PlayerWarp Index, .Map, .X, .Y
                            End If
                        Else
                            Partmap Index
                            .Map = MapNum
                            JoinMap Index
                        End If
                    Case 1
                        If Map(MapNum).ExitDown > 0 And Map(MapNum).ExitDown <= MaxMaps Then
                            If .Y = 11 Then
                                Partmap Index
                                .Map = Map(MapNum).ExitDown
                                .Y = 0
                                JoinMap Index
                            Else
                                PlayerWarp Index, .Map, .X, .Y
                            End If
                        Else
                            Partmap Index
                            .Map = MapNum
                            JoinMap Index
                        End If
                    Case 2
                        If Map(MapNum).ExitLeft > 0 And Map(MapNum).ExitLeft <= MaxMaps Then
                            If .X = 0 Then
                                Partmap Index
                                .Map = Map(MapNum).ExitLeft
                                .X = 11
                                JoinMap Index
                            Else
                                PlayerWarp Index, .Map, .X, .Y
                            End If
                        Else
                            Partmap Index
                            .Map = MapNum
                            JoinMap Index
                        End If
                    Case 3
                        If Map(MapNum).ExitRight > 0 And Map(MapNum).ExitRight <= MaxMaps Then
                            If .X = 11 Then
                                Partmap Index
                                .Map = Map(MapNum).ExitRight
                                .X = 0
                                JoinMap Index
                            Else
                                PlayerWarp Index, .Map, .X, .Y
                            End If
                        Else
                            Partmap Index
                            .Map = MapNum
                            JoinMap Index
                        End If
                    End Select
                Else
                    Hacker Index, "A.23"
                End If

            Case 14    'Tell
                If Not .Flag(40) > 0 And Not .Flag(41) > 0 Then
                    .FloodTimer = .FloodTimer + 1000
                    If Len(St) >= 2 And Len(St) <= 513 Then
                        A = Asc(Mid$(St, 1, 1))
                        If A >= 1 And A <= MaxUsers Then
                            If Player(A).Mode = modePlaying Then
                                PrintChat "Tell", .Name + " tells " + Player(A).Name + ", '" + Mid$(St, 2) + "'"
                                SendSocket A, Chr$(25) + Chr$(Index) + Mid$(St, 2)
                                If .Mana > 2 Then
                                    .Mana = .Mana - 2
                                Else
                                    .Mana = 0
                                End If
                                SendSocket Index, Chr$(48) + Chr$(.Mana)
                            End If
                        End If
                    Else
                        Hacker Index, "A.24"
                    End If
                Else
                    SendSocket Index, Chr$(16) + Chr$(41)
                End If
            Case 15    'Broadcast
                If Len(St) >= 1 And Len(St) <= 512 Then
                    If Not .Flag(40) > 0 And Not .Flag(41) > 0 And .IsDead = False Then
                        .FloodTimer = .FloodTimer + 1500

                        A = SysAllocStringByteLen(St, Len(St))
                        Parameter(0) = Index
                        Parameter(1) = A
                        SysFreeString A

                        If RunScript("BROADCAST") = 0 Then
                            SendAllBut Index, Chr$(26) + Chr$(Index) + St
                            PrintLog .Name + ": " + St
                            PrintChat "Broadcast", .Name + ": " + St
                            If .Mana > 5 Then
                                .Mana = .Mana - 5
                            Else
                                .Mana = 0
                            End If
                            SendSocket Index, Chr$(48) + Chr$(.Mana)
                        End If
                    Else
                        SendSocket Index, Chr$(16) + Chr$(41)
                    End If
                Else
                    Hacker Index, "A.25"
                End If

            Case 16    'Emote
                If Not .Flag(40) > 0 And Not .Flag(41) > 0 And .IsDead = False Then
                    .FloodTimer = .FloodTimer + 1000
                    If Len(St) >= 1 And Len(St) <= 512 Then
                        SendToMapAllBut MapNum, Index, Chr$(27) + Chr$(Index) + St
                        PrintChat "Emote", .Name + " emotes, '" + St + "'"
                    Else
                        Hacker Index, "A.26"
                    End If
                Else
                    SendSocket Index, Chr$(16) + Chr$(41)
                End If
            Case 17    'Yell
                If Not .Flag(40) > 0 And Not .Flag(41) > 0 And .IsDead = False Then
                    .FloodTimer = .FloodTimer + 1200
                    If Len(St) >= 1 And Len(St) <= 512 Then
                        SendToMapAllBut MapNum, Index, Chr$(28) + Chr$(Index) + St
                        A = MapNum
                        With Map(MapNum)
                            B = .ExitUp
                            C = .ExitDown
                            D = .ExitLeft
                            E = .ExitRight
                        End With
                        If B <> MapNum Then SendToMap B, Chr$(28) + Chr$(Index) + St
                        If C <> MapNum And C <> B Then SendToMap C, Chr$(28) + Chr$(Index) + St
                        If D <> MapNum And D <> B And D <> C Then SendToMap D, Chr$(28) + Chr$(Index) + St
                        If E <> MapNum And E <> B And E <> C And E <> D Then SendToMap E, Chr$(28) + Chr$(Index) + St
                        PrintChat "Yell", .Name + " yells, '" + St + "'"
                    Else
                        Hacker Index, "A.27"
                    End If
                Else
                    SendSocket Index, Chr$(16) + Chr$(41)
                End If

            Case 18    'God Commands
                If .Access > 0 And Len(St) >= 1 Then
                    ProcessGodCommand Index, St
                Else
                    Hacker Index, "A.42"
                End If

            Case 19    'Edit Object
                If Len(St) = 2 And .Access > 0 Then
                    A = GetInt(Mid$(St, 1, 2))
                    If A >= 1 Then
                        With Object(A)
                            SendSocket Index, Chr$(33) + DoubleChar$(A) + Chr$(.flags) + Chr$(.Data(0)) + Chr$(.Data(1)) + Chr$(.Data(2)) + Chr$(.Data(3)) + Chr$(.ClassReq) + Chr$(.LevelReq)
                        End With
                    End If
                Else
                    Hacker Index, "A.43"
                End If

            Case 20    'Edit Monster
                If Len(St) = 2 And .Access > 0 Then
                    A = GetInt(Mid$(St, 1, 2))
                    If A >= 1 And A <= MaxTotalMonsters Then
                        With Monster(A)
                            SendSocket Index, Chr$(34) + DoubleChar$(A) + DoubleChar$(CLng(.HP)) + Chr$(.Strength) + Chr$(.Armor) + Chr$(.Speed) + Chr$(.Sight) + Chr$(.Agility) + Chr$(.flags) + DoubleChar$(.Object(0)) + Chr$(.Value(0)) + DoubleChar$(.Object(1)) + Chr$(.Value(1)) + DoubleChar$(.Object(2)) + Chr$(.Value(2)) + Chr$(.Experience / 10) + Chr$(.MagicDefense)
                        End With
                    End If
                Else
                    Hacker Index, "A.44"
                End If

            Case 21    'Save Object
                If Len(St) >= 8 And .Access >= 2 Then
                    A = GetInt(Mid$(St, 1, 2))
                    If A >= 1 Then
                        PrintGod .User, " (Save Object) Object #: " + CStr(A)
                        With Object(A)
                            .Picture = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                            .Type = Asc(Mid$(St, 5, 1))
                            .flags = Asc(Mid$(St, 6, 1))
                            .Data(0) = Asc(Mid$(St, 7, 1))
                            .Data(1) = Asc(Mid$(St, 8, 1))
                            .Data(2) = Asc(Mid$(St, 9, 1))
                            .Data(3) = Asc(Mid$(St, 10, 1))
                            .ClassReq = Asc(Mid$(St, 11, 1))
                            .LevelReq = Asc(Mid$(St, 12, 1))
                            .SellPrice = Asc(Mid$(St, 13, 1)) * 256 + Asc(Mid$(St, 14, 1))
                            If Not .Version = 255 Then .Version = .Version + 1 Else .Version = 1
                            If Len(St) >= 15 Then
                                .Name = Mid$(St, 15)
                            Else
                                .Name = ""
                            End If
                            ObjectRS.Seek "=", A
                            If ObjectRS.NoMatch Then
                                ObjectRS.AddNew
                                ObjectRS!number = A
                            Else
                                ObjectRS.Edit
                            End If
                            ObjectRS!Name = .Name
                            ObjectRS!Picture = .Picture
                            ObjectRS!Type = .Type
                            ObjectRS!flags = .flags
                            ObjectRS!Data1 = .Data(0)
                            ObjectRS!Data2 = .Data(1)
                            ObjectRS!Data3 = .Data(2)
                            ObjectRS!Data4 = .Data(3)
                            ObjectRS!ClassReq = .ClassReq
                            ObjectRS!LevelReq = .LevelReq
                            ObjectRS!Version = .Version
                            ObjectRS!SellPrice = .SellPrice
                            ObjectRS.Update
                            SendAll Chr$(31) + DoubleChar$(A) + DoubleChar$(CLng(.Picture)) + Chr$(.Type) + Chr$(.Data(0)) + Chr$(.Data(1)) + Chr$(.Data(2)) + Chr$(.flags) + Chr$(.ClassReq) + Chr$(.LevelReq) + Chr$(.Version) + DoubleChar$(.SellPrice) + .Name
                            GenerateObjectVersionList
                        End With
                    End If
                Else
                    Hacker Index, "A.45"
                End If

            Case 22    'Save Monster
                If Len(St) >= 16 And .Access >= 2 Then
                    ProcessSaveMonster Index, St
                Else
                    Hacker Index, "A.46"
                End If

            Case 24    'Scan Echo
                If Len(St) > 2 Then
                    A = Asc(Mid$(St, 1, 1))
                    SendSocket A, Chr$(102) + Mid$(St, 2)
                End If

            Case 25    'Attack Player
                If Len(St) = 1 Then
                    If .LastMsg > .TimeLeft And .IsDead = False Then
                        If ExamineBit(Map(MapNum).flags, 0) = False Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= MaxUsers Then
                                If Player(A).IsDead = False Then
                                    .TimeLeft = .LastMsg + 850
                                    If NoDirectionalWalls(CLng(.Map), CLng(.X), CLng(.Y), CLng(.D)) Then CombatAttackPlayer Index, A, PlayerDamage(Index)
                                End If
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(9)    'Friendly Zone
                        End If
                    End If
                Else
                    Hacker Index, "A.47"
                End If

            Case 26    'Attack Monster
                If Len(St) = 1 Then
                    If .LastMsg > .TimeLeft And .IsDead = False Then
                        If ExamineBit(Map(MapNum).flags, 5) = False Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= MaxMonsters Then
                                If Map(MapNum).Monster(A).Monster > 0 Then
                                    If Sqr((CSng(Map(MapNum).Monster(A).X) - CSng(.X)) ^ 2 + (CSng(Map(MapNum).Monster(A).Y) - CSng(.Y)) ^ 2) <= LagHitDistance Then
                                        If NoDirectionalWalls(CLng(.Map), CLng(.X), CLng(.Y), CLng(.D)) Then
                                            Parameter(0) = Index
                                            Parameter(1) = Map(MapNum).Monster(A).Monster
                                            Parameter(2) = .Map
                                            Parameter(3) = A
                                            If RunScript("ATTACKMONSTER") = 0 Then
                                        
                                                .TimeLeft = .LastMsg + 850
                                                With Monster(Map(MapNum).Monster(A).Monster)
                                                    Dim AgilityChance As Integer
                                                    If CInt(.Agility) - CInt(statPlayerAgility) > 0 Then AgilityChance = CInt(.Agility) - CInt(statPlayerAgility) Else AgilityChance = 0
                                                    If Int(Rnd * 100) > AgilityChance Then
                                                        'Hit Target
                                                        B = 0
                                                        C = PlayerDamage(Index) - .Armor
                                                        If C < 0 Then C = 0
                                                        If C > 255 Then C = 255
                                                    Else
                                                        'Missed
                                                        B = 1
                                                        C = 0
                                                    End If
                                                End With

                                                With Map(MapNum).Monster(A)
                                                    .Target = Index
                                                    .TargetIsMonster = False
                                                    If .HP > C Then
                                                        .HP = .HP - C
                                                        'Attacked Monster
                                                        SendToMap MapNum, Chr$(44) + Chr$(Index) + Chr$(B) + Chr$(A) + Chr$(C) + DoubleChar$(CLng(.HP))
                                                    Else
                                                        'Attacked Monster
                                                        SendToMap MapNum, Chr$(44) + Chr$(Index) + Chr$(B) + Chr$(A) + Chr$(C) + DoubleChar$(CLng(.HP))

                                                        'Monster Died
                                                        SendToMapAllBut MapNum, Index, Chr$(39) + Chr$(A)    'Monster Died

                                                        'Experience
                                                        If ExamineBit(Monster(.Monster).flags, 4) = False Then
                                                            GainExp Index, CLng(Monster(.Monster).Experience)
                                                        Else
                                                            GainEliteExp Index, CLng(Monster(.Monster).Experience)
                                                        End If
                                                        
                                                        SendSocket Index, Chr$(51) + Chr$(A) + QuadChar(Player(Index).Experience)    'You killed monster

                                                        D = Int(Rnd * 3)
                                                        E = Monster(.Monster).Object(D)
                                                        If E > 0 Then
                                                            NewMapObject MapNum, E, Monster(.Monster).Value(D), CLng(.X), CLng(.Y), False
                                                        End If

                                                        Parameter(0) = Index
                                                        Parameter(1) = .Monster
                                                        Parameter(2) = MapNum
                                                        Parameter(3) = A
                                                        RunScript "MONSTERDIE"
                                                        
                                                        .Monster = 0
                                                    End If
                                                End With
                                            End If 'AttackMonster Script
                                        End If    'Directional Walls Check
                                    Else
                                        SendSocket Index, Chr$(16) + Chr$(7)    'Too far away
                                    End If
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(5)    'No such monster
                                End If
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(12)    'Can't attack monsters here
                        End If
                    End If
                Else
                    Hacker Index, "A.48"
                End If

            Case 27    'Look at player
                If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 And A <= MaxUsers Then
                        Parameter(0) = Index
                        Parameter(1) = A
                        RunScript "CLICKPLAYER"
                        SendSocket Index, Chr$(56) + Chr$(15) + Player(A).Name + "'s Description:  " + Player(A).desc
                    End If
                Else
                    Hacker Index, "A.49"
                End If

            Case 28    'Describe
                If Len(St) >= 1 And Len(St) <= 255 Then
                    .desc = St
                Else
                    Hacker Index, "A.50"
                End If

            Case 29    'Pong
                If Len(St) = 0 Then

                Else
                    SendSocket Index, Chr$(255)
                End If

            Case 31    'Join Guild
                If Len(St) = 0 Then
                    If .Guild = 0 Then
                        If .Level >= World.GuildJoinLevel Then
                            If .JoinRequest > 0 Then
                                If Guild(.JoinRequest).Name <> "" Then
                                    Parameter(0) = Index
                                    If RunScript("GUILDJOIN") = 0 Then
                                        A = FreeGuildMemberNum(CLng(.JoinRequest))
                                        If A >= 0 Then
                                            B = FindInvObject(Index, CLng(World.ObjMoney))
                                            If B > 0 Then
                                                If .Inv(B).Value >= World.GuildJoinPrice Then
                                                    With .Inv(B)
                                                        .Value = .Value - World.GuildJoinPrice
                                                        If .Value <= 0 Then
                                                            .Object = 0
                                                            .ItemPrefix = 0
                                                            .ItemSuffix = 0
                                                        Else
                                                            SendSocket Index, Chr$(17) + Chr$(B) + Chr$(.Object) + QuadChar(.Value) + Chr$(.ItemPrefix) + Chr$(.ItemSuffix)    'Change inv object
                                                        End If
                                                    End With
                                                    .Guild = .JoinRequest
                                                    .JoinRequest = 0
                                                    .GuildRank = 0
                                                    .GuildSlot = A
                                                    If Guild(.Guild).Sprite > 0 Then
                                                        .Sprite = Guild(.Guild).Sprite
                                                        SendToMap .Map, Chr$(63) + Chr$(Index) + DoubleChar$(CLng(.Sprite))
                                                    End If
                                                    Guild(.Guild).Member(A).Name = .Name
                                                    Guild(.Guild).Member(A).Rank = 0
                                                    Guild(.Guild).Member(A).JoinDate = CLng(Date)
                                                    Guild(.Guild).Member(A).Kills = 0
                                                    Guild(.Guild).Member(A).Deaths = 0
    
                                                    GuildRS.Bookmark = Guild(.Guild).Bookmark
                                                    GuildRS.Edit
                                                    GuildRS("MemberName" + CStr(A)) = .Name
                                                    GuildRS("MemberRank" + CStr(A)) = 0
                                                    GuildRS("MemberJoinDate" + CStr(A)) = CLng(Date)
                                                    GuildRS("MemberKills" + CStr(A)) = 0
                                                    GuildRS("MemberDeaths" + CStr(A)) = 0
                                                    GuildRS.Update
    
                                                    Guild(.Guild).MemberCount = CountGuildMembers(CLng(.Guild))
                                                    SendAll Chr$(70) + Chr$(.Guild) + Chr$(Guild(.Guild).MemberCount) + Guild(.Guild).Name
    
                                                    SendSocket Index, Chr$(72) + Chr$(.Guild)    'Change guild
                                                    SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(.Guild)    'Player changed guild
                                                    For A = 0 To DeclarationCount
                                                        With Guild(.Guild).Declaration(A)
                                                            SendSocket Index, Chr$(71) + Chr$(A) + Chr$(.Guild) + Chr$(.Type)
                                                        End With
                                                    Next A
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(15)    'Not enough money
                                                End If
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(15)    'Not enough money
                                            End If
                                        Else
                                            SendSocket Index, Chr$(16) + Chr$(17)    'Guild is full
                                        End If
                                    End If
                                Else
                                    .JoinRequest = 0
                                    SendSocket Index, Chr$(16) + Chr$(14)    'You have not been invited
                                End If
                            Else
                                SendSocket Index, Chr$(16) + Chr$(14)    'You have not been invited
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(30)    'You must be Level 5 to join
                        End If
                    Else
                        SendSocket Index, Chr$(16) + Chr$(16)    'You are already in a guild
                    End If
                Else
                    Hacker Index, "A.53"
                End If

            Case 32    'Leave Guild
                If Len(St) = 0 Then
                    If .Guild > 0 Then
                        Parameter(0) = Index
                        If RunScript("GUILDLEAVE") = 0 Then
                            A = FindGuildMember(.Name, CLng(.Guild))
                            If A >= 0 Then
                                With Guild(.Guild).Member(A)
                                    .Name = ""
                                    .Rank = 0
                                    .JoinDate = 0
                                    .Kills = 0
                                    .Deaths = 0
                                End With
                                GuildRS.Bookmark = Guild(.Guild).Bookmark
                                GuildRS.Edit
                                GuildRS("MemberName" + CStr(A)) = ""
                                GuildRS("MemberRank" + CStr(A)) = 0
                                GuildRS("MemberJoinDate" + CStr(A)) = 0
                                GuildRS("MemberKills" + CStr(A)) = 0
                                GuildRS("MemberDeaths" + CStr(A)) = 0
                                GuildRS.Update
                            End If
                            SendSocket Index, Chr$(72) + vbNullChar
                            SendAllBut Index, Chr$(73) + Chr$(Index) + vbNullChar
                            If Guild(.Guild).Sprite > 0 Then
                                .Sprite = .Class * 2 + .Gender - 1
                                SendToMap .Map, Chr$(63) + Chr$(Index) + DoubleChar$(CLng(.Sprite))
                            End If
    
                            Guild(.Guild).MemberCount = CountGuildMembers(CLng(.Guild))
                            SendAll Chr$(70) + Chr$(.Guild) + Chr$(Guild(.Guild).MemberCount) + Guild(.Guild).Name
    
                            CheckGuild CLng(.Guild)
    
                            .Guild = 0
                        End If
                    End If
                Else
                    Hacker Index, "A.54"
                End If

            Case 33    'Start New Guild
                ProcessStartNewGuild Index, St

            Case 34    'Invite Player to Guild
                If Len(St) = 1 Then
                    If .Guild > 0 Then
                        If .GuildRank >= 2 Then
                            Parameter(0) = Index
                            If RunScript("GUILDINVITE") = 0 Then
                                A = Asc(Mid$(St, 1, 1))
                                If A >= 1 And A <= MaxUsers Then
                                    If Player(A).Mode = modePlaying Then
                                        Player(A).JoinRequest = .Guild
                                        SendSocket A, Chr$(77) + Chr$(.Guild) + Chr$(Index)    'Invited to join guild
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "A.55"
                End If

            Case 35    'Kick player from guild
                If Len(St) = 1 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        If A <= 19 Then
                            B = .Guild
                            If Guild(B).Member(A).Rank <= .GuildRank Then
                                With Guild(B).Member(A)
                                    St1 = .Name
                                    .Name = ""
                                    .Rank = 0
                                    .JoinDate = 0
                                    .Kills = 0
                                    .Deaths = 0
                                End With
                                GuildRS.Bookmark = Guild(B).Bookmark
                                GuildRS.Edit
                                GuildRS("MemberName" + CStr(A)) = ""
                                GuildRS("MemberRank" + CStr(A)) = 0
                                GuildRS("MemberJoinDate" + CStr(A)) = 0
                                GuildRS("MemberKills" + CStr(A)) = 0
                                GuildRS("MemberDeaths" + CStr(A)) = 0
                                GuildRS.Update
                                A = FindPlayer(St1)
                                If A > 0 Then
                                    With Player(A)
                                        .Guild = 0
                                        .GuildRank = 0
                                        SendSocket A, Chr$(72) + vbNullChar
                                        SendAllBut A, Chr$(73) + Chr$(A) + vbNullChar
                                        If Guild(B).Sprite > 0 Then
                                            .Sprite = .Class * 2 + .Gender - 1
                                            SendToMap .Map, Chr$(63) + Chr$(A) + DoubleChar$(CLng(.Sprite))
                                        End If
                                    End With
                                ElseIf Guild(B).Sprite > 0 Then
                                    UserRS.Index = "Name"
                                    UserRS.Seek "=", St1
                                    If UserRS.NoMatch = False Then
                                        A = UserRS!Class * 2 + UserRS!Gender - 1
                                        If A >= 1 And A <= MaxSprite Then
                                            UserRS.Edit
                                            UserRS!Sprite = A
                                            UserRS.Update
                                        End If
                                    End If
                                End If
                                CheckGuild B
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "A.56"
                End If

            Case 36    'Change player's rank
                If Len(St) = 2 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        B = Asc(Mid$(St, 2, 1))
                        D = .Guild
                        If A <= 19 And B <= .GuildRank Then
                            If Guild(D).Member(A).Rank <= .GuildRank Then
                                With Guild(D).Member(A)
                                    If .Name <> "" Then
                                        .Rank = B
                                        C = FindPlayer(.Name)
                                        If C > 0 Then
                                            Player(C).GuildRank = B
                                            SendSocket C, Chr$(76) + Chr$(B)    'Rank Changed
                                        End If
                                    End If
                                End With
                                GuildRS.Bookmark = Guild(D).Bookmark
                                GuildRS.Edit
                                GuildRS("MemberRank" + CStr(A)) = B
                                GuildRS.Update
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "A.57"
                End If

            Case 37    'Add Declaration
                If Len(St) = 2 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        B = Asc(Mid$(St, 2, 1))
                        If A >= 1 And (B = 0 Or B = 1) Then
                            D = .Guild
                            C = FreeGuildDeclarationNum(D)
                            If C >= 0 Then
                                With Guild(D).Declaration(C)
                                    .Guild = A
                                    .Type = B
                                    .Date = CLng(Date)
                                    .Kills = 0
                                    .Deaths = 0
                                End With
                                SendToGuild D, Chr$(71) + Chr$(C) + Chr$(A) + Chr$(B)

                                GuildRS.Bookmark = Guild(D).Bookmark
                                GuildRS.Edit
                                GuildRS("DeclarationGuild" + CStr(C)) = A
                                GuildRS("DeclarationType" + CStr(C)) = B
                                GuildRS("DeclarationDate" + CStr(C)) = CLng(Date)
                                GuildRS("DeclarationKills" + CStr(C)) = 0
                                GuildRS("DeclarationDeaths" + CStr(C)) = 0
                                GuildRS.Update
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "A.58"
                End If

            Case 38    'Remove Declaration
                If Len(St) = 1 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        If A <= 4 Then
                            B = .Guild
                            With Guild(B).Declaration(A)
                                .Guild = 0
                                .Type = 0
                            End With

                            SendToGuild B, Chr$(71) + Chr$(A) + vbNullChar + vbNullChar

                            GuildRS.Bookmark = Guild(B).Bookmark
                            GuildRS.Edit
                            GuildRS("DeclarationGuild" + CStr(A)) = 0
                            GuildRS("DeclarationType" + CStr(A)) = 0
                            GuildRS("DeclarationDate" + CStr(A)) = 0
                            GuildRS("DeclarationKills" + CStr(A)) = 0
                            GuildRS("DeclarationDeaths" + CStr(A)) = 0
                            GuildRS.Update
                        End If
                    End If
                Else
                    Hacker Index, "A.59"
                End If

            Case 39    'View Guild Data
                If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 Then
                        With Guild(A)
                            St1 = DoubleChar$(87) + Chr$(78) + Chr$(A) + DoubleChar$(CLng(.Sprite)) + QuadChar$(.CreationDate) + QuadChar$(.Kills) + QuadChar$(.Deaths) + Chr$(.Hall)
                            For B = 0 To DeclarationCount
                                St1 = St1 + Chr$(.Declaration(B).Guild) + Chr$(.Declaration(B).Type) + QuadChar$(.Declaration(B).Date) + QuadChar$(.Declaration(B).Kills) + QuadChar$(.Declaration(B).Deaths)
                            Next B
                            For B = 0 To 19
                                If Not .Member(B).Name = "" Then
                                    St1 = St1 + DoubleChar$(15 + Len(.Member(B).Name)) + Chr$(144) + Chr$(B) + Chr$(.Member(B).Rank) + QuadChar$(.Member(B).Kills) + QuadChar$(.Member(B).Deaths) + QuadChar$(.Member(B).JoinDate) + .Member(B).Name
                                End If
                            Next B
                        End With
                        SendRaw Index, St1
                    End If
                Else
                    Hacker Index, "A.60"
                End If

            Case 40    'Pay guild balance
                If Len(St) = 4 Then
                    If .Guild > 0 Then
                        If Guild(.Guild).Name <> "" Then
                            C = .Guild
                            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            If A > 0 Then
                                B = FindInvObject(Index, CLng(World.ObjMoney))
                                If B > 0 Then
                                    If .Inv(B).Value >= A Then
                                        With .Inv(B)
                                            .Value = .Value - A
                                            If .Value = 0 Then
                                                .Object = 0
                                                .ItemPrefix = 0
                                                .ItemSuffix = 0
                                            End If
                                            SendSocket Index, Chr$(17) + Chr$(B) + DoubleChar$(CLng(.Object)) + QuadChar$(.Value) + Chr$(.ItemPrefix) + Chr$(.ItemSuffix)    'Change inv object
                                        End With
                                        With Guild(C)
                                            If CSng(.Bank) + CSng(A) >= 2147483647 Then
                                                .Bank = 2147483647
                                            Else
                                                .Bank = .Bank + A
                                            End If

                                            GuildRS.Bookmark = .Bookmark
                                            GuildRS.Edit
                                            GuildRS!Bank = .Bank
                                            GuildRS.Update

                                            If .Bank >= 0 Then
                                                SendToGuild C, Chr$(152) + QuadChar$(.Bank) + QuadChar$(GetGuildUpkeep(C)) + Chr$(Index) + QuadChar$(A)
                                            Else
                                                SendToGuild C, Chr$(74) + QuadChar$(Abs(.Bank)) + QuadChar$(.DueDate) + Chr$(Index) + QuadChar$(A)
                                            End If
                                        End With
                                    Else
                                        SendSocket Index, Chr$(16) + Chr$(15)    'Not enough money
                                    End If
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(15)    'Not enough money
                                End If
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "A.61"
                End If

            Case 41    'Guild Chat
                If Not .Flag(40) > 0 And Not .Flag(41) > 0 And .IsDead = False Then
                    .FloodTimer = .FloodTimer + 1000
                    If Len(St) >= 1 Then
                        If .Guild > 0 Then
                            PrintChat "Guild", .Name + " <" + Guild(.Guild).Name + ">: " + St
                            SendToGuildAllBut Index, CLng(.Guild), Chr$(79) + Chr$(Index) + St
                        End If
                    Else
                        Hacker Index, "A.62"
                    End If
                Else
                    SendSocket Index, Chr$(16) + Chr$(41)
                End If

            Case 42    'Disband Guild
                If Len(St) = 0 Then
                    If .Guild > 0 And .GuildRank = 3 Then
                        DeleteGuild CLng(.Guild), 2
                    End If
                Else
                    Hacker Index, "A.63"
                End If

            Case 43    'Buy guild hall
                If Len(St) = 0 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        D = .Guild
                        If Guild(D).Hall = 0 Then
                            A = Map(MapNum).Hall
                            If A > 0 Then
                                C = 0
                                For B = 1 To MaxGuilds
                                    With Guild(B)
                                        If .Name <> "" And .Hall = A Then
                                            C = 1
                                            Exit For
                                        End If
                                    End With
                                Next B
                                If C = 0 Then
                                    With Guild(D)
                                        If CountGuildMembers(D) >= 5 Then
                                            If .Bank >= Hall(A).Price Then
                                                .Bank = .Bank - Hall(A).Price
                                                SendToGuild D, Chr$(154) + QuadChar(.Bank) + QuadChar$(GetGuildUpkeep(D))
                                                .Hall = Map(Player(Index).Map).Hall
                                                SendToGuild D, Chr$(81) + Chr$(0)
                                                GuildRS.Bookmark = .Bookmark
                                                GuildRS.Edit
                                                GuildRS!Bank = .Bank
                                                GuildRS!Hall = .Hall
                                                GuildRS.Update
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(24)    'Cost 20k to buy hall
                                            End If
                                        Else
                                            SendSocket Index, Chr$(16) + Chr$(26)    'Need 3 members
                                        End If
                                    End With
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(22)    'Hall already owned
                                End If
                            Else
                                SendSocket Index, Chr$(16) + Chr$(21)    'Not in a hall
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(23)    'Already have a hall
                        End If
                    End If
                Else
                    Hacker Index, "A.64"
                End If

            Case 44    'Leave guild hall
                If Len(St) = 0 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = .Guild
                        With Guild(A)
                            If .Hall > 0 Then
                                .Hall = 0
                                GuildRS.Bookmark = .Bookmark
                                GuildRS.Edit
                                GuildRS!Hall = 0
                                GuildRS.Update
                                SendToGuild A, Chr$(81) + Chr$(1)
                            End If
                        End With
                    End If
                Else
                    Hacker Index, "A.65"
                End If

            Case 45    'Request Map
                If Len(St) = 0 Then
                    MapRS.Seek "=", .Map
                    If MapRS.NoMatch Then
                        SendSocket Index, Chr$(21) + String$(2388, vbNullChar)
                    Else
                        SendSocket Index, Chr$(21) + MapRS!Data
                    End If
                Else
                    Hacker Index, "A.66"
                End If

            Case 46    'Guild Balance
                If Len(St) = 0 Then
                    If .Guild > 0 Then
                        With Guild(.Guild)
                            If .Bank >= 0 Then
                                SendSocket Index, Chr$(152) + QuadChar(.Bank) + QuadChar$(GetGuildUpkeep(CLng(Player(Index).Guild)))
                            Else
                                SendSocket Index, Chr$(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                            End If
                        End With
                    End If
                Else
                    Hacker Index, "A.67"
                End If

            Case 47    'Guild Hall Info
                If Len(St) = 0 Then
                    A = Map(Player(Index).Map).Hall
                    If A >= 1 Then
                        With Hall(A)
                            C = 0
                            For B = 1 To MaxGuilds
                                With Guild(B)
                                    If .Name <> "" And .Hall = A Then
                                        C = B
                                        Exit For
                                    End If
                                End With
                            Next B
                            SendSocket Index, Chr$(84) + Chr$(A) + Chr$(C) + QuadChar(Hall(A).Price) + QuadChar(Hall(A).Upkeep)
                        End With
                    Else
                        SendSocket Index, Chr$(16) + Chr$(21)    'Not in a hall
                    End If
                Else
                    Hacker Index, "A.68"
                End If

            Case 48    'Edit Guild Hall
                If Len(St) = 1 And .Access >= 2 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 Then
                        With Hall(A)
                            SendSocket Index, Chr$(83) + Chr$(A) + QuadChar(.Price) + QuadChar(.Upkeep) + DoubleChar(CLng(.StartLocation.Map)) + Chr$(.StartLocation.X) + Chr$(.StartLocation.Y)
                        End With
                    End If
                Else
                    Hacker Index, "A.69"
                End If

            Case 49    'Upload Guild hall data
                If Len(St) >= 13 And Len(St) <= 28 And .Access >= 2 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 Then
                        With Hall(A)
                            .Price = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                            .Upkeep = Asc(Mid$(St, 6, 1)) * 16777216 + Asc(Mid$(St, 7, 1)) * 65536 + Asc(Mid$(St, 8, 1)) * 256& + Asc(Mid$(St, 9, 1))
                            With .StartLocation
                                .Map = Asc(Mid$(St, 10, 1)) * 256 + Asc(Mid$(St, 11, 1))
                                .X = Asc(Mid$(St, 12, 1))
                                .Y = Asc(Mid$(St, 13, 1))
                            End With
                            If Len(St) >= 14 Then
                                .Name = Mid$(St, 14)
                            Else
                                .Name = ""
                            End If
                            HallRS.Seek "=", A
                            If HallRS.NoMatch = True Then
                                HallRS.AddNew
                                HallRS!number = A
                            Else
                                HallRS.Edit
                            End If
                            HallRS!Name = .Name
                            HallRS!Price = .Price
                            HallRS!Upkeep = .Upkeep
                            HallRS!StartLocationMap = .StartLocation.Map
                            HallRS!StartLocationX = .StartLocation.X
                            HallRS!StartLocationY = .StartLocation.Y
                            If Not .Version = 255 Then .Version = .Version + 1 Else .Version = 1
                            HallRS!Version = .Version
                            HallRS.Update

                            SendAll Chr$(82) + Chr$(A) + Chr$(.Version) + .Name
                            GenerateHallVersionList
                        End With
                    End If
                Else
                    Hacker Index, "A.70"
                End If

            Case 50    'Edit NPC Data
                If Len(St) = 2 And .Access >= 2 Then
                    A = GetInt(Mid$(St, 1, 2))
                    If A >= 1 Then
                        With NPC(A)
                            St1 = Chr$(87) + DoubleChar$(A) + Chr$(.flags)
                            St1 = St1 + .JoinText + vbNullChar + .LeaveText + vbNullChar + .SayText(0) + vbNullChar + .SayText(1) + vbNullChar + .SayText(2) + vbNullChar + .SayText(3) + vbNullChar + .SayText(4)
                            SendSocket Index, St1
                        End With
                    End If
                Else
                    Hacker Index, "A.71"
                End If

            Case 51    'Upload NPC Data
                If Len(St) >= 109 And .Access >= 2 Then
                    A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                    If A >= 1 Then
                        PrintGod .User, " (Save NPC Data) NPC: " + CStr(A)
                        With NPC(A)
                            .flags = Asc(Mid$(St, 3, 1))
                            For B = 0 To 9
                                With .SaleItem(B)
                                    .GiveObject = GetInt(Mid$(St, 4 + B * 12, 2))
                                    .GiveValue = Asc(Mid$(St, 6 + B * 12, 1)) * 16777216 + Asc(Mid$(St, 7 + B * 12, 1)) * 65536 + Asc(Mid$(St, 8 + B * 12, 1)) * 256& + Asc(Mid$(St, 9 + B * 12, 1))
                                    .TakeObject = GetInt(Mid$(St, 10 + B * 12, 2))
                                    .TakeValue = Asc(Mid$(St, 12 + B * 12, 1)) * 16777216 + Asc(Mid$(St, 13 + B * 12, 1)) * 65536 + Asc(Mid$(St, 14 + B * 12, 1)) * 256& + Asc(Mid$(St, 15 + B * 12, 1))
                                End With
                            Next B
                            '124
                            GetSections Mid$(St, 124)
                            .Name = Word(1)
                            .JoinText = Word(2)
                            .LeaveText = Word(3)
                            .SayText(0) = Word(4)
                            .SayText(1) = Word(5)
                            .SayText(2) = Word(6)
                            .SayText(3) = Word(7)
                            .SayText(4) = Word(8)
                            NPCRS.Seek "=", A
                            If NPCRS.NoMatch = True Then
                                NPCRS.AddNew
                                NPCRS!number = A
                            Else
                                NPCRS.Edit
                            End If
                            NPCRS!Name = .Name
                            NPCRS!flags = .flags
                            NPCRS!JoinText = .JoinText
                            NPCRS!LeaveText = .LeaveText
                            NPCRS!SayText0 = .SayText(0)
                            NPCRS!SayText1 = .SayText(1)
                            NPCRS!SayText2 = .SayText(2)
                            NPCRS!SayText3 = .SayText(3)
                            NPCRS!SayText4 = .SayText(4)
                            St2 = vbNullString
                            For B = 0 To 9
                                With .SaleItem(B)
                                    NPCRS("GiveObject" + CStr(B)) = .GiveObject
                                    NPCRS("GiveValue" + CStr(B)) = .GiveValue
                                    NPCRS("TakeObject" + CStr(B)) = .TakeObject
                                    NPCRS("TakeValue" + CStr(B)) = .TakeValue
                                    St2 = St2 + DoubleChar$(.GiveObject) + QuadChar$(.GiveValue) + DoubleChar$(.TakeObject) + QuadChar$(.TakeValue)
                                End With
                            Next B
                            If Not .Version = 255 Then .Version = .Version + 1 Else .Version = 1
                            NPCRS!Version = .Version
                            NPCRS.Update
                            SendAll Chr$(85) + DoubleChar$(A) + Chr$(.Version) + Chr$(.flags) + St2 + .Name
                            GenerateNPCVersionList
                        End With
                    End If
                Else
                    Hacker Index, "A.72"
                End If

            Case 53    'trade item(s)
                If Len(St) = 1 Then
                    A = Map(MapNum).NPC
                    If A >= 1 Then
                        B = Asc(Mid$(St, 1, 1))
                        If B <= 9 Then
                        
                            With NPC(A).SaleItem(B)
                                C = .GiveObject
                                D = .GiveValue
                                E = .TakeObject
                                F = .TakeValue
                            End With
                            
                            If C >= 1 And E >= 1 Then
                                G = FindUnEquipInvObject(Index, E)
                                If G > 0 Then
                                    If Object(E).Type = 6 Or Object(E).Type = 11 Then
                                        If .Inv(G).Value >= F Then
                                            H = 1
                                        Else
                                            H = 0
                                        End If
                                    Else
                                        H = 1
                                    End If
                                    If H = 1 Then
                                        If Object(C).Type = 6 Or Object(C).Type = 11 Then
                                            I = FindInvObject(Index, C)
                                            If I = 0 Then
                                                I = FreeInvNum(Index)
                                                If I > 0 Then
                                                    .Inv(I).Value = 0
                                                End If
                                            End If
                                        Else
                                            I = FreeInvNum(Index)
                                        End If
                                        If I > 0 Then
                                            Parameter(0) = Index
                                            Parameter(1) = C
                                            Parameter(2) = D
                                            GetObjResult = RunScript("GetObj")
                                            
                                            Parameter(1) = E
                                            Parameter(2) = F
                                            DropObjResult = RunScript("DropObj")
                                                                                        
                                            If GetObjResult = 0 And DropObjResult = 0 Then
                                                With .Inv(G)
                                                    If Object(E).Type = 6 Or Object(E).Type = 11 Then
                                                        .Value = .Value - F
                                                        If .Value = 0 Then
                                                            .Object = 0
                                                            .ItemPrefix = 0
                                                            .ItemSuffix = 0
                                                        End If
                                                    Else
                                                        .Object = 0
                                                        .ItemPrefix = 0
                                                        .ItemSuffix = 0
                                                    End If
                                                End With
                                                CalculateStats Index
                                                With .Inv(I)
                                                    .Object = C
                                                    Select Case Object(C).Type
                                                    Case 1, 2, 3, 4    'Weapon, Shield, Armor, Helmut
                                                        .Value = CLng(Object(C).Data(0)) * 10
                                                    Case 6, 11    'Money
                                                        If CDbl(.Value) + CDbl(D) >= 2147483647# Then
                                                            .Value = 2147483647
                                                        Else
                                                            .Value = .Value + D
                                                            If .Value = 0 Then .Value = 1
                                                        End If
                                                    Case 8    'Ring
                                                        .Value = CLng(Object(C).Data(1)) * 10
                                                    Case Else
                                                        .Value = 0
                                                    End Select
                                                End With
                                                SendRaw Index, DoubleChar(10) + Chr$(17) + Chr$(G) + DoubleChar$(CLng(.Inv(G).Object)) + QuadChar(.Inv(G).Value) + Chr$(.Inv(G).ItemPrefix) + Chr$(.Inv(G).ItemSuffix) + DoubleChar(10) + Chr$(17) + Chr$(I) + DoubleChar$(CLng(.Inv(I).Object)) + QuadChar(.Inv(I).Value) + Chr$(.Inv(I).ItemPrefix) + Chr$(.Inv(I).ItemSuffix)    'Change inv objects
                                            End If
                                        Else
                                            SendSocket Index, Chr$(16) + Chr$(1)    'Inventory Full
                                        End If
                                    Else
                                        SendSocket Index, Chr$(16) + Chr$(27)    'Can't afford that
                                    End If
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(27)    'Can't afford that
                                End If
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "A.74"
                End If
            Case 54    'Bank Deposit
                If Len(St) = 4 Then
                    A = Map(MapNum).NPC
                    If A >= 1 Then
                        If ExamineBit(NPC(A).flags, 0) = True Then
                            B = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            If B > 0 Then
                                C = FindInvObject(Index, CLng(World.ObjMoney))
                                If C > 0 Then
                                    If .Inv(C).Value >= B And B > 0 Then
                                        With .Inv(C)
                                            .Value = .Value - B
                                            If .Value = 0 Then
                                                .Object = 0
                                                .ItemPrefix = 0
                                                .ItemSuffix = 0
                                            End If
                                        End With
                                        If CDbl(.Bank) + CDbl(B) >= 2147483647# Then
                                            .Bank = 2147483647
                                        Else
                                            .Bank = .Bank + B
                                        End If
                                        SendRaw Index, DoubleChar(10) + Chr$(17) + Chr$(C) + Chr$(.Inv(C).Object) + QuadChar(.Inv(C).Value) + Chr$(.Inv(C).ItemPrefix) + Chr$(.Inv(C).ItemSuffix) + DoubleChar(5) + Chr$(89) + QuadChar(.Bank)    'Change inv object / Bank Balance
                                    Else
                                        SendSocket Index, Chr$(16) + Chr$(15)    'Not enough money
                                    End If
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(15)    'Not enough money
                                End If
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(28)    'Not in a bank
                        End If
                    End If
                Else
                    Hacker Index, "A.75"
                End If

            Case 55    'Bank Data
                If Len(St) > 1 Then
                    A = Map(MapNum).NPC
                    If A >= 1 Then
                        If ExamineBit(NPC(A).flags, 0) = True Then
                            ProcessBankData Index, St
                        Else
                            SendSocket Index, Chr$(16) + Chr$(28)    'Not in a bank
                        End If
                    End If
                Else
                    Hacker Index, "Bank Deposit"
                End If
            Case 56    'Send Bank
                If Len(St) = 0 Then
                    A = Map(MapNum).NPC
                    If A >= 1 Then
                        If ExamineBit(NPC(A).flags, 0) = True Then
                            SendBankData Index
                        Else
                            SendSocket Index, Chr$(16) + Chr$(28)    'Not in a bank
                        End If
                    End If
                Else
                    Hacker Index, "Bank View"
                End If

            Case 57    'Edit Ban
                If Len(St) = 1 And .Access >= 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 And A <= 255 Then
                        With Ban(A)
                            If Not .Banner = "SuperAdmin" Or Player(Index).Access = 4 Then
                                B = .UnbanDate - CLng(Date)
                                If B < 0 Then B = 0
                                If B > 255 Then B = 255
                                SendSocket Index, Chr$(92) + Chr$(A) + Chr$(B) + .Name + vbNullChar + .Banner + vbNullChar + .Reason + vbNullChar + .ComputerID + vbNullChar + .IPAddress
                            End If
                        End With
                    End If
                Else
                    Hacker Index, "A.80"
                End If

            Case 58    'Change Ban
                If Len(St) >= 4 And .Access >= 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 And A <= 50 Then
                        With Ban(A)
                            GetSections Mid$(St, 3)
                            .UnbanDate = CLng(Date) + Asc(Mid$(St, 2, 1))
                            .Name = Word(1)
                            .Banner = Word(2)
                            .Reason = Word(3)
                            .ComputerID = Word(4)
                            .IPAddress = Word(5)
                            BanRS.Seek "=", A
                            If BanRS.NoMatch Then
                                BanRS.AddNew
                                BanRS!number = A
                            Else
                                BanRS.Edit
                            End If
                            BanRS!Name = .Name
                            BanRS!Reason = .Reason
                            BanRS!UnbanDate = .UnbanDate
                            BanRS!Banner = .Banner
                            BanRS!ComputerID = .ComputerID
                            BanRS!IPAddress = .IPAddress
                            BanRS.Update
                        End With
                    End If
                Else
                    Hacker Index, "A.81"
                End If

            Case 59    'Edit Script
                If Len(St) >= 1 And .Access >= 2 Then
                    PrintGod Player(Index).User, " (Edit Script) " + St
                    ScriptRS.Seek "=", St
                    If ScriptRS.NoMatch = False Then
                        SendSocket Index, Chr$(94) + St + vbNullChar + ScriptRS!Source
                    Else
                        SendSocket Index, Chr$(94) + St + vbNullChar
                    End If
                Else
                    Hacker Index, "A.87"
                End If

            Case 60    'Change Script
                If Len(St) >= 3 And .Access >= 2 Then
                    A = InStr(St, vbNullChar)
                    If A >= 2 Then
                        B = InStr(A + 1, St, vbNullChar)
                        If B > 0 Then
                            PrintGod Player(Index).User, " (Change Script) " + Left$(St, A - 1)
                            ScriptRS.Seek "=", Left$(St, A - 1)
                            St1 = Mid$(St, A + 1, B - A - 1)
                            St2 = Mid$(St, B + 1)
                            If St1 = "" And St2 = "" Then
                                If ScriptRS.NoMatch = False Then
                                    ScriptRS.Delete
                                End If
                            Else
                                If ScriptRS.NoMatch Then
                                    ScriptRS.AddNew
                                    ScriptRS!Name = Left$(St, A - 1)
                                Else
                                    ScriptRS.Edit
                                End If
                                ScriptRS!Source = St1
                                PrintGodSilent .User, St1
                                ScriptRS!Data = St2
                                ScriptRS.Update
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "A.88"
                End If

            Case 61    'Request Durability
                If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A > 0 And A <= 20 Then
                        If .Inv(A).Object > 0 Then
                            B = GetObjectDur(Index, A)
                            SendSocket Index, Chr$(100) + Chr$(B)
                        End If
                    Else
                        Hacker Index, "Durability: Invalid Slot"
                    End If
                Else
                    Hacker Index, "Durability Request"
                End If

            Case 62    'Command
                If Len(St) >= 1 Then
                    GetSections St
                    A = SysAllocStringByteLen(Word(1), Len(Word(1)))
                    B = SysAllocStringByteLen(Word(2), Len(Word(2)))
                    C = SysAllocStringByteLen(Word(3), Len(Word(3)))
                    D = SysAllocStringByteLen(Word(4), Len(Word(4)))
                    If .Access > 0 Then
                        PrintGod .User, " (Command) " + Word(1) + " : " + Word(2) + " : " + Word(3) + " : " + Word(4)
                    End If
                    Parameter(0) = Index
                    Parameter(1) = A
                    Parameter(2) = B
                    Parameter(3) = C
                    Parameter(4) = D
                    E = RunScript("COMMAND")
                    SysFreeString D
                    SysFreeString C
                    SysFreeString B
                    SysFreeString A
                    If E = 0 Then
                        SendSocket Index, Chr$(56) + Chr$(14) + "Invalid command."
                    End If
                End If

            Case 65    'Repairing
                If Len(St) >= 1 Then
                    Select Case Asc(Mid$(St, 1, 1))
                    Case 1    'NPC Repair Display
                        If ExamineBit(NPC(Map(MapNum).NPC).flags, 1) = True Then
                            SendSocket Index, Chr$(98) + Chr$(1)
                        Else
                            SendSocket Index, Chr$(16) + Chr$(32)
                        End If
                    Case 2    'NPC Repair the Object
                        .CurrentRepairTar = Asc(Mid$(St, 2, 1))
                        If ExamineBit(NPC(Map(MapNum).NPC).flags, 1) = True Then
                            If .CurrentRepairTar <= 20 Then
                                If Not ExamineBit(Object(.Inv(.CurrentRepairTar).Object).flags, 0) = 255 Then
                                    RepairItem Index
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(34)
                                End If
                            Else
                                If Not ExamineBit(Object(.EquippedObject(.CurrentRepairTar - 20).Object).flags, 0) = 255 Then
                                    RepairItem Index
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(34)
                                End If
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(34)
                        End If
                    Case 3    'Repair all Items
                        RepairAll Index
                    End Select
                Else
                    Hacker Index, "R.1"
                End If

            Case 70    'Inform
                If Len(St) > 0 Then
                    BootPlayer Index, 0, St
                    PrintCheat .Name & "  " & .IP & "  " & St
                End If

            Case 72    'Projectile Shot
                If .LastMsg > .ShootTimer Then
                    .ShootTimer = .LastMsg + 850
                    If .EquippedObject(6).Object > 0 Then
                        If .EquippedObject(1).Object > 0 Then
                            If .Inv(.EquippedObject(6).Object).Object > 0 Then
                                SendToMap MapNum, Chr$(99) + Chr$(5) + Chr$(.D) + Chr$(.X) + Chr$(.Y) + Chr$(Object(.Inv(.EquippedObject(6).Object).Object).Data(2)) + Chr$(Index)
                                If .Inv(.EquippedObject(6).Object).Value = 1 Then
                                    TakeObj Index, .Inv(.EquippedObject(6).Object).Object, 1
                                    .EquippedObject(6).Object = 0
                                    .EquippedObject(6).ItemPrefix = 0
                                    .EquippedObject(6).ItemSuffix = 0
                                Else
                                    If .Inv(.EquippedObject(6).Object).Value > 0 Then
                                        .Inv(.EquippedObject(6).Object).Value = .Inv(.EquippedObject(6).Object).Value - 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
            Case 73    'Projectile Hit
                If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    ProjectileAttackMonster Index, A
                End If
            Case 74    'Hit Player
                If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    ProjectileAttackPlayer Index, A
                End If
            Case 78    'Skill Data
                If Len(St) > 0 Then
                    ProcessSkillData Index, St, Tick
                End If
            Case 79    'Script Projectile Attack
                If Len(St) = 5 Then
                    A = Asc(Mid$(St, 1, 1))
                    If (CheckSum(Mid$(St, 2, 4)) Mod 256) = A Then
                        ProcessScriptProjectile Index, Mid$(St, 3)
                    Else
                        BanPlayer Index, 0, 100, "ScriptProjectilePacketChecksum", "SuperAdmin"
                    End If
                Else
                    BanPlayer Index, 0, 100, "ScriptProjectilePacketLength", "SuperAdmin"
                End If
            Case 81    'Choose Guild Sprite
                If Len(St) = 2 Then
                    A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                    If .Guild > 0 And .GuildRank >= 2 Then
                        If Not GuildRS.BOF Then
                            GuildRS.MoveFirst
                            While Not GuildRS.EOF
                                If A = GuildRS!Sprite Then A = 0
                                GuildRS.MoveNext
                            Wend
                        End If
                        If Not A = 0 Then
                            If Guild(.Guild).Bank >= 100000 Then
                                With Guild(.Guild)
                                    .Bank = .Bank - 100000
                                    SendToGuild CLng(Player(Index).Guild), Chr$(152) + QuadChar(.Bank) + QuadChar$(GetGuildUpkeep(CLng(Player(Index).Guild)))
                                    SendToGuild CLng(Player(Index).Guild), Chr$(97) + Chr$(6)
                                    If .Name <> "" Then
                                        .Sprite = A
                                        GuildRS.Bookmark = .Bookmark
                                        GuildRS.Edit
                                        GuildRS!Sprite = A
                                        GuildRS.Update

                                        For C = 0 To 19
                                            With .Member(C)
                                                If .Name <> "" Then
                                                    D = FindPlayer(.Name)
                                                    If D > 0 Then
                                                        With Player(D)
                                                            If A > 0 Then
                                                                .Sprite = A
                                                            Else
                                                                .Sprite = .Class * 2 + .Gender - 1
                                                            End If
                                                            SendToMap .Map, Chr$(63) + Chr$(D) + DoubleChar$(CLng(.Sprite))
                                                        End With
                                                    Else
                                                        UserRS.Seek "=", .Name
                                                        If UserRS.NoMatch = False Then
                                                            If A > 0 Then
                                                                D = A
                                                            Else
                                                                D = UserRS!Class * 2 + UserRS!Gender - 1
                                                            End If
                                                            If D >= 1 And D <= MaxSprite Then
                                                                UserRS.Edit
                                                                UserRS!Sprite = D
                                                                UserRS.Update
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End With
                                        Next C
                                    End If
                                End With
                            Else
                                SendSocket Index, Chr$(97) + Chr$(5)
                            End If
                        Else
                            SendSocket Index, Chr$(97) + Chr$(7)
                        End If
                    End If
                End If
            Case 82    'Edit Magic
                If Len(St) = 2 And .Access > 1 Then
                    A = GetInt(Mid$(St, 1, 2))
                    If A >= 1 Then
                        With Magic(A)
                            PrintGod Player(Index).User, " (Edit Magic) " + .Name
                            SendSocket Index, Chr$(128) + DoubleChar$(A)
                        End With
                    End If
                Else
                    Hacker Index, "Edit Magic"
                End If
            Case 83    'Save Magic
                If Len(St) > 2 And .Access > 1 Then
                    A = GetInt(Mid$(St, 1, 2))
                    With Magic(A)
                        PrintGod Player(Index).User, " (Save Magic) " + .Name
                        .Level = Asc(Mid$(St, 3, 1))
                        .Class = Asc(Mid$(St, 4, 1))
                        .Icon = Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1))
                        .IconType = Asc(Mid$(St, 7, 1))
                        .CastTimer = Asc(Mid$(St, 8, 1)) * 256 + Asc(Mid$(St, 9, 1))
                        B = InStr(10, St, vbNullChar)
                        If B > 10 And B < Len(St) Then
                            .Name = Mid$(St, 10, B - 10)
                            .Description = Mid$(St, B + 1)

                            MagicRS.Seek "=", A
                            If MagicRS.NoMatch = True Then
                                MagicRS.AddNew
                                MagicRS!number = A
                            Else
                                MagicRS.Edit
                            End If
                            MagicRS!Name = .Name
                            MagicRS!Level = .Level
                            MagicRS!Class = .Class
                            MagicRS!Icon = .Icon
                            MagicRS!IconType = .IconType
                            MagicRS!CastTimer = .CastTimer
                            MagicRS!Description = .Description
                            If Not .Version = 255 Then .Version = .Version + 1 Else .Version = 1
                            MagicRS!Version = .Version
                            MagicRS.Update

                            SendAll Chr$(127) + DoubleChar$(A) + Chr$(.Class) + Chr$(.Level) + Chr$(.Version) + DoubleChar$(CLng(.Icon)) + Chr$(.IconType) + DoubleChar$(CLng(.CastTimer)) + .Name + vbNullChar + .Description
                            GenerateMagicVersionList
                        End If
                    End With
                Else
                    Hacker Index, "Save Magic"
                End If
            Case 84    'Use Magic
                If Len(St) = 1 Then
                    If .LastMsg > .CastTimer Or .Access > 0 Then
                        If .IsDead = False Then
                            .CastTimer = .LastMsg + 450
                            A = Asc(Mid$(St, 1, 1))
                            If ExamineBit(Magic(A).Class, .Class - 1) = True And .Level >= Magic(A).Level Then
                                Parameter(0) = Index
                                Parameter(1) = A
                                RunScript "Spell"
                            End If
                        End If
                    End If
                End If
            Case 86    'Edit Prefix
                If Len(St) = 1 Then
                    If .Access > 0 Then
                        A = Asc(Mid$(St, 1, 1))
                        SendSocket Index, Chr$(135) + Chr$(A)
                    Else
                        Hacker Index, "Edit Prefix No Access"
                    End If
                Else
                    Hacker Index, "Edit Prefix"
                End If
            Case 87    'Save Prefix
                If Len(St) >= 4 Then
                    If .Access > 0 Then
                        A = Asc(Mid$(St, 1, 1))
                        With ItemPrefix(A)
                            .ModificationType = Asc(Mid$(St, 2, 1))
                            .ModificationValue = Asc(Mid$(St, 3, 1))
                            .OccursNaturally = Asc(Mid$(St, 4, 1))
                            If Len(St) > 4 Then
                                .Name = Mid$(St, 5)
                            Else
                                .Name = ""
                            End If

                            PrefixRS.Seek "=", A
                            If PrefixRS.NoMatch = True Then
                                PrefixRS.AddNew
                                PrefixRS!number = A
                            Else
                                PrefixRS.Edit
                            End If
                            PrefixRS!Name = .Name
                            PrefixRS!ModificationType = .ModificationType
                            PrefixRS!ModificationValue = .ModificationValue
                            PrefixRS!OccursNaturally = .OccursNaturally
                            If Not .Version = 255 Then .Version = .Version + 1 Else .Version = 1
                            PrefixRS!Version = .Version
                            PrefixRS.Update
                            SendAll Chr$(133) + Chr$(A) + Chr$(.ModificationType) + Chr$(.ModificationValue) + Chr$(.OccursNaturally) + Chr$(.Version) + .Name
                            GeneratePrefixVersionList
                        End With
                    Else
                        Hacker Index, "Edit Prefix No Access"
                    End If
                Else
                    Hacker Index, "Edit Prefix"
                End If
            Case 88    'Edit Suffix
                If Len(St) = 1 Then
                    If .Access > 0 Then
                        A = Asc(Mid$(St, 1, 1))
                        SendSocket Index, Chr$(136) + Chr$(A)
                    Else
                        Hacker Index, "Edit Suffix No Access"
                    End If
                Else
                    Hacker Index, "Edit Suffix"
                End If
            Case 89    'Save Suffix
                If Len(St) >= 4 Then
                    If .Access > 0 Then
                        A = Asc(Mid$(St, 1, 1))
                        With ItemSuffix(A)
                            .ModificationType = Asc(Mid$(St, 2, 1))
                            .ModificationValue = Asc(Mid$(St, 3, 1))
                            .OccursNaturally = Asc(Mid$(St, 4, 1))
                            If Len(St) > 4 Then
                                .Name = Mid$(St, 5)
                            Else
                                .Name = ""
                            End If

                            SuffixRS.Seek "=", A
                            If SuffixRS.NoMatch = True Then
                                SuffixRS.AddNew
                                SuffixRS!number = A
                            Else
                                SuffixRS.Edit
                            End If
                            SuffixRS!Name = .Name
                            SuffixRS!ModificationType = .ModificationType
                            SuffixRS!ModificationValue = .ModificationValue
                            SuffixRS!OccursNaturally = .OccursNaturally
                            If Not .Version = 255 Then .Version = .Version + 1 Else .Version = 1
                            SuffixRS!Version = .Version
                            SuffixRS.Update
                            SendAll Chr$(134) + Chr$(A) + Chr$(.ModificationType) + Chr$(.ModificationValue) + Chr$(.OccursNaturally) + Chr$(.Version) + .Name
                            GenerateSuffixVersionList
                        End With
                    Else
                        Hacker Index, "Save Suffix No Access"
                    End If
                Else
                    Hacker Index, "Save Suffix"
                End If
            Case 90    'Forward
                If Len(St) >= 2 Then
                    A = Asc(Mid$(St, 1, 1))
                    SendSocket A, Mid$(St, 2)
                End If
            Case 91    'Click Tile
                If Len(St) >= 2 Then
                    A = Asc(Mid$(St, 1, 1))
                    B = Asc(Mid$(St, 2, 1))
                    Parameter(0) = Index
                    Parameter(1) = .Map
                    Parameter(2) = A
                    Parameter(3) = B
                    RunScript ("MAPCLICK")
                End If
            Case 92    'Speedhack Check
                A = Tick - .SpeedHackTimer
                .SpeedHackTimer = Tick
                If A < 90000 Then
                    PrintCheat "Speed hack timer " + .Name + " " + CStr(A) + " ms out of 120000"
                    .SpeedStrikes = .SpeedStrikes + 1
                    If .SpeedStrikes > 5 Then
                        Hacker Index, "Speedhack Detected"
                    End If
                End If
            Case 93    'Change Guild MOTD
                If .Guild > 0 And .GuildRank >= 2 And Len(St) > 1 Then
                    Guild(.Guild).MOTD = Mid$(St, 1)
                    Guild(.Guild).MOTDDate = CLng(Date)
                    Guild(.Guild).MOTDCreator = .Name
                    GuildRS.Bookmark = Guild(.Guild).Bookmark
                    GuildRS.Edit
                    GuildRS!MOTD = Guild(.Guild).MOTD
                    GuildRS!MOTDDate = Guild(.Guild).MOTDDate
                    GuildRS!MOTDCreator = Guild(.Guild).MOTDCreator
                    GuildRS.Update
                    SendToGuild CLng(.Guild), Chr$(16) + vbNullChar + Chr$(15) + Guild(.Guild).MOTDCreator + " - " + CStr(Guild(.Guild).MOTDDate) + " - " + Guild(.Guild).MOTD
                End If
            Case 94    'View Guild MOTD
                If .Guild > 0 Then
                    If Not Guild(.Guild).MOTD = "" Then
                        SendSocket Index, Chr$(56) + Chr$(14) + .Name + " - " + Guild(.Guild).MOTDCreator + " - " + CStr(CDate(Guild(.Guild).MOTDDate)) + " - " + Guild(.Guild).MOTD
                    Else
                        SendSocket Index, Chr$(16) + Chr$(43)
                    End If
                End If
            Case 95    'Guild Sprite Command
                If .Guild > 0 Then
                    SetPlayerSprite Index, 0
                End If
            Case 96    'Ping
                SendSocket Index, Chr$(149)
            Case 97    'Sell Item
                If Len(St) = 2 Then
                    A = Asc(Mid$(St, 1, 1))    'Action
                    B = Asc(Mid$(St, 2, 1))    'Inventory Slot
                    ProcessSellItem Index, A, B
                Else
                    Hacker Index, "Sell Item"
                End If
            Case 98    'Cheating report
                If Len(St) > 0 Then
                    If Len(St) < 255 Then
                        'BanPlayer Index, 0, 2, St, "Server"
                    Else
                        'BanPlayer Index, 0, 2, "Cheating", "Server"
                    End If
                    PrintCheat .Name & "  " & .IP & "  " & St
                End If
            Case 99    'Cheat Report
                If Len(St) > 0 Then
                    PrintCheat .Name & "  " & .IP & "  " & St
                End If
            Case 100    'Debug Report
                If Len(St) > 0 Then
                    PrintDebug "(Player Submitted) - " + .Name + "  " + .IP + "  " & St
                End If
            Case 101    'Choose Guild Sprite
                If Len(St) = 0 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        With Guild(.Guild)
                            If .Name <> "" Then
                                .Sprite = 0
                                GuildRS.Bookmark = .Bookmark
                                GuildRS.Edit
                                GuildRS!Sprite = 0
                                GuildRS.Update

                                For C = 0 To 19
                                    With .Member(C)
                                        If .Name <> "" Then
                                            D = FindPlayer(.Name)
                                            If D > 0 Then
                                                With Player(D)
                                                    .Sprite = .Class * 2 + .Gender - 1
                                                    SendToMap .Map, Chr$(63) + Chr$(D) + DoubleChar$(CLng(.Sprite))
                                                End With
                                            Else
                                                UserRS.Seek "=", .Name
                                                If UserRS.NoMatch = False Then
                                                    D = UserRS!Class * 2 + UserRS!Gender - 1
                                                    If D >= 1 And D <= MaxSprite Then
                                                        UserRS.Edit
                                                        UserRS!Sprite = D
                                                        UserRS.Update
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End With
                                Next C
                            End If
                        End With
                    End If
                End If
            Case 102    'Reset Guild Stats
                If Len(St) = 0 Then
                    If .Guild > 0 And .GuildRank >= 3 Then
                        With Guild(.Guild)
                            If .Name <> "" Then
                                .Kills = 0
                                .Deaths = 0
                                .UpdateFlag = True
                            End If
                        End With
                    End If
                End If
            Case 103 'Update e-mail address
                If Len(St) > 2 And Len(St) < 102 Then
                    If Len(.Email) < 2 Then
                        .Email = Mid$(St, 1)
                        'E-mail has been updated
                         SendSocket Index, Chr$(16) + Chr$(46)
                    Else
                        'E-mail is already saved
                        SendSocket Index, Chr$(16) + Chr$(45)
                    End If
                End If
            Case 104 'Send uncompressed map
                If Len(St) = 0 Then
                    MapRS.Seek "=", .Map
                    If MapRS.NoMatch Then
                        SendSocket Index, Chr$(21) + String$(2677, vbNullChar)
                    Else
                        St = UncompressString(MapRS!Data)
                        SendSocket Index, Chr$(21) + St
                    End If
                Else
                    Hacker Index, "A.66"
                End If

            Case Else
                Hacker Index, "B.3"
            End Select
        End Select
    End With

    Exit Sub

LogDatShit:
    SendToGods Chr$(16) + Chr$(0) + "WARNING:  Server Crashed on Packet # " + CStr(PacketID) & " from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    PrintLog "WARNING:  Server Crashed on Packet # " + CStr(PacketID) & " from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    PrintDebug "WARNING:  Server Crashed on Packet # " + CStr(PacketID) & " from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    BootPlayer Index, 0, "Crashed Server"
End Sub

Sub ProcessRawData(Index As Long, SocketData As String)
    On Error GoTo LogDatShit

    Dim St As String, PacketLength As Long, PacketID As Long
LoopRead:
    With Player(Index)
        If Len(SocketData) >= 3 Then
            PacketLength = GetInt(Mid$(SocketData, 1, 2))
            If PacketLength >= 5072 And .Access = 0 Then
                Hacker Index, "C.1.1"
                Exit Sub
            End If
            If Len(SocketData) - 2 >= PacketLength Then
                St = Mid$(SocketData, 3, PacketLength)
                SocketData = Mid$(SocketData, PacketLength + 3)
                If PacketLength > 0 Then
                    PacketID = Asc(Mid$(St, 1, 1))
                    If Len(St) > 1 Then
                        St = Mid$(St, 2)
                    Else
                        St = ""
                    End If
                    ProcessString Index, PacketID, St
                End If
                GoTo LoopRead
            End If
        End If
    End With

    Exit Sub

LogDatShit:
    SendToGods Chr$(16) & vbNullChar & "WARNING:  Server Crashed in ProcessRawData from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    PrintLog "WARNING:  Server Crashed in ProcessRawData from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    PrintDebug "WARNING:  Server Crashed in ProcessRawData from " & Player(Index).Name & "  DATA:  " & St & "  " & Err.Description
    BootPlayer Index, 0, "Crashed Server"
End Sub

Sub ProcessStartNewGuild(Index As Long, St As String)
    Dim A As Long, B As Long, C As Long
    With Player(Index)
        If Len(St) >= 1 And Len(St) <= 15 Then
            If ValidName(St) = True Then
                If .Guild = 0 Then
                    C = FindInvObject(Index, CLng(World.ObjMoney))    'Money Slot
                    If C > 0 Then    'Has money
                        If .Inv(C).Value >= World.GuildNewPrice Then    'Has the Cash
                            A = FreeGuildNum
                            If A > 0 Then
                                UserRS.Index = "Name"
                                UserRS.Seek "=", St
                                If UserRS.NoMatch = True And GuildNum(St) = 0 And NPCNum(St) = 0 Then
                                    TakeObj Index, CLng(World.ObjMoney), World.GuildNewPrice    'Take Cash
                                    GuildRS.AddNew
                                    GuildRS!number = A
                                    With Guild(A)
                                        .Name = St
                                        GuildRS!Name = St
                                        .Bank = 0
                                        GuildRS!Bank = 0
                                        .DueDate = 0
                                        GuildRS!DueDate = 0
                                        .Hall = 0
                                        GuildRS!Hall = 0
                                        .Sprite = 0
                                        GuildRS!Sprite = 0
                                        .CreationDate = CLng(Date)
                                        GuildRS!CreationDate = CLng(Date)
                                        .Kills = 0
                                        GuildRS!Kills = 0
                                        .Deaths = 0
                                        GuildRS!Deaths = 0
                                        .MOTD = ""
                                        GuildRS!MOTD = ""
                                        .MOTDCreator = ""
                                        GuildRS!MOTDCreator = ""
                                        .MOTDDate = 0
                                        GuildRS!MOTDDate = 0
                                        For B = 0 To DeclarationCount
                                            .Declaration(B).Guild = 0
                                            .Declaration(B).Type = 0
                                            .Declaration(B).Date = 0
                                            .Declaration(B).Kills = 0
                                            .Declaration(B).Deaths = 0
                                            GuildRS("DeclarationGuild" + CStr(B)) = 0
                                            GuildRS("DeclarationType" + CStr(B)) = 0
                                            GuildRS("DeclarationDate" + CStr(B)) = 0
                                            GuildRS("DeclarationKills" + CStr(B)) = 0
                                            GuildRS("DeclarationDeaths" + CStr(B)) = 0
                                        Next B
                                        .Member(0).Name = Player(Index).Name
                                        .Member(0).Rank = 3
                                        .Member(0).JoinDate = CLng(Date)
                                        .Member(0).Kills = 0
                                        .Member(0).Deaths = 0
                                        GuildRS!MemberName0 = Player(Index).Name
                                        GuildRS!MemberRank0 = 3
                                        GuildRS!MemberJoinDate0 = CLng(Date)
                                        GuildRS!MemberKills0 = 0
                                        GuildRS!MemberDeaths0 = 0
                                        For B = 1 To 19
                                            .Member(B).Name = ""
                                            .Member(B).Rank = 0
                                            .Member(B).JoinDate = 0
                                            .Member(B).Kills = 0
                                            .Member(B).Deaths = 0
                                            .MemberCount = CountGuildMembers(B)
                                            GuildRS("MemberName" + CStr(B)) = ""
                                            GuildRS("MemberRank" + CStr(B)) = 0
                                            GuildRS("MemberJoinDate" + CStr(B)) = 0
                                            GuildRS("MemberKills" + CStr(B)) = 0
                                            GuildRS("MemberDeaths" + CStr(B)) = 0
                                        Next B
                                        GuildRS.Update
                                        GuildRS.Seek "=", A
                                        Guild(A).Bookmark = GuildRS.Bookmark
    
                                        Player(Index).Guild = A
                                        Player(Index).GuildRank = 3
    
                                        SendAll Chr$(70) + Chr$(A) + Chr$(.MemberCount) + St    'Guild Data
                                        SendSocket Index, Chr$(80) + Chr$(A)    'Guild Created
                                        SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(A)    'Player changed guild
                                    End With
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(16)    'Name in use
                                End If
                            Else
                                SendSocket Index, Chr$(16) + Chr$(18)    'Too many guilds
                            End If
                        Else
                            SendSocket Index, Chr$(16) & Chr$(15)
                        End If
                    Else
                        SendSocket Index, Chr$(16) & Chr$(15)
                    End If
                End If
            End If
        Else
            Hacker Index, "A.82"
        End If
    End With
End Sub

Sub ProcessSellItem(Index As Long, Action As Long, Slot As Long)
    Dim GoldSlot As Long, FreeSlot As Long, SellPrice As Long
    Dim A As Long
    If Index >= 1 And Index <= MaxUsers And Slot >= 1 And Slot <= 20 Then
        With Player(Index)
            If Map(.Map).NPC > 0 Then
                If ExamineBit(NPC(Map(.Map).NPC).flags, 2) = True Then
                    GoldSlot = FindInvObject(Index, CLng(World.ObjMoney))
                    FreeSlot = FreeInvNum(Index)
                    If .Inv(Slot).Object > 0 Then
                        If Object(.Inv(Slot).Object).SellPrice > 0 Then
                            SellPrice = Object(.Inv(Slot).Object).SellPrice
                            Select Case Action
                            Case 1    'Sell one
                                Select Case Object(.Inv(Slot).Object).Type
                                Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helm
                                    If Object(.Inv(Slot).Object).Data(0) * 10 > 0 Then
                                        A = (.Inv(Slot).Value / (Object(.Inv(Slot).Object).Data(0) * 10)) * Object(.Inv(Slot).Object).SellPrice
                                        SendSocket Index, Chr$(18) + Chr$(Slot)
                                        .Inv(Slot).Object = 0
                                        .Inv(Slot).Value = 0
                                        .Inv(Slot).ItemPrefix = 0
                                        .Inv(Slot).ItemSuffix = 0
                                        If A > 0 Then GiveObj Index, World.ObjMoney, A
                                    End If
                                Case 8 'Ring
                                    If Object(.Inv(Slot).Object).Data(1) * 10 > 0 Then
                                        A = (.Inv(Slot).Value / (Object(.Inv(Slot).Object).Data(1) * 10)) * Object(.Inv(Slot).Object).SellPrice
                                        SendSocket Index, Chr$(18) + Chr$(Slot)
                                        .Inv(Slot).Object = 0
                                        .Inv(Slot).Value = 0
                                        .Inv(Slot).ItemPrefix = 0
                                        .Inv(Slot).ItemSuffix = 0
                                        If A > 0 Then GiveObj Index, World.ObjMoney, A
                                    End If
                                Case 6, 11
                                    If .Inv(Slot).Value > 1 Then
                                        If GoldSlot > 0 Or FreeSlot > 0 Then
                                            .Inv(Slot).Value = .Inv(Slot).Value - 1
                                            SendSocket Index, Chr$(17) + Chr$(Slot) + DoubleChar$(CLng(.Inv(Slot).Object)) + QuadChar(.Inv(Slot).Value) + Chr$(.Inv(Slot).ItemPrefix) + Chr$(.Inv(Slot).ItemSuffix)
                                            GiveObj Index, World.ObjMoney, SellPrice
                                        Else
                                            SendSocket Index, Chr$(16) + Chr$(1)    'Inventory is full
                                        End If
                                    Else
                                        SendSocket Index, Chr$(18) + Chr$(Slot)
                                        .Inv(Slot).Object = 0
                                        .Inv(Slot).Value = 0
                                        .Inv(Slot).ItemPrefix = 0
                                        .Inv(Slot).ItemSuffix = 0
                                        GiveObj Index, World.ObjMoney, SellPrice
                                    End If
                                Case Else
                                    SendSocket Index, Chr$(18) + Chr$(Slot)
                                    .Inv(Slot).Object = 0
                                    .Inv(Slot).Value = 0
                                    .Inv(Slot).ItemPrefix = 0
                                    .Inv(Slot).ItemSuffix = 0
                                    GiveObj Index, World.ObjMoney, SellPrice
                                End Select
                            Case 2    'Sell all
                                Select Case Object(.Inv(Slot).Object).Type
                                Case 6, 11
                                    If .Inv(Slot).Value > 1 Then
                                        GiveObj Index, World.ObjMoney, SellPrice * .Inv(Slot).Value
                                        SendSocket Index, Chr$(18) + Chr$(Slot)
                                        .Inv(Slot).Object = 0
                                        .Inv(Slot).Value = 0
                                        .Inv(Slot).ItemPrefix = 0
                                        .Inv(Slot).ItemSuffix = 0
                                    Else
                                        SendSocket Index, Chr$(18) + Chr$(Slot)
                                        .Inv(Slot).Object = 0
                                        .Inv(Slot).Value = 0
                                        .Inv(Slot).ItemPrefix = 0
                                        .Inv(Slot).ItemSuffix = 0
                                        GiveObj Index, World.ObjMoney, SellPrice
                                    End If
                                End Select
                            End Select
                        End If
                    End If
                End If
            End If
        End With
    End If
End Sub

Sub ProcessVersion(Index As Long, St As String)
    Dim A As Long
    With Player(Index)
        .ClientVer = Asc(Mid$(St, 1, 1))
        .ComputerID = Mid$(St, 3)
        A = CheckSum(.ComputerID) Mod 256
        If Not A = Asc(Mid$(St, 2, 1)) Then
            SendToGods Chr$(56) + Chr$(7) + "WARNING:  Player " & .Name & " from IP " & .IP & " attempted to login with an invalid checksum."
            PrintCheat "WARNING:  Player " & .Name & " from IP " & .IP & " attempted to login with an invalid checksum."
            AddSocketQue Index
        End If
        If Not .ClientVer = CurrentClientVer Then
            SendSocket Index, Chr$(116)
        Else
            If CheckBan(Index, "NotLoggedIn", .ComputerID, .IP) = False Then
    
            End If
        End If
    End With
End Sub

Sub ProcessSaveMonster(Index As Long, St As String)
    Dim A As Long
    With Player(Index)
        A = GetInt(Mid$(St, 1, 2))
        If A >= 1 Then
            With Monster(A)
                .Sprite = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                .HP = GetInt(Mid$(St, 5, 2))
                .Strength = Asc(Mid$(St, 7, 1))
                .Armor = Asc(Mid$(St, 8, 1))
                .Speed = Asc(Mid$(St, 9, 1))
                .Sight = Asc(Mid$(St, 10, 1))
                .Agility = Asc(Mid$(St, 11, 1))
                .flags = Asc(Mid$(St, 12, 1))
                .Object(0) = GetInt(Mid$(St, 13, 2))
                .Value(0) = Asc(Mid$(St, 15, 1))
                .Object(1) = GetInt(Mid$(St, 16, 2))
                .Value(1) = Asc(Mid$(St, 18, 1))
                .Object(2) = GetInt(Mid$(St, 19, 2))
                .Value(2) = Asc(Mid$(St, 21, 1))
                .Experience = Asc(Mid$(St, 22, 1)) * 10
                .MagicDefense = Asc(Mid$(St, 23, 1))
                If Len(St) >= 24 Then
                    .Name = Mid$(St, 24)
                Else
                    .Name = ""
                End If
                MonsterRS.Seek "=", A
                If MonsterRS.NoMatch Then
                    MonsterRS.AddNew
                    MonsterRS!number = A
                Else
                    MonsterRS.Edit
                End If
                MonsterRS!Name = .Name
                MonsterRS!Sprite = .Sprite
                MonsterRS!HP = .HP
                MonsterRS!Strength = .Strength
                MonsterRS!Armor = .Armor
                MonsterRS!Speed = .Speed
                MonsterRS!Sight = .Sight
                MonsterRS!Agility = .Agility
                MonsterRS!flags = .flags
                MonsterRS!Object0 = .Object(0)
                MonsterRS!Value0 = .Value(0)
                MonsterRS!Object1 = .Object(1)
                MonsterRS!Value1 = .Value(1)
                MonsterRS!Object2 = .Object(2)
                MonsterRS!Value2 = .Value(2)
                MonsterRS!Experience = .Experience
                MonsterRS!MagicDefense = .MagicDefense
                If Not .Version = 255 Then .Version = .Version + 1 Else .Version = 1
                MonsterRS!Version = .Version
                MonsterRS.Update
                PrintGod Player(Index).User, " (Save Monster) Monster #: " + CStr(A) + ", Experience: " + CStr(.Experience) + ", Drops: " + GetMonsterDrops(A)
                SendAll Chr$(32) + DoubleChar$(A) + DoubleChar$(CLng(.Sprite)) + Chr$(.Version) + DoubleChar$(CLng(.HP)) + Chr$(.flags) + .Name
                GenerateMonsterVersionList
            End With
        End If
    End With
End Sub

Function GetMonsterDrops(TheMonster As Long) As String
    Dim A As Long
    Dim St1 As String
    
    If TheMonster > 0 Then
        With Monster(TheMonster)
            For A = 0 To 2
                If Monster(TheMonster).Object(A) > 0 Then
                    St1 = St1 + Object(.Object(A)).Name + " (" + CStr(.Value(A)) + "), "
                End If
            Next A
            
        End With
    End If
    
    GetMonsterDrops = St1
End Function

Sub ProcessDropObject(Index As Long, MapNum As Long, St As String)
    If Not Len(St) = 5 Then
        Hacker Index, "A.18"
        Exit Sub
    End If

    Dim InvSlot As Long
    Dim Obj As Long, Value As Long
    Dim MapObj As Long
    Dim Prefix As Long, Suffix As Long
    
    InvSlot = Asc(Mid$(St, 1, 1))
    If InvSlot >= 1 And InvSlot <= 20 Then
    Else
        Exit Sub
    End If
    
    With Player(Index)
        Obj = .Inv(InvSlot).Object
        If Not Obj > 0 Then
            SendSocket Index, Chr$(16) + Chr$(3)    'No such object
            Exit Sub
        End If
        
        If Object(Obj).Type = 6 Or Object(Obj).Type = 11 Then
            If Asc(Mid$(St, 2, 1)) < 120 Then    'Crash Attempt
                Value = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                If Not Value < 0 Then
                    If Not Value < .Inv(InvSlot).Value Then
                        Value = .Inv(InvSlot).Value
                        Prefix = 0
                        Suffix = 0
                    End If
                Else
                    BanPlayer Index, 0, 99, "Definite Hacking Attempt (Dupe)", "Server"
                    Exit Sub
                End If
            Else
                BanPlayer Index, 0, 99, "Definite Hacking Attempt (Dupe/Crash)", "Server"
                Exit Sub
            End If
        Else 'Type != 6 or Type != 11
            Value = .Inv(InvSlot).Value
            Prefix = .Inv(InvSlot).ItemPrefix
            Suffix = .Inv(InvSlot).ItemSuffix
        End If 'Object Type

        Parameter(0) = Index
        Parameter(1) = Obj
        Parameter(2) = Value
        If Not RunScript("DROPOBJ") = 0 Then
            Exit Sub
        End If
        
        MapObj = FreeMapObj(MapNum)
        If Not MapObj >= 0 Then
            SendSocket Index, Chr$(16) + Chr$(2)    'Map full
            Exit Sub
        End If
        
        If .Access > 0 Then
            PrintGod .User, " (Drop) Object: " + Object(Obj).Name + "  Value: " + CStr(Value)
        End If
        
        PrintItem .User + " - " + .Name + " (Drop) " + Object(Obj).Name + " (" + CStr(Value) + ") - Map: " + CStr(.Map)
        
        If .EquippedObject(6).Object = InvSlot Then 'Ammo?
            .EquippedObject(6).Object = 0
            .EquippedObject(6).ItemPrefix = 0
            .EquippedObject(6).ItemSuffix = 0
        End If
        
        If Value < .Inv(InvSlot).Value Then
            .Inv(InvSlot).Value = .Inv(InvSlot).Value - Value
            SendSocket Index, Chr$(17) + Chr$(InvSlot) + DoubleChar$(Obj) + QuadChar(.Inv(InvSlot).Value) + Chr$(.Inv(InvSlot).ItemPrefix) + Chr$(.Inv(InvSlot).ItemSuffix)    'Update inv obj
        Else
            .Inv(InvSlot).Object = 0
            .Inv(InvSlot).Value = 0
            .Inv(InvSlot).ItemPrefix = 0
            .Inv(InvSlot).ItemSuffix = 0
            SendSocket Index, Chr$(18) + Chr$(InvSlot)    'Erase Inv Obj
        End If
        
        With Map(MapNum).Object(MapObj)
            .Object = Obj
            .ItemPrefix = Prefix
            .ItemSuffix = Suffix
            .Value = Value
            .TimeStamp = Player(Index).LastMsg + Int(Rnd * 60000) - 30000
        End With
        
        Map(MapNum).Object(MapObj).X = .X
        Map(MapNum).Object(MapObj).Y = .Y
        SendToMap MapNum, Chr$(14) + Chr$(MapObj) + DoubleChar$(Obj) + Chr$(.X) + Chr$(.Y) + Chr$(Map(MapNum).Object(MapObj).ItemPrefix) + Chr$(Map(MapNum).Object(MapObj).ItemSuffix) + QuadChar$(Map(MapNum).Object(MapObj).Value)    'New Map Obj
    End With
End Sub
