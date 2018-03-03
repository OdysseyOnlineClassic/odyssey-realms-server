Attribute VB_Name = "modServer"
Option Explicit

'Game Constants
Public Const TitleString = "Odyssey Realms"
Public Const MaxUsers = 80
Public Const DownloadSite = "http://www.odysseyclassic.info"

Public Const CurrentClientVer = 4
Public Const PrefixSuffixChance = 40

'Misc Variables
Public BackupCounter As Long
Public Startup As Boolean

Public Word(1 To 50) As String
Public Prefix As String
Public Suffix As String

Public OutdoorLight As Byte

'Sockets
Public ListeningSocket As Long

Sub SaveFlags()
    Dim A As Long, St As String
    For A = 0 To 255
        If World.Flag(A) > 0 Then
            St = St + QuadChar$(World.Flag(A))
        Else
            St = St + QuadChar$(0)
        End If
    Next A
    DataRS.Edit
    DataRS!flags = St
    DataRS.Update
End Sub
Sub SaveObjects()
    Dim A As Long, B As Long, St As String
    For A = 1 To MaxMaps
        With Map(A)
            If .Keep = True Then
                For B = 0 To MaxMapObjects
                    With .Object(B)
                        If .Object > 0 Then
                            If .Value >= 0 Then
                                If Map(A).Tile(.X, .Y).Att = 5 Or Map(A).Tile(.X, .Y).Att2 = 5 Then
                                    St = St + DoubleChar(A) + Chr$(B) + Chr$(.X) + Chr$(.Y) + DoubleChar$(.Object) + QuadChar(.Value) + Chr$(.ItemPrefix) + Chr$(.ItemSuffix)
                                End If
                            End If
                        End If
                    End With
                Next B
            End If
        End With
    Next A
    DataRS.Edit
    DataRS!ObjectData = St
    DataRS.Update
End Sub

Sub SaveBugReports()
Dim A As Long, B As Long, St As String
On Error GoTo ExitSub
    For A = 1 To UBound(Bug)
        With Bug(A)
            BugsRS.Seek "=", A
            If .Status > 0 Then
                If BugsRS.NoMatch Then
                    BugsRS.AddNew
                Else
                    BugsRS.Edit
                End If
                BugsRS!ID = A
                BugsRS!PlayerUser = .PlayerUser
                BugsRS!PlayerName = .PlayerName
                BugsRS!PlayerIP = .PlayerIP
                BugsRS!Title = .Title
                BugsRS!Description = .Description
                BugsRS!Status = .Status
                BugsRS!ResolverUser = .ResolverUser
                BugsRS!ResolverName = .ResolverName
                BugsRS.Update
            Else
                If Not BugsRS.NoMatch Then BugsRS.Delete
            End If
        End With
    Next A
ExitSub:
End Sub

Sub CheckGuild(Index As Long)
    If Guild(Index).Name <> vbNullString Then
        If CountGuildMembers(Index) < 1 Then
            'Not enough players -- delete guild
            DeleteGuild Index, 1
        End If
    End If
End Sub

Function CheckSum(St As String) As Long
    Dim A As Long, B As Long
    For A = 1 To Len(St)
        B = B + Asc(Mid$(St, A, 1))
    Next A
    CheckSum = B
End Function
Function CountGuildMembers(Index As Long) As Long
    Dim A As Long, B As Long
    With Guild(Index)
        If .Name <> vbNullString Then
            B = 0
            For A = 0 To 19
                If .Member(A).Name <> vbNullString Then
                    B = B + 1
                End If
            Next A
            CountGuildMembers = B
        End If
    End With
End Function

Sub DeleteCharacter()
    Dim A As Long, B As Long, St As String
    On Error Resume Next

    St = UserRS!Name
    For A = 1 To MaxGuilds
        With Guild(A)
            If .Name <> vbNullString Then
                For B = 0 To 19
                    With .Member(B)
                        If .Name = St Then
                            .Name = vbNullString
                            CheckGuild A
                        End If
                    End With
                Next B
            End If
        End With
    Next A

    On Error GoTo 0
End Sub
Sub DeleteAccount()
    On Error Resume Next

    If UserRS!Class > 0 Then
        DeleteCharacter
    End If

    UserRS.Delete

    On Error GoTo 0
End Sub
Sub LoadObjectData(ObjectData As String)
    Dim A As Long, NumObjects As Long
    NumObjects = Len(ObjectData) / 13 - 1
    For A = 0 To NumObjects
        With Map(Asc(Mid$(ObjectData, A * 13 + 1, 1)) * 256 + Asc(Mid$(ObjectData, A * 13 + 2, 1))).Object(Asc(Mid$(ObjectData, A * 13 + 3, 1)))
            .X = Asc(Mid$(ObjectData, A * 13 + 4, 1))
            .Y = Asc(Mid$(ObjectData, A * 13 + 5, 1))
            .Object = GetInt(Mid$(ObjectData, A * 13 + 6, 2))
            .Value = Asc(Mid$(ObjectData, A * 13 + 8, 1)) * 16777216 + Asc(Mid$(ObjectData, A * 13 + 9, 1)) * 65536 + Asc(Mid$(ObjectData, A * 13 + 10, 1)) * 256& + Asc(Mid$(ObjectData, A * 13 + 11, 1))
            .ItemPrefix = Asc(Mid$(ObjectData, A * 13 + 12, 1))
            .ItemSuffix = Asc(Mid$(ObjectData, A * 13 + 13, 1))
        End With
    Next A
End Sub
Function NPCNum(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To MaxNPCs
        With NPC(A)
            If UCase$(.Name) = Name Then
                NPCNum = A
                Exit Function
            End If
        End With
    Next A
End Function

Function FindInvObject(Index As Long, ObjectNum As Long) As Long
    Dim A As Long
    With Player(Index)
        For A = 1 To 20
            If .Inv(A).Object = ObjectNum Then
                FindInvObject = A
                Exit Function
            End If
        Next A
    End With
End Function

Function FindPlayer(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .InUse = True Then
                If UCase$(.Name) = Name Then
                    FindPlayer = A
                    Exit Function
                End If
            End If
        End With
    Next A
End Function

Function FreeBanNum() As Long
    Dim A As Long
    For A = 1 To 50
        If Ban(A).InUse = False Then
            FreeBanNum = A
            Exit For
        End If
    Next A
End Function
Function FreeInvNum(Index As Long) As Long
    Dim A As Long
    With Player(Index)
        For A = 1 To 20
            If .Inv(A).Object = 0 Then
                FreeInvNum = A
                Exit Function
            End If
        Next A
    End With
End Function
Function FreeMapDoorNum(MapNum As Long) As Long
    Dim A As Long
    With Map(MapNum)
        For A = 0 To 9
            If .Door(A).Att = 0 Then
                FreeMapDoorNum = A
                Exit Function
            End If
        Next A
    End With
    FreeMapDoorNum = -1
End Function
Function FreeMapObj(MapNum As Long) As Long
    Dim A As Long
    If MapNum >= 1 Then
        With Map(MapNum)
            For A = 0 To MaxMapObjects
                If .Object(A).Object = 0 Then
                    .Object(A).Value = 0
                    .Object(A).ItemPrefix = 0
                    .Object(A).ItemSuffix = 0
                    FreeMapObj = A
                    Exit Function
                End If
            Next A
        End With
    End If
    FreeMapObj = -1
End Function
Function FreePlayer() As Long
    Dim A As Long
    For A = 1 To MaxUsers
        If Player(A).InUse = False Then
            FreePlayer = A
            Exit Function
        End If
    Next A
End Function
Sub GainExp(Index As Long, Exp As Long)
    With Player(Index)
        If .Level < 80 Then
            If CDbl(.Experience) + CDbl(Exp) > 2147483647# Then
                .Experience = 2147483647
            Else
                .Experience = .Experience + Exp
            End If
            'Floating text
            SendToMap .Map, Chr$(112) + Chr$(13) + Chr$(.X) + Chr$(.Y) + CStr(Exp)
            If .Experience >= Int(1000 * CLng(.Level) ^ 1.3) Then
                If .Level < World.MaxLevel Then

                    .Level = .Level + 1
                    .Experience = 0

                    CalculateStats Index

                    SendSocket Index, Chr$(59) + Chr$(.MaxHP) + Chr$(.MaxEnergy) + Chr$(.MaxMana)
                End If
            End If
        End If
    End With
End Sub

Sub GainEliteExp(Index As Long, Exp As Long)
    With Player(Index)
        If .Level < World.MaxLevel Then
            If CDbl(.Experience) + CDbl(Exp) > 2147483647# Then
                .Experience = 2147483647
            Else
                .Experience = .Experience + Exp
            End If
            'Floating text
            SendToMap .Map, Chr$(112) + Chr$(13) + Chr$(.X) + Chr$(.Y) + CStr(Exp)
            If .Experience >= Int(1000 * CLng(.Level) ^ 1.3) Then
                If .Level < World.MaxLevel Then

                    .Level = .Level + 1
                    .Experience = 0

                    CalculateStats Index

                    SendSocket Index, Chr$(59) + Chr$(.MaxHP) + Chr$(.MaxEnergy) + Chr$(.MaxMana)
                End If
            End If
        End If
    End With
End Sub

Function GuildNum(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To MaxGuilds
        With Guild(A)
            If UCase$(.Name) = Name Then
                GuildNum = A
                Exit Function
            End If
        End With
    Next A
End Function
Function IsVacant(MapNum As Long, X As Long, Y As Long) As Boolean
    Dim A As Long

    If X < 0 Or Y < 0 Or X > 11 Or Y > 11 Then Exit Function

    With Map(MapNum)
        Select Case .Tile(X, Y).Att
        Case 1, 2, 3, 10, 13, 14, 15, 16    'Wall / Warp / Door / No Monsters
            Exit Function
        Case 19    'Light
            If ExamineBit(.Tile(X, Y).AttData(2), 0) Then
                Exit Function
            End If
        Case 20    'Light Dampening
            If ExamineBit(.Tile(X, Y).AttData(3), 0) Then
                Exit Function
            End If
        End Select
        Select Case .Tile(X, Y).Att2
        Case 1, 10, 13, 14, 15, 16
            Exit Function
        End Select

        For A = 0 To MaxMonsters
            With .Monster(A)
                If .Monster > 0 Then
                    If .X = X Then
                        If .Y = Y Then
                            Exit Function
                        End If
                    End If
                End If
            End With
        Next A

        For A = 1 To MaxUsers
            With Player(A)
                If .Map = MapNum Then
                    If .X = X Then
                        If .Y = Y Then
                            If .IsDead = False Then
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End With
        Next A
    End With

    IsVacant = True
End Function
Function PlayerIsVacant(MapNum As Long, X As Long, Y As Long) As Boolean
    Dim A As Long

    With Map(MapNum)
        Select Case .Tile(X, Y).Att
        Case 1, 2, 3, 13, 14, 15, 16    'Wall / Warp
            Exit Function
        Case 19    'Light
            If ExamineBit(.Tile(X, Y).AttData(2), 0) Then
                Exit Function
            End If
        Case 20    'Light Dampening
            If ExamineBit(.Tile(X, Y).AttData(3), 0) Then
                Exit Function
            End If
        End Select
        Select Case .Tile(X, Y).Att2
        Case 1, 2, 3, 13, 14, 15, 16
            Exit Function
        End Select

        For A = 0 To MaxMonsters
            With .Monster(A)
                If .Monster > 0 Then
                    If .X = X Then
                        If .Y = Y Then
                            Exit Function
                        End If
                    End If
                End If
            End With
        Next A

        For A = 1 To MaxUsers
            With Player(A)
                If .Map = MapNum Then
                    If .X = X Then
                        If .Y = Y Then
                            If Not .Status = 25 Then
                                If .IsDead = False Then
                                    If .Guild > 0 Then
                                        If Player(A).Guild = 0 Then
                                            If ExamineBit(Map(.Map).flags, 0) = False And ExamineBit(Map(.Map).flags, 6) = False Then

                                            Else
                                                Exit Function
                                            End If
                                        Else
                                            Exit Function
                                        End If
                                    Else
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next A
    End With

    PlayerIsVacant = True
End Function
Sub JoinGame(Index As Long)
    Dim A As Long, St1 As String, Tick As Currency
    
    Tick = getTime()

    With Player(Index)
        .SpeedHackTimer = Tick - 120000
        If .Class = 0 Then
            .Mode = modeBanned
            Hacker Index, "C.0"
            Exit Sub
        End If
        .Mode = modePlaying
        SendAllBut Index, Chr$(6) + Chr$(Index) + DoubleChar$(CLng(.Sprite)) + Chr$(.Status) + Chr$(.Guild) + Chr$(.MaxHP) + .Name
        SendToGods Chr$(16) + Chr$(0) + .User + " - " + .IP
        
        If .Access > 0 Then PrintGodSilent .User, " (Joined Game) "
        PrintAccount "User: " + .User + " - Name " + .Name + " - IP: " + .IP + " - ID: " + .ComputerID
        
        St1 = DoubleChar(1) + Chr$(24)

        A = .Map
        If Map(A).BootLocation.Map > 0 Then
            'Move player if not allowed to join on this map
            .Map = Map(A).BootLocation.Map
            .X = Map(A).BootLocation.X
            .Y = Map(A).BootLocation.Y
        End If

        If .Map < 1 Then .Map = 1
        If .Map > MaxMaps Then .Map = MaxMaps
        If .X > 11 Then .X = 11
        If .Y > 11 Then .Y = 11

        'Send Player Data
        For A = 1 To MaxUsers
            If A <> Index Then
                With Player(A)
                    If .Mode = modePlaying Then
                        St1 = St1 + DoubleChar(7 + Len(.Name)) + Chr$(6) + Chr$(A) + DoubleChar$(CLng(.Sprite)) + Chr$(.Status) + Chr$(.Guild) + Chr$(.MaxHP) + .Name
                        If Len(St1) > 1024 Then
                            SendRaw Index, St1
                            St1 = vbNullString
                        End If
                    End If
                End With
            End If
        Next A

        'Send Inventory Data
        For A = 1 To 20
            If .Inv(A).Object > 0 Then
                St1 = St1 + DoubleChar$(10) + Chr$(17) + Chr$(A) + DoubleChar$(CLng(.Inv(A).Object)) + QuadChar(.Inv(A).Value) + Chr$(.Inv(A).ItemPrefix) + Chr$(.Inv(A).ItemSuffix)
                If Len(St1) > 1024 Then
                    SendRaw Index, St1
                    St1 = vbNullString
                End If
            End If
        Next A

        If .EquippedObject(6).Object > 0 Then St1 = St1 + DoubleChar(2) + Chr$(19) + Chr$(.EquippedObject(6).Object)

        For A = 1 To 5
            If .EquippedObject(A).Object > 0 Then
                St1 = St1 + DoubleChar(9) + Chr$(115) + DoubleChar$(CLng(.EquippedObject(A).Object)) + QuadChar(.EquippedObject(A).Value) + Chr$(.EquippedObject(A).ItemPrefix) + Chr$(.EquippedObject(A).ItemSuffix)
            End If
        Next A

        If Len(St1) > 0 Then
            SendRaw Index, St1
        End If

        SendSocket Index, Chr$(143) + Chr$(OutdoorLight)

        JoinMap Index

        Parameter(0) = Index
        RunScript "JOINGAME"

        CalculateStats Index

        'Send Guild Data
        If .Guild > 0 Then
            St1 = vbNullString
            With Guild(.Guild)
                For A = 0 To DeclarationCount
                    With .Declaration(A)
                        St1 = St1 + DoubleChar(4) + Chr$(71) + Chr$(A) + Chr$(.Guild) + Chr$(.Type)
                    End With
                Next A

                If .Bank >= 0 Then
                    St1 = St1 + DoubleChar(9) + Chr$(152) + QuadChar(.Bank) + QuadChar$(GetGuildUpkeep(CLng(Player(Index).Guild)))
                Else
                    St1 = St1 + DoubleChar(9) + Chr$(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                End If

                If Not .MOTD = "" Then
                    SendSocket Index, Chr$(56) + Chr$(14) + .Name + " - " + .MOTDCreator + " - " + CStr(CDate(.MOTDDate)) + " - " + .MOTD
                Else
                    SendSocket Index, Chr$(16) + Chr$(43)
                End If
            End With
            If Len(St1) > 0 Then
                SendRaw Index, St1
            End If
        End If
    End With
End Sub

Sub JoinMap(Index As Long)
    Dim A As Long, MapNum As Long, St1 As String, Tick As Currency

    Tick = getTime

    With Player(Index)
        MapNum = .Map

        If Map(MapNum).NumPlayers = 0 And Map(MapNum).ResetTimer > 0 And Tick - Map(MapNum).ResetTimer >= 120000 And ExamineBit(Map(MapNum).Flags2, 2) = 0 Then
            ResetMap MapNum
        End If

        With Map(MapNum)
            .NumPlayers = .NumPlayers + 1
        End With
        St1 = DoubleChar(14) + Chr$(12) + DoubleChar(CLng(MapNum)) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + QuadChar(Map(MapNum).Version) + QuadChar(Map(MapNum).CheckSum)

        With Map(MapNum)
            For A = 0 To 9
                If .Door(A).Att > 0 Then
                    If Tick - .Door(A).T > 10000 Then
                        .Tile(.Door(A).X, .Door(A).Y).Att = .Door(A).Att
                        .Door(A).Att = 0
                    End If
                End If
            Next A
        End With

        'Send Door Data
        For A = 0 To 9
            With Map(MapNum).Door(A)
                If .Att > 0 Then
                    St1 = St1 + DoubleChar(4) + Chr$(36) + Chr$(A) + Chr$(.X) + Chr$(.Y)
                End If
            End With
        Next A

        'Send Player Data
        For A = 1 To MaxUsers
            If Player(A).Mode = modePlaying And Player(A).Map = MapNum And A <> Index Then
                With Player(A)
                    St1 = St1 + DoubleChar(8) + Chr$(8) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + DoubleChar$(CLng(.Sprite)) + Chr$(.Status)
                End With
                If Len(St1) > 1024 Then
                    SendRaw Index, St1
                    St1 = vbNullString
                End If
            End If
        Next A

        With Map(MapNum)
            'Send Map Monster Data
            For A = 0 To MaxMonsters
                With .Monster(A)
                    If .Monster > 0 Then
                        St1 = St1 + DoubleChar(9) + Chr$(38) + Chr$(A) + DoubleChar$(CLng(.Monster)) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + DoubleChar$(CLng(.HP))
                    End If
                End With
            Next A

            'Send Map Object Data
            For A = 0 To MaxMapObjects
                With .Object(A)
                    If .Object > 0 Then
                        St1 = St1 + DoubleChar(12) + Chr$(14) + Chr$(A) + DoubleChar$(CLng(.Object)) + Chr$(.X) + Chr$(.Y) + Chr$(.ItemPrefix) + Chr$(.ItemSuffix) + QuadChar(.Value)
                    End If
                    If Len(St1) > 1024 Then
                        SendRaw Index, St1
                        St1 = vbNullString
                    End If
                End With
            Next A
        End With
        
        A = Map(MapNum).NPC
        If A >= 1 Then
            With NPC(A)
                If .JoinText <> vbNullString Then
                    St1 = St1 + DoubleChar(3 + Len(.JoinText)) + Chr$(88) + DoubleChar$(A) + .JoinText
                End If
            End With
        End If

        If St1 <> vbNullString Then
            SendRaw Index, St1
        End If
        SendToMapAllBut MapNum, Index, Chr$(8) + Chr$(Index) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + DoubleChar$(CLng(.Sprite)) + Chr$(.Status)

        Parameter(0) = Index
        RunScript "JOINMAP" + CStr(MapNum)
    End With
End Sub

Sub MapWarp(Index As Long)
    With Player(Index)
        SendSocket Index, Chr$(147) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
        SendToMapAllBut .Map, Index, Chr$(8) + Chr$(Index) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + DoubleChar$(CLng(.Sprite)) + Chr$(.Status)
    End With
End Sub

Sub LoadMap(MapNum As Long, MapData As String)
    Dim A As Long, X As Long, Y As Long
    If Len(MapData) = 2677 Then
        'Characters 1-30 = Name
        '36 = Midi
        With Map(MapNum)
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .CheckSum = CheckSum(MapData)
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1)) * 256 + Asc(Mid$(MapData, 36, 1))
            .Midi = Asc(Mid$(MapData, 37, 1))
            .ExitUp = Asc(Mid$(MapData, 38, 1)) * 256 + Asc(Mid$(MapData, 39, 1))
            .ExitDown = Asc(Mid$(MapData, 40, 1)) * 256 + Asc(Mid$(MapData, 41, 1))
            .ExitLeft = Asc(Mid$(MapData, 42, 1)) * 256 + Asc(Mid$(MapData, 43, 1))
            .ExitRight = Asc(Mid$(MapData, 44, 1)) * 256 + Asc(Mid$(MapData, 45, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 46, 1)) * 256 + Asc(Mid$(MapData, 47, 1))
            .BootLocation.X = Asc(Mid$(MapData, 48, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 49, 1))
            .DeathLocation.Map = Asc(Mid$(MapData, 50, 1)) * 256 + Asc(Mid$(MapData, 51, 1))
            .DeathLocation.X = Asc(Mid$(MapData, 52, 1))
            .DeathLocation.Y = Asc(Mid$(MapData, 53, 1))
            .flags = Asc(Mid$(MapData, 54, 1))
            .Flags2 = Asc(Mid$(MapData, 55, 1))
            For A = 0 To 9    '56 - 86
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 56 + A * 3)) * 256 + Asc(Mid$(MapData, 57 + A * 3))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 58 + A * 3))
            Next A
            '86
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 86 + Y * 216 + X * 18
                        '1-10 = Tiles
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 9, 1))
                        .FGTile2 = Asc(Mid$(MapData, A + 10, 1)) * 256 + Asc(Mid$(MapData, A + 11, 1))
                        .Att = Asc(Mid$(MapData, A + 12, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 13, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 14, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 15, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 16, 1))
                        .Att2 = Asc(Mid$(MapData, A + 17, 1))
                        Select Case .Att
                        Case 5
                            Map(MapNum).Keep = True
                        Case 8
                            If .AttData(2) > 0 Then
                                Map(MapNum).Hall = .AttData(2)
                            End If
                        End Select
                        Select Case .Att2
                        Case 5
                            Map(MapNum).Keep = True
                        End Select
                    End With
                Next X
            Next Y
        End With
    End If
End Sub

Sub LoadMapOld(MapNum As Long, MapData As String)
    Dim A As Long, X As Long, Y As Long
    MsgBox Len(MapData)
    If Len(MapData) = 2388 Then
        'Characters 1-30 = Name
        '36 = Midi
        With Map(MapNum)
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .CheckSum = CheckSum(MapData)
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1))
            .Midi = Asc(Mid$(MapData, 36, 1))
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256 + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256 + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256 + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256 + Asc(Mid$(MapData, 44, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 45, 1)) * 256 + Asc(Mid$(MapData, 46, 1))
            .BootLocation.X = Asc(Mid$(MapData, 47, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 48, 1))
            .DeathLocation.Map = Asc(Mid$(MapData, 49, 1)) * 256 + Asc(Mid$(MapData, 50, 1))
            .DeathLocation.X = Asc(Mid$(MapData, 51, 1))
            .DeathLocation.Y = Asc(Mid$(MapData, 52, 1))
            .flags = Asc(Mid$(MapData, 53, 1))
            .Flags2 = Asc(Mid$(MapData, 54, 1))
            For A = 0 To 9    '55 - 85
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 55 + A * 3)) * 256 + Asc(Mid$(MapData, 56 + A * 3))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 57 + A * 3))
            Next A
            '86
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 85 + Y * 192 + X * 16
                        '1-8 = Tiles
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 9, 1))
                        .Att = Asc(Mid$(MapData, A + 10, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 11, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 12, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 13, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 14, 1))
                        .Att2 = Asc(Mid$(MapData, A + 15, 1))
                        Select Case .Att
                        Case 5
                            Map(MapNum).Keep = True
                        Case 8
                            If .AttData(2) > 0 Then
                                Map(MapNum).Hall = .AttData(2)
                            End If
                        End Select
                        Select Case .Att2
                        Case 5
                            Map(MapNum).Keep = True
                        End Select
                    End With
                Next X
            Next Y
        End With
    End If
End Sub

Sub LoadMapOld2008(MapNum As Long, MapData As String)
    Dim A As Long, X As Long, Y As Long
    If Len(MapData) = 2359 Then
        'Characters 1-30 = Name
        '36 = Midi
        With Map(MapNum)
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .CheckSum = CheckSum(MapData)
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1))
            .Midi = Asc(Mid$(MapData, 36, 1))
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256 + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256 + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256 + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256 + Asc(Mid$(MapData, 44, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 45, 1)) * 256 + Asc(Mid$(MapData, 46, 1))
            .BootLocation.X = Asc(Mid$(MapData, 47, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 48, 1))
            .flags = Asc(Mid$(MapData, 49, 1))
            For A = 0 To 2 '50 - 55
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 50 + A * 2))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 51 + A * 2))
            Next A
            '56
            .Keep = False
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 56 + Y * 192 + X * 16
                        '1-8 = Tiles
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 9, 1))
                        .FGTile2 = 0
                        .Att = Asc(Mid$(MapData, A + 10, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 11, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 12, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 13, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 14, 1))
                        .Att2 = Asc(Mid$(MapData, A + 15, 1))
                        Select Case .Att
                            Case 5
                                Map(MapNum).Keep = True
                            Case 8
                                If .AttData(2) > 0 Then
                                    Map(MapNum).Hall = .AttData(2)
                                End If
                        End Select
                        Select Case .Att2
                            Case 5
                                Map(MapNum).Keep = True
                        End Select
                    End With
                Next X
            Next Y
        End With
    End If
End Sub

Sub LoadMapOld1997(MapNum As Long, MapData As String)
    Dim A As Long, X As Long, Y As Long
    If Len(MapData) = 1927 Then
        'Characters 1-30 = Name
        '36 = Midi
        With Map(MapNum)
            .CheckSum = CheckSum(MapData)
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1))
            .Midi = Asc(Mid$(MapData, 36, 1))
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256 + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256 + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256 + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256 + Asc(Mid$(MapData, 44, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 45, 1)) * 256 + Asc(Mid$(MapData, 46, 1))
            .BootLocation.X = Asc(Mid$(MapData, 47, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 48, 1))
            .flags = Asc(Mid$(MapData, 49, 1))
            For A = 0 To 2 '50 - 55
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 50 + A * 2))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 51 + A * 2))
            Next A
            '56
            .Keep = False
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 56 + Y * 156 + X * 13
                        '1-8 = Tiles
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = 0
                        .BGTile1 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .FGTile = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile2 = 0
                        .Att = Asc(Mid$(MapData, A + 8, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 9, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 10, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 11, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 12, 1))
                        .Att2 = 0
                        Select Case .Att
                            Case 5
                                Map(MapNum).Keep = True
                            Case 8
                                If .AttData(2) > 0 Then
                                    Map(MapNum).Hall = .AttData(2)
                                End If
                        End Select
                    End With
                Next X
            Next Y
        End With
    Else
        MsgBox Len(MapData)
    End If
End Sub

Sub Main()

    Randomize
    timeBeginPeriod 1
    
    Dim A As Long
    Dim St As String
    Dim LingerType As LingerType

    Startup = True

    InitFunctionTable

    frmLoading.Show
    frmLoading.Refresh

    On Error Resume Next
    MkDir "log"
    MkDir "log/debug"
    MkDir "log/god"
    MkDir "log/password"
    MkDir "log/cheat"
    MkDir "log/script"
    MkDir "log/account"
    MkDir "log/items"
    MkDir "log/chat"
    MkDir "log/chat/guild"
    MkDir "log/chat/god"
    MkDir "log/chat/say"
    MkDir "log/chat/yell"
    MkDir "log/chat/emote"
    MkDir "log/chat/broadcast"
    MkDir "log/chat/tell"
    MkDir "scriptini"
    On Error GoTo 0

    LoadDatabase

    For A = 1 To MaxMaps
        ResetMap A
    Next A

    frmLoading.lblStatus = "Initializing Sockets.."
    frmLoading.lblStatus.Refresh

    Load frmMain
    frmMain.Caption = TitleString + " [0]"
    Hook
    StartWinsock St

    'Listen for connections
    With LingerType
        .l_onoff = 1
        .l_linger = 0
    End With

    ListeningSocket = ListenForConnect(World.ServerPort, gHW, 1025)
    If ListeningSocket = INVALID_SOCKET Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If
    If setsockopt(ListeningSocket, SOL_SOCKET, SO_LINGER, LingerType, 4) <> 0 Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
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
    
    Unload frmLoading

    frmMain.Show
    Startup = False
    PrintLog ("Odyssey Realms Registry Version B" + CStr(CurrentClientVer) + ".")
    PrintLog "Attempting to connect to the Registry..."
    ConnectToRegistry
End Sub
Function NewMapMonster(MapNum As Long, MonsterNum As Long) As String
    Dim TX As Long, TY As Long, TriesLeft As Long
    Dim MonsterType As Long, MonsterFlags As Byte

    If Int(MonsterNum / 2) * 2 = MonsterNum Or ExamineBit(Map(MapNum).flags, 4) = True Then
        MonsterType = Map(MapNum).MonsterSpawn(Int(MonsterNum / 2)).Monster
        If MonsterType > 0 Then
            MonsterFlags = Monster(MonsterType).flags
            Randomize
            TX = Int(Rnd * 12)
            TY = Int(Rnd * 12)
            TriesLeft = 20
            While TriesLeft > 0 And (Map(MapNum).Tile(TX, TY).Att > 0 Or Map(MapNum).Tile(TX, TY).Att2 > 0)
                TX = Int(Rnd * 12)
                TY = Int(Rnd * 12)
                TriesLeft = TriesLeft - 1
            Wend
            If TriesLeft > 0 Then
                NewMapMonster = SpawnMapMonster(MapNum, MonsterNum, MonsterType, TX, TY)
            End If
        End If
    End If
End Function
Function NewMapObject(MapNum As Long, ObjectNum As Long, Value As Long, X As Long, Y As Long, Infinite As Boolean) As Long
    Dim A As Long
    If MapNum >= 1 Then
        A = FreeMapObj(MapNum)
        If A >= 0 Then
            With Map(MapNum).Object(A)
                .Object = ObjectNum
                .X = X
                .Y = Y
                If Infinite = True Then
                    .TimeStamp = 0
                Else
                    .TimeStamp = getTime() + Int(Rnd * 60000) - 30000
                End If
                Select Case Object(ObjectNum).Type
                Case 1, 2, 3, 4    'Weapon, Shield, Armor, Helmut
                    .Value = CLng(Object(ObjectNum).Data(0)) * 10
                    .ItemPrefix = RandomPrefix
                    .ItemSuffix = RandomSuffix
                Case 6, 11    'Money
                    If Value < 1 Then Value = 1
                    .Value = .Value + Value
                Case 8    'Ring
                    .Value = CLng(Object(ObjectNum).Data(1)) * 10
                    .ItemPrefix = RandomPrefix
                    .ItemSuffix = RandomSuffix
                Case Else
                    .Value = 0
                End Select
                SendToMap MapNum, Chr$(14) + Chr$(A) + DoubleChar$(ObjectNum) + Chr$(X) + Chr$(Y) + Chr$(.ItemPrefix) + Chr$(.ItemSuffix) + QuadChar$(.Value)
            End With
            NewMapObject = True
        End If
    End If
End Function
Sub Partmap(Index As Long)
    Dim A As Long, MapNum As Long

    With Player(Index)
        MapNum = .Map
        If MapNum > 0 Then
            Parameter(0) = Index
            RunScript "PARTMAP" + CStr(MapNum)

            With Map(MapNum)
                .NumPlayers = .NumPlayers - 1
                For A = 0 To MaxMonsters
                    With .Monster(A)
                        If .Target = Index And .TargetIsMonster = False And .Monster > 0 Then
                            .Target = 0
                            .TargetIsMonster = False
                        End If
                    End With
                Next A
                If .NumPlayers = 0 Then
                    .ResetTimer = getTime()
                End If
            End With
            SendToMapAllBut MapNum, Index, Chr$(9) + Chr$(Index)

            If .Socket <> INVALID_SOCKET Then
                A = Map(MapNum).NPC
                If A >= 1 Then
                    With NPC(A)
                        If .LeaveText <> vbNullString Then
                            SendSocket Index, Chr$(88) + DoubleChar$(A) + .LeaveText
                        End If
                    End With
                End If
            End If

            .Map = 0
        End If
    End With
End Sub

Function PlayerDied(Index As Long, Killer As Long) As Boolean
    PlayerDied = False
    Dim A As Long, B As Long, C As Long, D As Long, St1 As String, St2 As String, Tick As Currency
    Tick = getTime()
    Dim DontDropOnGround As Boolean
    Dim MapNum As Long
    Parameter(0) = Index
    Player(Index).IsDead = True
    Player(Index).DeadTick = Tick + World.DeathTime * 1000
    If ExamineBit(Map(Player(Index).Map).flags, 7) = True Then    'Map is an arena
        PlayerDied = False
        Exit Function
    End If
    If Not RunScript("PLAYERDIE") = 0 Then
        PlayerDied = False
        Exit Function
    End If
    If Not Index = Killer Then
        If Player(Index).Status = 1 Then Player(Index).Status = 0
    End If
    SetPlayerStatus Index, Player(Index).Status

    With Player(Index)
        St1 = vbNullString
        St2 = vbNullString
        MapNum = .Map
        For A = 1 To 20
            If .Inv(A).Object > 0 Then
                C = 0
                If .EquippedObject(6).Object = A Then C = 1
                Parameter(0) = Index
                Parameter(1) = .Inv(A).Value
                If Not ExamineBit(Object(.Inv(A).Object).flags, 2) = 255 And C = 1 And RunScript("DROPOBJ" + CStr(.Inv(A).Object)) = 0 Then
                    DontDropOnGround = False
                    If Killer > -1 And Not Killer = Index Then
                        Select Case Object(.Inv(A).Object).Type
                        Case 6, 11
                            D = FindInvObject(Killer, CLng(.Inv(A).Object))
                            If D = 0 Then D = FreeInvNum(Killer)
                        Case Else
                            D = FreeInvNum(Killer)
                        End Select

                        If D > 0 Then
                            Parameter(0) = Killer
                            Parameter(1) = .Inv(A).Value
                            If RunScript("GETOBJ" + CStr(.Inv(A).Object)) = 0 Then
                                Select Case Object(.Inv(A).Object).Type
                                    Case 6, 11
                                        If Player(Killer).Inv(D).Object > 0 Then
                                            Player(Killer).Inv(D).Value = Player(Killer).Inv(D).Value + .Inv(A).Value
                                        Else
                                            Player(Killer).Inv(D).Object = .Inv(A).Object
                                            Player(Killer).Inv(D).Value = .Inv(A).Value
                                            Player(Killer).Inv(D).ItemPrefix = .Inv(A).ItemPrefix
                                            Player(Killer).Inv(D).ItemSuffix = .Inv(A).ItemSuffix
                                        End If
                                    Case Else
                                        Player(Killer).Inv(D).Object = .Inv(A).Object
                                        Player(Killer).Inv(D).Value = .Inv(A).Value
                                        Player(Killer).Inv(D).ItemPrefix = .Inv(A).ItemPrefix
                                        Player(Killer).Inv(D).ItemSuffix = .Inv(A).ItemSuffix
                                End Select
                                SendSocket Killer, Chr$(17) + Chr$(D) + DoubleChar$(CLng(Player(Killer).Inv(D).Object)) + QuadChar(Player(Killer).Inv(D).Value) + Chr$(Player(Killer).Inv(D).ItemPrefix) + Chr$(Player(Killer).Inv(D).ItemSuffix)
                                DontDropOnGround = True
                            End If
                        End If
                    End If

                    If DontDropOnGround = False Then
                        B = FreeMapObj(MapNum)
                        If B >= 0 Then
                            Map(MapNum).Object(B).X = .X
                            Map(MapNum).Object(B).Y = .Y
                            Map(MapNum).Object(B).ItemPrefix = .Inv(A).ItemPrefix
                            Map(MapNum).Object(B).ItemSuffix = .Inv(A).ItemSuffix
                            Map(MapNum).Object(B).Object = .Inv(A).Object
                            Map(MapNum).Object(B).Value = .Inv(A).Value
                            Map(MapNum).Object(B).TimeStamp = 1
                            St1 = St1 + DoubleChar(12) + Chr$(14) + Chr$(B) + DoubleChar$(CLng(.Inv(A).Object)) + Chr$(.X) + Chr$(.Y) + Chr$(Map(MapNum).Object(B).ItemPrefix) + Chr$(Map(MapNum).Object(B).ItemSuffix) + QuadChar$(Map(MapNum).Object(B).Value)
                        End If
                    End If

                    .Inv(A).Object = 0
                    .Inv(A).Value = 0
                    .Inv(A).ItemPrefix = 0
                    .Inv(A).ItemSuffix = 0
                    St2 = St2 + DoubleChar(2) + Chr$(18) + Chr$(A)
                End If
            End If
        Next A

        Dim RandomDrop As Byte
        Randomize
        RandomDrop = Random(5) + 1
        If .EquippedObject(RandomDrop).Object > 0 Then
            Parameter(0) = Index
            Parameter(1) = .EquippedObject(RandomDrop).Value
            If Not ExamineBit(Object(.EquippedObject(RandomDrop).Object).flags, 2) And RunScript("DROPOBJ" + CStr(.EquippedObject(RandomDrop).Object)) = 0 Then
                DontDropOnGround = False
                If Killer > -1 And Not Killer = Index Then
                    Select Case Object(.EquippedObject(RandomDrop).Object).Type
                    Case 6, 11
                        D = FindInvObject(Killer, CLng(.EquippedObject(RandomDrop).Object))
                        If D = 0 Then D = FreeInvNum(Killer)
                    Case Else
                        D = FreeInvNum(Killer)
                    End Select

                    If D > 0 Then
                        Parameter(0) = Killer
                        Parameter(1) = .EquippedObject(RandomDrop).Value
                        If RunScript("GETOBJ" + CStr(.EquippedObject(RandomDrop).Object)) = 0 Then
                            Select Case Object(.EquippedObject(RandomDrop).Object).Type
                            Case 6, 11
                                If Player(Killer).Inv(D).Object > 0 Then
                                    Player(Killer).Inv(D).Value = Player(Killer).Inv(D).Value + .EquippedObject(RandomDrop).Value
                                Else
                                    Player(Killer).Inv(D).Object = .EquippedObject(RandomDrop).Object
                                    Player(Killer).Inv(D).Value = .EquippedObject(RandomDrop).Value
                                    Player(Killer).Inv(D).ItemPrefix = .EquippedObject(RandomDrop).ItemPrefix
                                    Player(Killer).Inv(D).ItemSuffix = .EquippedObject(RandomDrop).ItemSuffix
                                End If
                            Case Else
                                Player(Killer).Inv(D).Object = .EquippedObject(RandomDrop).Object
                                Player(Killer).Inv(D).Value = .EquippedObject(RandomDrop).Value
                                Player(Killer).Inv(D).ItemPrefix = .EquippedObject(RandomDrop).ItemPrefix
                                Player(Killer).Inv(D).ItemSuffix = .EquippedObject(RandomDrop).ItemSuffix
                            End Select
                            SendSocket Killer, Chr$(17) + Chr$(D) + DoubleChar$(CLng(Player(Killer).Inv(D).Object)) + QuadChar(Player(Killer).Inv(D).Value) + Chr$(Player(Killer).Inv(D).ItemPrefix) + Chr$(Player(Killer).Inv(D).ItemSuffix)
                            DontDropOnGround = True
                        End If
                    End If
                End If

                If DontDropOnGround = False Then
                    B = FreeMapObj(MapNum)
                    If B >= 0 Then
                        Map(MapNum).Object(B).X = .X
                        Map(MapNum).Object(B).Y = .Y
                        Map(MapNum).Object(B).ItemPrefix = .EquippedObject(RandomDrop).ItemPrefix
                        Map(MapNum).Object(B).ItemSuffix = .EquippedObject(RandomDrop).ItemSuffix
                        Map(MapNum).Object(B).Object = .EquippedObject(RandomDrop).Object
                        Map(MapNum).Object(B).Value = .EquippedObject(RandomDrop).Value
                        Map(MapNum).Object(B).TimeStamp = 1
                        St1 = St1 + DoubleChar(12) + Chr$(14) + Chr$(B) + DoubleChar$(CLng(.EquippedObject(RandomDrop).Object)) + Chr$(.X) + Chr$(.Y) + Chr$(Map(MapNum).Object(B).ItemPrefix) + Chr$(Map(MapNum).Object(B).ItemSuffix) + QuadChar$(Map(MapNum).Object(B).Value)
                    End If
                End If

                .EquippedObject(RandomDrop).Object = 0
                .EquippedObject(RandomDrop).Value = 0
                .EquippedObject(RandomDrop).ItemPrefix = 0
                .EquippedObject(RandomDrop).ItemSuffix = 0
                St2 = St2 + DoubleChar(2) + Chr$(20) + Chr$(RandomDrop + 20)
            Else
                For A = 1 To 5
                    If .EquippedObject(A).Object > 0 Then
                        Parameter(0) = Index
                        Parameter(1) = .EquippedObject(A).Value
                        If Not ExamineBit(Object(.EquippedObject(A).Object).flags, 2) And RunScript("DROPOBJ" + CStr(.EquippedObject(A).Object)) = 0 Then
                            DontDropOnGround = False

                            If Killer > -1 And Not Killer = Index Then
                                Select Case Object(.EquippedObject(A).Object).Type
                                Case 6, 11
                                    D = FindInvObject(Killer, CLng(.EquippedObject(A).Object))
                                    If D = 0 Then D = FreeInvNum(Killer)
                                Case Else
                                    D = FreeInvNum(Killer)
                                End Select

                                If D > 0 Then
                                    Parameter(0) = Killer
                                    Parameter(1) = .EquippedObject(A).Value
                                    If RunScript("GETOBJ" + CStr(.EquippedObject(A).Object)) = 0 Then
                                        Select Case Object(.EquippedObject(A).Object).Type
                                        Case 6, 11
                                            If Player(Killer).Inv(D).Object > 0 Then
                                                Player(Killer).Inv(D).Value = Player(Killer).Inv(D).Value + .EquippedObject(A).Value
                                            Else
                                                Player(Killer).Inv(D).Object = .EquippedObject(A).Object
                                                Player(Killer).Inv(D).Value = .EquippedObject(A).Value
                                                Player(Killer).Inv(D).ItemPrefix = .EquippedObject(A).ItemPrefix
                                                Player(Killer).Inv(D).ItemSuffix = .EquippedObject(A).ItemSuffix
                                            End If
                                        Case Else
                                            Player(Killer).Inv(D).Object = .EquippedObject(A).Object
                                            Player(Killer).Inv(D).Value = .EquippedObject(A).Value
                                            Player(Killer).Inv(D).ItemPrefix = .EquippedObject(A).ItemPrefix
                                            Player(Killer).Inv(D).ItemSuffix = .EquippedObject(A).ItemSuffix
                                        End Select
                                        SendSocket Killer, Chr$(17) + Chr$(D) + DoubleChar$(CLng(Player(Killer).Inv(D).Object)) + QuadChar(Player(Killer).Inv(D).Value) + Chr$(Player(Killer).Inv(D).ItemPrefix) + Chr$(Player(Killer).Inv(D).ItemSuffix)
                                        DontDropOnGround = True
                                    End If
                                End If
                            End If

                            If DontDropOnGround = False Then
                                B = FreeMapObj(MapNum)
                                If B >= 0 Then
                                    Map(MapNum).Object(B).X = .X
                                    Map(MapNum).Object(B).Y = .Y
                                    Map(MapNum).Object(B).ItemPrefix = .EquippedObject(A).ItemPrefix
                                    Map(MapNum).Object(B).ItemSuffix = .EquippedObject(A).ItemSuffix
                                    Map(MapNum).Object(B).Object = .EquippedObject(A).Object
                                    Map(MapNum).Object(B).Value = .EquippedObject(A).Value
                                    Map(MapNum).Object(B).TimeStamp = 1
                                    St1 = St1 + DoubleChar(12) + Chr$(14) + Chr$(B) + DoubleChar$(CLng(.EquippedObject(A).Object)) + Chr$(.X) + Chr$(.Y) + Chr$(Map(MapNum).Object(B).ItemPrefix) + Chr$(Map(MapNum).Object(B).ItemSuffix) + QuadChar$(Map(MapNum).Object(B).Value)
                                End If
                            End If

                            .EquippedObject(A).Object = 0
                            .EquippedObject(A).Value = 0
                            .EquippedObject(A).ItemPrefix = 0
                            .EquippedObject(A).ItemSuffix = 0
                            St2 = St2 + DoubleChar(2) + Chr$(20) + Chr$(A + 20)
                            Exit For
                        End If
                    End If
                Next A
            End If
        Else
            For A = 1 To 5
                If .EquippedObject(A).Object > 0 Then
                    Parameter(0) = Index
                    Parameter(1) = .EquippedObject(A).Value
                    If Not ExamineBit(Object(.EquippedObject(A).Object).flags, 2) = 255 And RunScript("DROPOBJ" + CStr(.EquippedObject(A).Object)) = 0 Then
                        DontDropOnGround = False

                        If Killer > -1 And Not Killer = Index Then
                            Select Case Object(.EquippedObject(A).Object).Type
                            Case 6, 11
                                D = FindInvObject(Killer, CLng(.EquippedObject(A).Object))
                                If D = 0 Then D = FreeInvNum(Killer)
                            Case Else
                                D = FreeInvNum(Killer)
                            End Select

                            If D > 0 Then
                                Parameter(0) = Killer
                                Parameter(1) = .EquippedObject(A).Value
                                If RunScript("GETOBJ" + CStr(.EquippedObject(A).Object)) = 0 Then
                                    Select Case Object(.EquippedObject(A).Object).Type
                                    Case 6, 11
                                        If Player(Killer).Inv(D).Object > 0 Then
                                            Player(Killer).Inv(D).Value = Player(Killer).Inv(D).Value + .EquippedObject(A).Value
                                        Else
                                            Player(Killer).Inv(D).Object = .EquippedObject(A).Object
                                            Player(Killer).Inv(D).Value = .EquippedObject(A).Value
                                            Player(Killer).Inv(D).ItemPrefix = .EquippedObject(A).ItemPrefix
                                            Player(Killer).Inv(D).ItemSuffix = .EquippedObject(A).ItemSuffix
                                        End If
                                    Case Else
                                        Player(Killer).Inv(D).Object = .EquippedObject(A).Object
                                        Player(Killer).Inv(D).Value = .EquippedObject(A).Value
                                        Player(Killer).Inv(D).ItemPrefix = .EquippedObject(A).ItemPrefix
                                        Player(Killer).Inv(D).ItemSuffix = .EquippedObject(A).ItemSuffix
                                    End Select
                                    SendSocket Killer, Chr$(17) + Chr$(D) + DoubleChar$(CLng(Player(Killer).Inv(D).Object)) + QuadChar(Player(Killer).Inv(D).Value) + Chr$(Player(Killer).Inv(D).ItemPrefix) + Chr$(Player(Killer).Inv(D).ItemSuffix)
                                    DontDropOnGround = True
                                End If
                            End If
                        End If

                        If DontDropOnGround = False Then
                            B = FreeMapObj(MapNum)
                            If B >= 0 Then
                                Map(MapNum).Object(B).X = .X
                                Map(MapNum).Object(B).Y = .Y
                                Map(MapNum).Object(B).ItemPrefix = .EquippedObject(A).ItemPrefix
                                Map(MapNum).Object(B).ItemSuffix = .EquippedObject(A).ItemSuffix
                                Map(MapNum).Object(B).Object = .EquippedObject(A).Object
                                Map(MapNum).Object(B).Value = .EquippedObject(A).Value
                                Map(MapNum).Object(B).TimeStamp = 1
                                St1 = St1 + DoubleChar(12) + Chr$(14) + Chr$(B) + DoubleChar$(CLng(.EquippedObject(A).Object)) + Chr$(.X) + Chr$(.Y) + Chr$(Map(MapNum).Object(B).ItemPrefix) + Chr$(Map(MapNum).Object(B).ItemSuffix) + QuadChar$(Map(MapNum).Object(B).Value)
                            End If
                        End If

                        .EquippedObject(A).Object = 0
                        .EquippedObject(A).Value = 0
                        .EquippedObject(A).ItemPrefix = 0
                        .EquippedObject(A).ItemSuffix = 0
                        St2 = St2 + DoubleChar(2) + Chr$(20) + Chr$(A + 20)
                        Exit For
                    End If
                End If
            Next A
        End If

        For A = 1 To 20
            Randomize
            If .Inv(A).Object > 0 Then
                Parameter(0) = Index
                Parameter(1) = .Inv(A).Value
                Randomize
                If (Rnd <= 0.3) And RunScript("DROPOBJ" + CStr(.Inv(A).Object)) = 0 Then
                    If Not ExamineBit(Object(.Inv(A).Object).flags, 2) = 255 Then
                        DontDropOnGround = False

                        If Killer > -1 And Not Killer = Index Then
                            Select Case Object(.Inv(A).Object).Type
                            Case 6, 11
                                D = FindInvObject(Killer, CLng(.Inv(A).Object))
                                If D = 0 Then D = FreeInvNum(Killer)
                            Case Else
                                D = FreeInvNum(Killer)
                            End Select

                            If D > 0 Then
                                Parameter(0) = Killer
                                Parameter(1) = .Inv(A).Value
                                If RunScript("GETOBJ" + CStr(.Inv(A).Object)) = 0 Then
                                    Select Case Object(.Inv(A).Object).Type
                                    Case 6, 11
                                        If Player(Killer).Inv(D).Object > 0 Then
                                            Player(Killer).Inv(D).Value = Player(Killer).Inv(D).Value + .Inv(A).Value
                                        Else
                                            Player(Killer).Inv(D).Object = .Inv(A).Object
                                            Player(Killer).Inv(D).Value = .Inv(A).Value
                                            Player(Killer).Inv(D).ItemPrefix = .Inv(A).ItemPrefix
                                            Player(Killer).Inv(D).ItemSuffix = .Inv(A).ItemSuffix
                                        End If
                                    Case Else
                                        Player(Killer).Inv(D).Object = .Inv(A).Object
                                        Player(Killer).Inv(D).Value = .Inv(A).Value
                                        Player(Killer).Inv(D).ItemPrefix = .Inv(A).ItemPrefix
                                        Player(Killer).Inv(D).ItemSuffix = .Inv(A).ItemSuffix
                                    End Select
                                    SendSocket Killer, Chr$(17) + Chr$(D) + DoubleChar$(CLng(Player(Killer).Inv(D).Object)) + QuadChar(Player(Killer).Inv(D).Value) + Chr$(Player(Killer).Inv(D).ItemPrefix) + Chr$(Player(Killer).Inv(D).ItemSuffix)
                                    DontDropOnGround = True
                                End If
                            End If
                        End If

                        If DontDropOnGround = False Then
                            B = FreeMapObj(MapNum)
                            If B >= 0 Then
                                Map(MapNum).Object(B).X = .X
                                Map(MapNum).Object(B).Y = .Y
                                Map(MapNum).Object(B).ItemPrefix = .Inv(A).ItemPrefix
                                Map(MapNum).Object(B).ItemSuffix = .Inv(A).ItemSuffix
                                Map(MapNum).Object(B).Object = .Inv(A).Object
                                Map(MapNum).Object(B).Value = .Inv(A).Value
                                Map(MapNum).Object(B).TimeStamp = 1
                                St1 = St1 + DoubleChar(12) + Chr$(14) + Chr$(B) + DoubleChar$(CLng(.Inv(A).Object)) + Chr$(.X) + Chr$(.Y) + Chr$(Map(MapNum).Object(B).ItemPrefix) + Chr$(Map(MapNum).Object(B).ItemSuffix) + QuadChar$(Map(MapNum).Object(B).Value)
                            End If
                        End If

                        .Inv(A).Object = 0
                        .Inv(A).Value = 0
                        .Inv(A).ItemPrefix = 0
                        .Inv(A).ItemSuffix = 0
                        St2 = St2 + DoubleChar(2) + Chr$(18) + Chr$(A)
                    End If
                End If
            End If
        Next A

        If St1 <> vbNullString Then
            SendToMapRaw MapNum, St1
        End If

        If St2 <> vbNullString Then
            SendRaw Index, St2
        End If

        '.Experience = Int((2 / 3) * .Experience)
        'SendSocket Index, Chr$(60) + QuadChar(.Experience)

        CalculateStats Index

        PlayerDied = True
    End With
End Function

Sub ResetMap(MapNum As Long)
    Dim A As Long, X As Long, Y As Long
    Dim NumPlayers As Long
    Dim St1 As String

    With Map(MapNum)
        NumPlayers = .NumPlayers
        For A = 0 To MaxMapObjects
            With .Object(A)
                If .Object > 0 Then
                    If Map(MapNum).Tile(.X, .Y).Att <> 5 And Map(MapNum).Tile(.X, .Y).Att2 <> 5 And Not .TimeStamp = 1 Then
                        .Object = 0
                        .Value = 0
                        .ItemPrefix = 0
                        .ItemSuffix = 0
                        If NumPlayers > 0 Then
                            St1 = St1 + DoubleChar(2) + Chr$(15) + Chr$(A)
                        End If
                    End If
                End If
            End With
        Next A
        For A = 0 To 9
            With .Door(A)
                If .Att > 0 Then
                    Map(MapNum).Tile(.X, .Y).Att = .Att
                    If NumPlayers > 0 Then
                        St1 = St1 + DoubleChar(2) + Chr$(37) + Chr$(A)
                    End If
                    .Att = 0
                End If
            End With
        Next A
        If ExamineBit(.flags, 3) = True Then
            'Create Monsters
            For A = 0 To MaxMonsters
                St1 = St1 + NewMapMonster(MapNum, A)
            Next A
        Else
            'Clear Monsters
            For A = 0 To MaxMonsters
                If .Monster(A).Monster > 0 Then
                    .Monster(A).Monster = 0
                    If NumPlayers > 0 Then
                        St1 = St1 + DoubleChar(2) + Chr$(39) + Chr$(A)
                    End If
                End If
            Next A
        End If
        If NumPlayers > 0 Then
            SendToMapRaw MapNum, St1
        End If
        For Y = 0 To 11
            For X = 0 To 11
                With Map(MapNum).Tile(X, Y)
                    If .Att = 7 Then
                        NewMapObject MapNum, CLng(.AttData(1)) * 256 + CLng(.AttData(0)), CLng(.AttData(2)) * 256& + CLng(.AttData(3)), X, Y, True
                    End If
                End With
            Next X
        Next Y
        .ResetTimer = 0
    End With
End Sub
Sub SendCharacterData(Index As Long)
    Dim St As String, A As Long
    With Player(Index)
        If .Class > 0 Then
            SendSocket Index, Chr$(3) + Chr$(.Class) + Chr$(.Gender) + DoubleChar$(CLng(.Sprite)) + Chr$(.Level) + Chr$(.Status) + Chr$(.Guild) + Chr$(.GuildRank) + Chr$(.Access) + Chr$(Index) + QuadChar(.Experience) + .Name + Chr$(0) + .desc
            For A = 1 To 10    'Send Skills
                St = St + DoubleChar(8) + Chr$(119) + Chr$(3) + Chr$(A) + Chr$(.Skill(A).Level) + QuadChar$(.Skill(A).Experience)
            Next A
            For A = 1 To MaxMagic    'Send Magic
                If .Level >= Magic(A).Level Then
                    If ExamineBit(Magic(A).Class, .Class - 1) = True Then
                        St = St + DoubleChar(9) + Chr$(153) + Chr$(3) + DoubleChar$(A) + Chr$(.MagicLevel(A).Level) + QuadChar$(.MagicLevel(A).Experience)
                    End If
                End If
            Next A
            SendRaw Index, St
        Else
            SendSocket Index, Chr$(3)
        End If
    End With
End Sub

Sub SendDataPacket(Index As Long, StartNum As Long)
    Dim A As Long, St1 As String

    For A = StartNum To 255
        If Guild(A).Name <> vbNullString Then
            With Guild(A)
                St1 = St1 + DoubleChar(3 + Len(.Name)) + Chr$(70) + Chr$(A) + Chr$(.MemberCount) + .Name
            End With
        End If
        If Len(St1) >= 700 Then
            If A < 255 Then
                St1 = St1 + DoubleChar(3) + Chr$(35) + Chr$(24) + Chr$(A + 1)
            Else
                St1 = St1 + DoubleChar(2) + Chr$(35) + Chr$(23)
            End If
            SendRaw Index, St1
            Exit Sub
        End If
    Next A
    St1 = St1 + DoubleChar(2) + Chr$(35) + Chr$(23)
    SendRaw Index, St1
End Sub

Function SpawnMapMonster(MapNum As Long, MonsterNum As Long, MonsterType As Long, TX As Long, TY As Long)
    With Map(MapNum).Monster(MonsterNum)
        .Monster = MonsterType
        .X = TX
        .Y = TY
        .HP = Monster(.Monster).HP
        .Target = 0
        .TargetIsMonster = False
        .MoveTimer = 0
        .AttackTimer = 0
        .D = Int(Rnd * 4)
        SpawnMapMonster = DoubleChar(9) + Chr$(38) + Chr$(MonsterNum) + DoubleChar$(CLng(.Monster)) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + DoubleChar$(CLng(.HP))
    End With
End Function

Function ValidName(St As String) As Boolean
    Dim A As Long, B As Long
    If Len(St) > 0 Then
        For A = 1 To Len(St)
            B = Asc(Mid$(St, A, 1))
            If (B < 48 Or B > 57) And (B < 65 Or B > 90) And (B < 97 Or B > 122) And B <> 32 And B <> 95 Then
                ValidName = False
                Exit Function
            End If
        Next
    End If
    ValidName = True
End Function
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim TempSocket As Long

    If uMsg >= 1029 And uMsg < 1029 + MaxUsers Then
        Dim Index As Long
        Index = uMsg - 1028
        Select Case lParam And 255
        Case FD_CLOSE
            AddSocketQue Index
        Case FD_READ
            ReadClientData Index
        End Select
    End If
    Select Case uMsg
    Case 1025    'Listening Socket
        Select Case lParam And 255
        Case FD_ACCEPT
            If lParam = FD_ACCEPT Then
                Dim NewPlayer As Long, Address As sockaddr
                Dim ClientIP As String

                NewPlayer = FreePlayer()
                If NewPlayer > 0 Then
                    With Player(NewPlayer)
                        .Socket = accept(ListeningSocket, Address, sockaddr_size)
                        If Not .Socket = INVALID_SOCKET Then
                            SetSockLinger .Socket, 1, 0
                            setsockopt .Socket, IPPROTO_TCP, TCP_NODELAY, 1&, 4
                            'setsockopt .Socket, SOL_SOCKET, SO_RCVBUF, 8192&, 4
                            'setsockopt .Socket, SOL_SOCKET, SO_SNDBUF, 8192&, 4
                            ClientIP = GetPeerAddress(.Socket)
                            If WSAAsyncSelect(.Socket, gHW, ByVal 1028 + NewPlayer, ByVal FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE) = 0 Then
                                .InUse = True
                                .Mode = modeNotConnected
                                .IP = ClientIP
                                .Class = 0
                                .SocketData = vbNullString
                                .LastMsg = getTime() - 50
                                .ClientVer = vbNullString
                                .FloodTimer = .LastMsg + 50
                                .PacketOrder = 0
                                .ServerPacketOrder = 0
                                PrintLog ("Connection accepted from " + .IP)
                                NumUsers = NumUsers + 1
                                frmMain.mnuDatabase.Enabled = False
                                frmMain.Caption = TitleString + " [" + CStr(NumUsers) + "]"
                            Else
                                closesocket .Socket
                                .Socket = INVALID_SOCKET
                            End If
                        Else
                            closesocket .Socket
                            .Socket = INVALID_SOCKET
                        End If
                    End With
                Else
                    TempSocket = accept(ListeningSocket, Address, sockaddr_size)
                    'SendData TempSocket, DoubleChar(2) + Chr$(0) + Chr$(4)
                    closesocket TempSocket
                    TempSocket = INVALID_SOCKET
                End If
            End If
        End Select
    End Select
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
Sub ParseString(St1)
    Dim St As String, A As Long, B As Long, C As Long
    If Mid$(St1, Len(St1), 1) = Chr$(13) Or Mid$(St1, Len(St1), 1) = Chr$(10) Then
        St1 = Mid$(St1, 1, Len(St1) - 1)
    End If
    If Mid$(St1, Len(St1), 1) = Chr$(13) Or Mid$(St1, Len(St1), 1) = Chr$(10) Then
        St1 = Mid$(St1, 1, Len(St1) - 1)
    End If
    For A = 1 To Len(St1)
        If Asc(Mid$(St1, 1, 1)) < 32 Then
            St1 = Mid$(St1, 2)
        Else
            Exit For
        End If
    Next A
    St = St1
    Suffix = vbNullString
    Prefix = vbNullString
    If Mid$(St, 1, 1) = ":" Then
        A = InStr(St, " ")
        Prefix = Mid$(St, 2, A - 2)
        St = Mid$(St, A + 1)
    End If
    St1 = St
    A = InStr(St, ":")
    If A > 0 Then
        Suffix = Mid$(St, A + 1, Len(St) - A)
        St = Mid$(St, 1, A - 1)
    End If
    B = 1
    Erase Word
    For A = 1 To 10
TryAgain9:
        C = InStr(B, St, " ")
        If C - B = 0 Then B = B + 1: GoTo TryAgain9
        If C <> 0 Then
            Word(A) = Mid$(St, B, C - B)
        Else
            Word(A) = Mid$(St, B, Len(St) - B + 1)
            Exit For
        End If
        B = C + 1
    Next A
End Sub
Function ExamineBit(bytByte As Byte, Bit As Byte) As Byte
    ExamineBit = ((bytByte And (2 ^ Bit)) > 0)
End Function
Sub SetBit(bytByte As Byte, Bit As Byte)
    bytByte = bytByte Or (2 ^ Bit)
End Sub
Sub ClearBit(bytByte As Byte, Bit As Byte)
    bytByte = bytByte And Not (2 ^ Bit)
End Sub
Sub CloseClientSocket(Index As Long)
    Dim A As Long
    With Player(Index)
        If .InUse = True Then
            'Decrement User Num
            NumUsers = NumUsers - 1
            If NumUsers = 0 Then frmMain.mnuDatabase.Enabled = True
            frmMain.Caption = TitleString + " [" + CStr(NumUsers) + "]"

            For A = 1 To MaxPlayerTimers
                If .ScriptTimer(A) > 0 Then
                    Parameter(0) = Index
                    .ScriptTimer(A) = 0
                    RunScript .Script(A)
                End If
            Next A

            If .Mode = modePlaying Then
                Parameter(0) = Index
                RunScript "PARTGAME"
            End If

            'Close Socket
            If Not .Socket = INVALID_SOCKET Then
                closesocket .Socket
                .Socket = INVALID_SOCKET
            End If

            If Not .Class = 0 Then
                If .Status = 2 Then .Status = 0
                If .IsDead = True Then
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
                        A = Int(Rnd * 2)

                        .Map = World.StartLocation(A).Map
                        .X = World.StartLocation(A).X
                        .Y = World.StartLocation(A).Y
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
                End If
                SavePlayerData Index
            End If

            PrintLog "Connection closed from " + .IP + " [" + Player(Index).Name + "]"

            If .Mode = modePlaying Then SendToGods Chr$(56) + Chr$(7) + .User + " - " + .IP
            If .Access > 0 And .Mode = modePlaying Then PrintGodSilent .User, " (Left Game) "

            'Send Quit Message
            If .Mode = modePlaying Then
                .Mode = modeNotConnected
                SendAll Chr$(7) + Chr$(Index)
                If .Map > 0 Then
                    Partmap Index
                    .Map = 0
                End If
            Else
                .Mode = modeNotConnected
            End If

            'Clear Socket Data
            .InUse = False
            .SocketData = vbNullString
            .Class = 0
            .MaxHP = 0
            .MaxMana = 0
            .MaxEnergy = 0
            .User = vbNullString
            .Name = vbNullString
            .IsDead = False
            .ComputerID = vbNullString
            For A = 0 To 29
                .ItemBank(A).Object = 0
                .ItemBank(A).Value = 0
                .ItemBank(A).ItemPrefix = 0
                .ItemBank(A).ItemSuffix = 0
            Next A
            For A = 1 To MaxSkill
                .Skill(A).Experience = 0
                .Skill(A).Level = 0
            Next A
            For A = 1 To MaxMagic
                .MagicLevel(A).Experience = 0
                .MagicLevel(A).Level = 0
            Next A
            .Bank = 0
            
            .CurPing = 0
            For A = 1 To 5
                .Ping(A) = 0
            Next A
            .LastPing = 0

            For A = 1 To MaxUsers
                If CloseSocketQue(A) = Index Then
                    CloseSocketQue(A) = 0
                End If
            Next A
        End If
    End With
End Sub
Function DoubleChar(Num As Long) As String
    DoubleChar = Chr$(Int(Num / 256)) + Chr$(Num Mod 256)
End Function
Function TripleChar(Num As Long) As String
    TripleChar = Chr$(Int(Num / 65536)) + Chr$(Int((Num Mod 65536) / 256)) + Chr$(Num Mod 256)
End Function
Function QuadChar(Num As Long) As String
    If Num < 0 Then
        SendToGods Chr$(56) + Chr$(7) + "WARNING:  QuadChar less than 0: " + CStr(Num)
        PrintLog "WARNING:  QuadChar less than 0    " + CStr(Num)
        PrintDebug "WARNING:  QuadChar less than 0   " + CStr(Num)
        QuadChar = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
    Else
        QuadChar = Chr$(Int(Num / 16777216) Mod 256) + Chr$(Int(Num / 65536) Mod 256) + Chr$(Int(Num / 256) Mod 256) + Chr$(Num Mod 256)
    End If
End Function
Function Exists(Filename As String) As Boolean
    Exists = (Dir(Filename) <> vbNullString)
End Function
Function GetInt(Chars As String) As Long
    GetInt = CLng(Asc(Mid$(Chars, 1, 1))) * 256& + CLng(Asc(Mid$(Chars, 2, 1)))
End Function
Sub GetWords(St As String)
    Dim A As Long, B As Long, C As Long
    B = 1
    Erase Word
    For A = 1 To 50
TryAgain:
        C = InStr(B, St, " ")
        If C - B = 0 Then B = B + 1: GoTo TryAgain
        If C <> 0 Then
            Word(A) = Mid$(St, B, C - B)
        Else
            Word(A) = Mid$(St, B, Len(St) - B + 1)
            Exit For
        End If
        B = C + 1
    Next A
End Sub
Sub GetSections(St)
    Dim A As Long, B As Long, C As Long
    B = 1
    Erase Word
    For A = 1 To 10
        C = InStr(B, St, vbNullChar)
        If C - B = 0 Then
            Word(A) = vbNullString
        ElseIf C <> 0 Then
            Word(A) = Mid$(St, B, C - B)
        Else
            Word(A) = Mid$(St, B, Len(St) - B + 1)
            Exit For
        End If
        B = C + 1
    Next A
End Sub
Function Nick(UserHost As String) As String
    Dim A As Long

    A = InStr(UserHost, "!")
    If A > 0 Then
        Nick = Mid$(UserHost, 1, A - 1)
    Else
        Nick = UserHost
    End If
End Function
Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub Unhook()
    SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
End Sub
Sub SavePlayerData(Index)
    Dim A As Long, St As String

    With Player(Index)
        If .LastSkillUse = 69 Then Exit Sub

        UserRS.Index = "User"
        UserRS.Seek "=", .User
        UserRS.Edit
        
        If .User = vbNullString Then
            PrintCheat ("Null Save Player Data found.")
            Exit Sub
        End If

        .Bookmark = UserRS.Bookmark
        UserRS!Access = .Access

        'Character Data
        UserRS!Name = .Name
        UserRS!Class = .Class
        UserRS!Gender = .Gender
        UserRS!Sprite = .Sprite
        UserRS!desc = .desc
        UserRS!Email = .Email

        'Position Data
        UserRS!Map = .Map
        UserRS!X = .X
        UserRS!Y = .Y
        UserRS!D = .D

        'Character Physical Stats
        UserRS!Level = .Level
        UserRS!Experience = .Experience

        'Misc. Data
        UserRS!Bank = .Bank
        UserRS!Status = .Status
        UserRS!LastPlayed = CLng(Date)

        'Inventory Data
        For A = 1 To 20
            UserRS.Fields("InvObject" + CStr(A)).Value = .Inv(A).Object
            UserRS.Fields("InvValue" + CStr(A)).Value = .Inv(A).Value
            If .Inv(A).Object > 0 Then
                If .Inv(A).ItemPrefix > 0 Then
                    If Len(ItemPrefix(.Inv(A).ItemPrefix).Name) = 0 Then .Inv(A).ItemPrefix = 0
                End If
                If .Inv(A).ItemSuffix > 0 Then
                    If Len(ItemSuffix(.Inv(A).ItemSuffix).Name) = 0 Then .Inv(A).ItemSuffix = 0
                End If
            End If
            UserRS.Fields("InvPrefix" + CStr(A)).Value = .Inv(A).ItemPrefix
            UserRS.Fields("InvSuffix" + CStr(A)).Value = .Inv(A).ItemSuffix
        Next A

        'Equipped Objects
        For A = 1 To 6
            UserRS.Fields("EquippedObject" + CStr(A)).Value = .EquippedObject(A).Object
            UserRS.Fields("EquippedVal" + CStr(A)).Value = .EquippedObject(A).Value
            If .EquippedObject(A).Object > 0 Then
                If .EquippedObject(A).ItemPrefix > 0 Then
                    If Len(ItemPrefix(.EquippedObject(A).ItemPrefix).Name) = 0 Then .EquippedObject(A).ItemPrefix = 0
                End If
                If .EquippedObject(A).ItemSuffix > 0 Then
                    If Len(ItemSuffix(.EquippedObject(A).ItemSuffix).Name) = 0 Then .EquippedObject(A).ItemSuffix = 0
                End If
            End If
            UserRS.Fields("EquippedPrefix" + CStr(A)).Value = .EquippedObject(A).ItemPrefix
            UserRS.Fields("EquippedSuffix" + CStr(A)).Value = .EquippedObject(A).ItemSuffix
        Next A

        'Item Bank
        For A = 0 To 29
            UserRS.Fields("BankObject" + CStr(A)).Value = .ItemBank(A).Object
            UserRS.Fields("BankValue" + CStr(A)).Value = .ItemBank(A).Value
            If .ItemBank(A).Object > 0 Then
                If .ItemBank(A).ItemPrefix > 0 Then
                    If Len(ItemPrefix(.ItemBank(A).ItemPrefix).Name) = 0 Then .ItemBank(A).ItemPrefix = 0
                End If
                If .ItemBank(A).ItemSuffix > 0 Then
                    If Len(ItemSuffix(.ItemBank(A).ItemSuffix).Name) = 0 Then .ItemBank(A).ItemSuffix = 0
                End If
            End If
            UserRS.Fields("BankPrefix" + CStr(A)).Value = .ItemBank(A).ItemPrefix
            UserRS.Fields("BankSuffix" + CStr(A)).Value = .ItemBank(A).ItemSuffix
            
        Next A

        'Flags
        St = vbNullString
        For A = 0 To MaxPlayerFlags
            If .Flag(A) > 0 Then
                St = St + DoubleChar$(A) + QuadChar$(.Flag(A))
            End If
        Next A
        UserRS!flags = St

        St = vbNullString
        For A = 1 To MaxSkill
            With .Skill(A)
                If .Experience < 0 Then .Experience = 0
                St = St + Chr$(.Level) + QuadChar$(.Experience)
            End With
        Next A
        If Len(St) > 0 Then UserRS!Skills = St
        
        St = vbNullString
        For A = 1 To MaxMagic
            With .MagicLevel(A)
                If .Experience < 0 Then .Experience = 0
                St = St + Chr$(.Level) + QuadChar$(.Experience)
            End With
        Next A
        If Len(St) > 0 Then UserRS!Magic = St

        UserRS.Update
    End With
End Sub
Sub SendAll(ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub
Sub SendToConnected(ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode > 0 Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub
Sub SendAllBut(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And A <> Index Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub
Sub SendAllButRaw(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And A <> Index Then
                SendRaw A, St
            End If
        End With
    Next A
End Sub
Sub SendAllButBut(ByVal Index1 As Long, ByVal Index2 As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And A <> Index1 And A <> Index2 Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub

Sub SendToGods(ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Access > 0 Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub

Sub SendToAdmins(ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Access > 2 Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub
Sub SendToGodsAllBut(Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Access > 0 And Index <> A Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub

Sub SendToMap(ByVal MapNum As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Map = MapNum Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub
Sub SendToMapRaw(ByVal MapNum As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Map = MapNum Then
                SendRaw A, St
            End If
        End With
    Next A
End Sub
Sub ShutdownServer()
    Dim A As Long, B As Long
    For A = 1 To MaxUsers
        If Player(A).InUse = True Then
            CloseClientSocket A
        End If
    Next A
    
    If ListeningSocket <> INVALID_SOCKET Then
        closesocket ListeningSocket
        ListeningSocket = INVALID_SOCKET
    End If
    EndWinsock
    Unhook
    
    frmMain.PlayerTimer.Enabled = False
    frmMain.MapTimer.Enabled = False
    frmMain.MinuteTimer.Enabled = False

    For A = 1 To MaxGuilds
        If Not Guild(A).Name = "" Then
            If Guild(A).UpdateFlag = True Then
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
    SaveBugReports

    UserRS.Close
    GuildRS.Close
    NPCRS.Close
    MonsterRS.Close
    ObjectRS.Close
    DataRS.Close
    MapRS.Close
    BanRS.Close
    PrefixRS.Close
    SuffixRS.Close
    MagicRS.Close
    DB.Close
    WS.Close

End Sub
Sub SendToMapAllBut(ByVal MapNum As Long, ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Map = MapNum And Index <> A Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub
Sub SendToMapAllButRaw(ByVal MapNum As Long, ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Map = MapNum And Index <> A Then
                SendRaw A, St
            End If
        End With
    Next A
End Sub
Sub SendSocket(ByVal Index As Long, ByVal St As String)
    If Index > 0 Then
        With Player(Index)
            If .InUse = True Then
                If SendData(.Socket, DoubleChar$(Len(St)) + Chr$(CheckSum(St) * 20 Mod 194) + Chr$(.ServerPacketOrder) + St) = SOCKET_ERROR Then
    
                End If
                .ServerPacketOrder = .ServerPacketOrder + 1
                If .ServerPacketOrder > 250 Then .ServerPacketOrder = 0
            End If
        End With
    End If
End Sub
Function GetSendSocket(ByVal Index As Long, ByVal St As String) As String
    Dim SendSt As String
    With Player(Index)
        If .InUse = True Then
            SendSt = DoubleChar$(Len(St)) + Chr$(CheckSum(St) * 20 Mod 194) + Chr$(.ServerPacketOrder) + St
            .ServerPacketOrder = .ServerPacketOrder + 1
            If .ServerPacketOrder > 250 Then .ServerPacketOrder = 0
            GetSendSocket = SendSt
        End If
    End With
End Function
Sub SendRaw(ByVal Index As Long, ByVal St As String)
    With Player(Index)
        If .InUse = True Then
            SendSocket Index, Chr$(170) + St
        End If
    End With
End Sub
Sub SendRawReal(ByVal Index As Long, ByVal St As String)
    With Player(Index)
        If .InUse = True Then
            SendData .Socket, St
        End If
    End With
End Sub
Sub PrintLog(St)
    With frmMain.lstLog
        .AddItem St
        If .ListCount > 30 Then .RemoveItem 0
        If .ListIndex = .ListCount - 2 Then .ListIndex = .ListCount - 1
    End With
End Sub

Function AddSocketQue(Index As Long) As Integer
    Dim A As Integer

    For A = 1 To MaxUsers
        If CloseSocketQue(A) = Index Then
            Exit Function
        End If
    Next A

    For A = 1 To MaxUsers
        If CloseSocketQue(A) = 0 Then
            CloseSocketQue(A) = Index
            Exit For
        End If
    Next A
End Function

Sub GiveStartingEQ(Index As Long)
    Dim A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers Then

        With Player(Index)
            For A = 1 To 8
                If World.StartObjects(A) > 0 Then
                    B = World.StartObjects(A)
                    C = World.StartObjValues(A)
                    .Inv(A).Object = B

                    Select Case Object(B).Type
                    Case 1, 2, 3, 4    'Weapon, Shield, Armor, Helmut
                        .Inv(A).Value = CLng(Object(B).Data(0)) * 10
                    Case 6    'Money
                        .Inv(A).Value = C
                    Case 8    'Ring
                        .Inv(A).Value = CLng(Object(B).Data(1)) * 10
                    Case Else
                        .Inv(A).Value = 0
                    End Select
                End If
            Next A
        End With
    End If
End Sub
Function GetRepairCost(Index As Long, Slot As Integer) As Long
    Dim A As Long, B As Long, C As Long

    If Slot = 0 Then Exit Function

    If Index >= 1 And Index <= MaxUsers Then
        If Slot >= 0 And Slot <= 20 Then
            Select Case Object(Player(Index).Inv(Slot).Object).Type
                Case 1, 2, 3, 4, 8    'Weapon, Shield, Armor, Helmet, Ring
                    A = Object(Player(Index).Inv(Slot).Object).Type
    
                    If ExamineBit(Object(Player(Index).Inv(Slot).Object).flags, 0) Or ExamineBit(Object(Player(Index).Inv(Slot).Object).flags, 1) Or Object(Player(Index).Inv(Slot).Object).SellPrice = 0 Then
                        A = 0
                    End If
                Case Else
                    A = 0
            End Select

            If A > 0 Then
                Select Case A
                    Case 1, 2, 3, 4   'Weapon, Shield, Armor, Helmet
                        'C = Object(Player(Index).Inv(Slot).Object).Data(0) * 10 - (Player(Index).Inv(Slot).Value)
                        'B = B + (C * World.Cost_Per_Durability)
                        'B = B + (C * (Object(Player(Index).Inv(Slot).Object).Data(1) * World.Cost_Per_Strength))
                        'If B > 0 Then B = B / 100
                        If Object(Player(Index).Inv(Slot).Object).Data(0) * 10 > 0 Then
                            C = Object(Player(Index).Inv(Slot).Object).SellPrice - (((Player(Index).Inv(Slot).Value) / (Object(Player(Index).Inv(Slot).Object).Data(0) * 10)) * Object(Player(Index).Inv(Slot).Object).SellPrice)
                            If C >= 0 Then
                                GetRepairCost = C
                            Else
                                GetRepairCost = 0
                            End If
                        Else
                            GetRepairCost = 0
                        End If
                        Exit Function
                    Case 8 'Ring
                        If Object(Player(Index).Inv(Slot).Object).Data(1) * 10 > 0 Then
                            C = Object(Player(Index).Inv(Slot).Object).SellPrice - (((Player(Index).Inv(Slot).Value) / (Object(Player(Index).Inv(Slot).Object).Data(1) * 10)) * Object(Player(Index).Inv(Slot).Object).SellPrice)
                            If C >= 0 Then
                                GetRepairCost = C
                            Else
                                GetRepairCost = 0
                            End If
                        Else
                            GetRepairCost = 0
                        End If
                        Exit Function
                    Case Else
                        GetRepairCost = 0
                End Select
            Else
                GetRepairCost = 0
            End If
        Else
            Select Case Object(Player(Index).EquippedObject(Slot - 20).Object).Type
            Case 1, 2, 3, 4, 8    'Weapon, Shield, Armor, Helmet, Ring
                A = Object(Player(Index).EquippedObject(Slot - 20).Object).Type

                If ExamineBit(Object(Player(Index).EquippedObject(Slot - 20).Object).flags, 0) Or ExamineBit(Object(Player(Index).EquippedObject(Slot - 20).Object).flags, 1) Or Object(Player(Index).EquippedObject(Slot - 20).Object).SellPrice = 0 Then
                    A = 0
                End If
            Case Else
                A = 0
            End Select

            If A > 0 Then
                Select Case A
                    Case 1, 2, 3, 4    'Weapon, Shield, Armor, Helmet, Ring
                        If Object(Player(Index).EquippedObject(Slot - 20).Object).Data(0) * 10 > 0 Then
                            C = Object(Player(Index).EquippedObject(Slot - 20).Object).SellPrice - ((Player(Index).EquippedObject(Slot - 20).Value / (Object(Player(Index).EquippedObject(Slot - 20).Object).Data(0) * 10)) * Object(Player(Index).EquippedObject(Slot - 20).Object).SellPrice)
                            If C >= 0 Then
                                GetRepairCost = C
                            Else
                                GetRepairCost = 0
                            End If
                        Else
                            GetRepairCost = 0
                        End If
                        Exit Function
                    Case 8 'Ring
                        If Object(Player(Index).EquippedObject(Slot - 20).Object).Data(1) * 10 > 0 Then
                            C = Object(Player(Index).EquippedObject(Slot - 20).Object).SellPrice - ((Player(Index).EquippedObject(Slot - 20).Value / (Object(Player(Index).EquippedObject(Slot - 20).Object).Data(1) * 10)) * Object(Player(Index).EquippedObject(Slot - 20).Object).SellPrice)
                            If C >= 0 Then
                                GetRepairCost = C
                            Else
                                GetRepairCost = 0
                            End If
                        Else
                            GetRepairCost = 0
                        End If
                        Exit Function
                End Select
            Else
                GetRepairCost = 0
            End If
        End If
    End If
End Function
Function GetObjectDur(ByVal Index As Long, ByVal Slot As Long) As Long
    Dim Percent As Single
    Select Case Object(Player(Index).Inv(Slot).Object).Type
    Case 1, 2, 3, 4
        Percent = Player(Index).Inv(Slot).Value / (Object(Player(Index).Inv(Slot).Object).Data(0) * 10)
        Percent = Percent * 100
        If Percent > 100 Then Percent = 100
        GetObjectDur = Percent
    Case 8
        Percent = Player(Index).Inv(Slot).Value / (Object(Player(Index).Inv(Slot).Object).Data(1) * 10)
        Percent = Percent * 100
        If Percent > 100 Then Percent = 100
        GetObjectDur = Percent
    Case Else
        GetObjectDur = 0
    End Select
End Function

Sub CalculateStats(Index As Long)
    Dim RunningTotal As Integer, MagicTotal As Integer, AttackTotal As Integer, A As Long
    Dim TotalMaxHP As Long, TotalMaxEnergy As Long, TotalMaxMana As Long
    Dim OldMaxHP As Integer, OldMaxEnergy As Integer, OldMaxMana As Integer
    Dim OldAttack As Integer, OldDefense As Integer, OldMagicDefense As Integer

    If Index > 0 Then
        With Player(Index)

            OldMaxHP = .MaxHP
            OldMaxEnergy = .MaxEnergy
            OldMaxMana = .MaxMana

            OldAttack = .PhysicalAttack
            OldDefense = .TotalDefense
            OldMagicDefense = .MagicDefense

            ''''Attack/Defense
            If .EquippedObject(1).Object > 0 Then
                If Object(.EquippedObject(1).Object).Type = 1 Then
                    AttackTotal = AttackTotal + Object(.EquippedObject(1).Object).Data(1)
                End If
                If .EquippedObject(1).ItemPrefix > 0 Then
                    Select Case ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                    End Select
                End If
                If .EquippedObject(1).ItemSuffix > 0 Then
                    Select Case ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                    End Select
                End If
            End If

            If .EquippedObject(2).Object > 0 Then    ' Shield
                If .EquippedObject(2).ItemPrefix > 0 Then
                    Select Case ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    End Select
                End If
                If .EquippedObject(2).ItemSuffix > 0 Then
                    Select Case ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    End Select
                End If
            End If

            For A = 3 To 4
                If .EquippedObject(A).Object > 0 Then
                    If .EquippedObject(A).ItemPrefix > 0 Then
                        Select Case ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationType
                        Case 8    'Max HP
                            TotalMaxHP = TotalMaxHP + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 9    'Max Energy
                            TotalMaxEnergy = TotalMaxEnergy + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 10    'Max Mana
                            TotalMaxMana = TotalMaxMana + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 11    'Damage
                            AttackTotal = AttackTotal + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 12    'Defense
                            RunningTotal = RunningTotal + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 13    'Magic Defense
                            MagicTotal = MagicTotal + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        End Select
                    End If
                    If .EquippedObject(A).ItemSuffix > 0 Then
                        Select Case ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationType
                        Case 8    'Max HP
                            TotalMaxHP = TotalMaxHP + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 9    'Max Energy
                            TotalMaxEnergy = TotalMaxEnergy + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 10    'Max Mana
                            TotalMaxMana = TotalMaxMana + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 11    'Damage
                            AttackTotal = AttackTotal + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 12    'Defense
                            RunningTotal = RunningTotal + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 13    'Magic Defense
                            MagicTotal = MagicTotal + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        End Select
                    End If
                    RunningTotal = RunningTotal + Object(.EquippedObject(A).Object).Data(1)
                    MagicTotal = MagicTotal + Object(.EquippedObject(A).Object).Data(2)
                End If
            Next A
            If .EquippedObject(5).Object > 0 Then    'Ring
                If .EquippedObject(5).ItemPrefix > 0 Then
                    Select Case ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    End Select
                End If
                If .EquippedObject(5).ItemSuffix > 0 Then
                    Select Case ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    End Select
                End If
                If Object(.EquippedObject(5).Object).Data(0) = 0 Then    'Attack
                    AttackTotal = AttackTotal + Object(.EquippedObject(5).Object).Data(2)
                End If
                If Object(.EquippedObject(5).Object).Data(0) = 1 Then    'Defense
                    RunningTotal = RunningTotal + Object(.EquippedObject(5).Object).Data(2)
                End If
                If Object(.EquippedObject(5).Object).Data(0) = 2 Then    'Magic Defense
                    MagicTotal = MagicTotal + Object(.EquippedObject(5).Object).Data(2)
                End If
            End If

            If RunningTotal > 255 Then RunningTotal = 255
            If MagicTotal > 255 Then MagicTotal = 255

            .TotalDefense = RunningTotal
            .MagicDefense = MagicTotal
            '''


            ''''HP/Mana
            Dim TempVar As Double

            'HP
            TempVar = Class(.Class).StartHP + CInt(CDbl(Class(.Class).MaxHP - Class(.Class).StartHP) * (CDbl(.Level) / CDbl(World.MaxLevel)))
            TotalMaxHP = TotalMaxHP + TempVar + CInt(World.StatConstitution)

            'Energy
            TempVar = Class(.Class).StartEnergy + CInt(CDbl(Class(.Class).MaxEnergy - Class(.Class).StartEnergy) * (CDbl(.Level) / CDbl(World.MaxLevel)))
            TotalMaxEnergy = TotalMaxEnergy + TempVar + CInt(World.StatStamina)

            'Mana
            TempVar = Class(.Class).StartMana + CInt(CDbl(Class(.Class).MaxMana - Class(.Class).StartMana) * (CDbl(.Level) / CDbl(World.MaxLevel)))
            TotalMaxMana = TotalMaxMana + TempVar + CInt(World.StatWisdom)
            '''

            If TotalMaxHP > 255 Then TotalMaxHP = 255
            If TotalMaxEnergy > 255 Then TotalMaxEnergy = 255
            If TotalMaxMana > 255 Then TotalMaxMana = 255

            ''Final Stats
            .MaxHP = TotalMaxHP
            .MaxEnergy = TotalMaxEnergy
            .MaxMana = TotalMaxMana

            AttackTotal = AttackTotal + (World.StatStrength)
            If AttackTotal > 255 Then AttackTotal = 255

            .PhysicalAttack = AttackTotal

            If .HP > .MaxHP Then
                .HP = .MaxHP
                SendSocket Index, Chr$(46) + Chr$(.HP)
            End If
            If .Energy > .MaxEnergy Then
                .Energy = .MaxEnergy
            End If
            If .Mana > .MaxMana Then
                .Mana = .MaxMana
                SendSocket Index, Chr$(48) + Chr$(.Mana)
            End If

            Dim StatsChanged As Boolean
            StatsChanged = False

            If Not OldMaxHP = .MaxHP Then StatsChanged = True
            If Not OldMaxEnergy = .MaxEnergy Then StatsChanged = True
            If Not OldMaxMana = .MaxMana Then StatsChanged = True

            If Not OldAttack = .PhysicalAttack Then StatsChanged = True
            If Not OldDefense = .TotalDefense Then StatsChanged = True
            If Not OldMagicDefense = .MagicDefense Then StatsChanged = True

            If StatsChanged = True Then
                SendSocket Index, Chr$(130) + Chr$(.MaxHP) + Chr$(.MaxEnergy) + Chr$(.MaxMana) + Chr$(.PhysicalAttack) + Chr$(.TotalDefense) + Chr$(.MagicDefense)
            End If

        End With

        RunningTotal = 0
        MagicTotal = 0
        AttackTotal = 0

        With Player(Index)
            OldMaxHP = .MaxHP
            OldMaxEnergy = .MaxEnergy
            OldMaxMana = .MaxMana

            OldAttack = .PhysicalAttack
            OldDefense = .TotalDefense
            OldMagicDefense = .MagicDefense

            'Set stats from base
            TotalMaxHP = 0
            TotalMaxEnergy = 0
            TotalMaxMana = 0

            ''''Attack/Defense
            If .EquippedObject(1).Object > 0 Then
                If Object(.EquippedObject(1).Object).Type = 1 Or Object(.EquippedObject(1).Object).Type = 1 Then
                    AttackTotal = AttackTotal + Object(.EquippedObject(1).Object).Data(1)
                    If .EquippedObject(1).ItemPrefix > 0 Then
                        Select Case ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationType
                        Case 8    'Max HP
                            TotalMaxHP = TotalMaxHP + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                        Case 9    'Max Energy
                            TotalMaxEnergy = TotalMaxEnergy + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                        Case 10    'Max Mana
                            TotalMaxMana = TotalMaxMana + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                        Case 11    'Damage
                            AttackTotal = AttackTotal + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                        Case 12    'Defense
                            RunningTotal = RunningTotal + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                        Case 13    'Magic Defense
                            MagicTotal = MagicTotal + ItemPrefix(.EquippedObject(1).ItemPrefix).ModificationValue
                        End Select
                    End If
                    If .EquippedObject(1).ItemSuffix > 0 Then
                        Select Case ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationType
                        Case 8    'Max HP
                            TotalMaxHP = TotalMaxHP + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                        Case 9    'Max Energy
                            TotalMaxEnergy = TotalMaxEnergy + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                        Case 10    'Max Mana
                            TotalMaxMana = TotalMaxMana + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                        Case 11    'Damage
                            AttackTotal = AttackTotal + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                        Case 12    'Defense
                            RunningTotal = RunningTotal + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                        Case 13    'Magic Defense
                            MagicTotal = MagicTotal + ItemSuffix(.EquippedObject(1).ItemSuffix).ModificationValue
                        End Select
                    End If
                End If
            End If
            If .EquippedObject(2).Object > 0 Then    ' Shield
                If .EquippedObject(2).ItemPrefix > 0 Then
                    Select Case ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemPrefix(.EquippedObject(2).ItemPrefix).ModificationValue
                    End Select
                End If
                If .EquippedObject(2).ItemSuffix > 0 Then
                    Select Case ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemSuffix(.EquippedObject(2).ItemSuffix).ModificationValue
                    End Select
                End If
            End If
            For A = 3 To 4
                If .EquippedObject(A).Object > 0 Then
                    If .EquippedObject(A).ItemPrefix > 0 Then
                        Select Case ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationType
                        Case 8    'Max HP
                            TotalMaxHP = TotalMaxHP + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 9    'Max Energy
                            TotalMaxEnergy = TotalMaxEnergy + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 10    'Max Mana
                            TotalMaxMana = TotalMaxMana + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 11    'Damage
                            AttackTotal = AttackTotal + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 12    'Defense
                            RunningTotal = RunningTotal + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        Case 13    'Magic Defense
                            MagicTotal = MagicTotal + ItemPrefix(.EquippedObject(A).ItemPrefix).ModificationValue
                        End Select
                    End If
                    If .EquippedObject(A).ItemSuffix > 0 Then
                        Select Case ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationType
                        Case 8    'Max HP
                            TotalMaxHP = TotalMaxHP + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 9    'Max Energy
                            TotalMaxEnergy = TotalMaxEnergy + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 10    'Max Mana
                            TotalMaxMana = TotalMaxMana + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 11    'Damage
                            AttackTotal = AttackTotal + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 12    'Defense
                            RunningTotal = RunningTotal + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        Case 13    'Magic Defense
                            MagicTotal = MagicTotal + ItemSuffix(.EquippedObject(A).ItemSuffix).ModificationValue
                        End Select
                    End If
                    RunningTotal = RunningTotal + Object(.EquippedObject(A).Object).Data(1)
                    MagicTotal = MagicTotal + Object(.EquippedObject(A).Object).Data(2)
                End If
            Next A
            If .EquippedObject(5).Object > 0 Then    'Ring
                If .EquippedObject(5).ItemPrefix > 0 Then
                    Select Case ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemPrefix(.EquippedObject(5).ItemPrefix).ModificationValue
                    End Select
                End If
                If .EquippedObject(5).ItemSuffix > 0 Then
                    Select Case ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationType
                    Case 8    'Max HP
                        TotalMaxHP = TotalMaxHP + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 9    'Max Energy
                        TotalMaxEnergy = TotalMaxEnergy + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 10    'Max Mana
                        TotalMaxMana = TotalMaxMana + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 11    'Damage
                        AttackTotal = AttackTotal + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 12    'Defense
                        RunningTotal = RunningTotal + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    Case 13    'Magic Defense
                        MagicTotal = MagicTotal + ItemSuffix(.EquippedObject(5).ItemSuffix).ModificationValue
                    End Select
                End If
                If Object(.EquippedObject(5).Object).Data(0) = 0 Then    'Attack
                    AttackTotal = AttackTotal + Object(.EquippedObject(5).Object).Data(2)
                End If
                If Object(.EquippedObject(5).Object).Data(0) = 1 Then    'Defense
                    RunningTotal = RunningTotal + Object(.EquippedObject(5).Object).Data(2)
                End If
                If Object(.EquippedObject(5).Object).Data(0) = 2 Then    'Magic Defense
                    MagicTotal = MagicTotal + Object(.EquippedObject(5).Object).Data(2)
                End If
            End If

            If RunningTotal > 255 Then RunningTotal = 255
            If MagicTotal > 255 Then MagicTotal = 255

            .TotalDefense = RunningTotal
            .MagicDefense = MagicTotal
            '''


            'HP
            TempVar = Class(.Class).StartHP + CInt(CDbl(Class(.Class).MaxHP - Class(.Class).StartHP) * (CDbl(.Level) / CDbl(World.MaxLevel)))
            TotalMaxHP = TotalMaxHP + TempVar + CInt(World.StatConstitution)

            'Energy
            TempVar = Class(.Class).StartEnergy + CInt(CDbl(Class(.Class).MaxEnergy - Class(.Class).StartEnergy) * (CDbl(.Level) / CDbl(World.MaxLevel)))
            TotalMaxEnergy = TotalMaxEnergy + TempVar + CInt(World.StatStamina)

            'Mana
            TempVar = Class(.Class).StartMana + CInt(CDbl(Class(.Class).MaxMana - Class(.Class).StartMana) * (CDbl(.Level) / CDbl(World.MaxLevel)))
            TotalMaxMana = TotalMaxMana + TempVar + CInt(World.StatWisdom)
            '''

            If TotalMaxHP > 255 Then TotalMaxHP = 255
            If TotalMaxEnergy > 255 Then TotalMaxEnergy = 255
            If TotalMaxMana > 255 Then TotalMaxMana = 255

            ''Final Stats
            .MaxHP = TotalMaxHP
            .MaxEnergy = TotalMaxEnergy
            .MaxMana = TotalMaxMana

            'Parameter(0) = Index
            'RunScript "CALCULATESTATS"

            AttackTotal = AttackTotal + (World.StatStrength)
            If AttackTotal > 255 Then AttackTotal = 255

            .PhysicalAttack = AttackTotal

            If .HP > .MaxHP Then
                .HP = .MaxHP
                SendSocket Index, Chr$(46) + Chr$(.HP)
            End If
            If .Energy > .MaxEnergy Then
                .Energy = .MaxEnergy
            End If
            If .Mana > .MaxMana Then
                .Mana = .MaxMana
                SendSocket Index, Chr$(48) + Chr$(.Mana)
            End If

            StatsChanged = False

            If Not OldMaxHP = .MaxHP Then StatsChanged = True
            If Not OldMaxEnergy = .MaxEnergy Then StatsChanged = True
            If Not OldMaxMana = .MaxMana Then StatsChanged = True

            If Not OldAttack = .PhysicalAttack Then StatsChanged = True
            If Not OldDefense = .TotalDefense Then StatsChanged = True
            If Not OldMagicDefense = .MagicDefense Then StatsChanged = True

            If StatsChanged = True Then
                SendSocket Index, Chr$(130) + Chr$(.MaxHP) + Chr$(.MaxEnergy) + Chr$(.MaxMana) + Chr$(.PhysicalAttack) + Chr$(.TotalDefense) + Chr$(.MagicDefense)
            End If

        End With
    End If
End Sub

Function ClipString(St As String) As String
    Dim A As Long
    For A = Len(St) To 1 Step -1
        If Mid$(St, A, 1) <> Chr$(32) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function

Sub EquipObject(Index As Long, Slot As Long)
    Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long
    With Player(Index)
        If .Inv(Slot).Object > 0 Then    'Has object
            Select Case Object(.Inv(Slot).Object).Type
            Case 1, 10
                If Object(.Inv(Slot).Object).LevelReq <= .Level Then
                    A = .Inv(Slot).Object
                    B = .Inv(Slot).Value
                    E = .Inv(Slot).ItemPrefix
                    F = .Inv(Slot).ItemSuffix
                    .Inv(Slot).Object = 0
                    .Inv(Slot).Value = 0
                    .Inv(Slot).ItemPrefix = 0
                    .Inv(Slot).ItemSuffix = 0
                    If .EquippedObject(1).Object > 0 Then
                        C = FreeInvNum(Index)
                        .Inv(C).Object = .EquippedObject(1).Object
                        .Inv(C).Value = .EquippedObject(1).Value
                        .Inv(C).ItemPrefix = .EquippedObject(1).ItemPrefix
                        .Inv(C).ItemSuffix = .EquippedObject(1).ItemSuffix
                        .EquippedObject(1).Object = 0
                        .EquippedObject(1).Value = 0
                        .EquippedObject(1).ItemPrefix = 0
                        .EquippedObject(1).ItemSuffix = 0
                        SendSocket Index, Chr$(20) + Chr$(21)    'Stop Using Object
                        SendSocket Index, Chr$(19) + Chr$(Slot)    'Use Object
                        SendSocket Index, Chr$(17) + Chr$(C) + DoubleChar$(CLng(.Inv(C).Object)) + QuadChar(.Inv(C).Value) + Chr$(.Inv(C).ItemPrefix) + Chr$(.Inv(C).ItemSuffix)    'New Inv Obj
                    Else
                        SendSocket Index, Chr$(19) + Chr$(Slot)    'Use Object
                    End If
                    .EquippedObject(1).Object = A
                    .EquippedObject(1).Value = B
                    .EquippedObject(1).ItemPrefix = E
                    .EquippedObject(1).ItemSuffix = F
                    CalculateStats Index
                Else
                    SendSocket Index, Chr$(16) & Chr$(35)    'not enough stats
                End If
            Case 2, 3, 4
                If Object(.Inv(Slot).Object).LevelReq <= .Level Then
                    A = .Inv(Slot).Object
                    B = .Inv(Slot).Value
                    E = .Inv(Slot).ItemPrefix
                    F = .Inv(Slot).ItemSuffix
                    D = Object(.Inv(Slot).Object).Type
                    .Inv(Slot).Object = 0
                    .Inv(Slot).Value = 0
                    .Inv(Slot).ItemPrefix = 0
                    .Inv(Slot).ItemSuffix = 0
                    If .EquippedObject(D).Object > 0 Then
                        C = FreeInvNum(Index)
                        .Inv(C).Object = .EquippedObject(D).Object
                        .Inv(C).Value = .EquippedObject(D).Value
                        .Inv(C).ItemPrefix = .EquippedObject(D).ItemPrefix
                        .Inv(C).ItemSuffix = .EquippedObject(D).ItemSuffix
                        .EquippedObject(D).Object = 0
                        .EquippedObject(D).Value = 0
                        .EquippedObject(D).ItemPrefix = 0
                        .EquippedObject(D).ItemSuffix = 0
                        SendSocket Index, Chr$(20) + Chr$(20 + D)    'Stop Using Object
                        SendSocket Index, Chr$(19) + Chr$(Slot)    'Use Object
                        SendSocket Index, Chr$(17) + Chr$(C) + DoubleChar$(CLng(.Inv(C).Object)) + QuadChar(.Inv(C).Value) + Chr$(.Inv(C).ItemPrefix) + Chr$(.Inv(C).ItemSuffix)    'New Inv Obj
                    Else
                        SendSocket Index, Chr$(19) + Chr$(Slot)    'Use Object
                    End If
                    .EquippedObject(D).Object = A
                    .EquippedObject(D).Value = B
                    .EquippedObject(D).ItemPrefix = E
                    .EquippedObject(D).ItemSuffix = F
                    CalculateStats Index
                Else
                    SendSocket Index, Chr$(16) & Chr$(35)    'not enough stats
                End If
            Case 8    'Ring
                If Object(.Inv(Slot).Object).LevelReq <= .Level Then
                    A = .Inv(Slot).Object
                    B = .Inv(Slot).Value
                    E = .Inv(Slot).ItemPrefix
                    F = .Inv(Slot).ItemSuffix
                    D = 5
                    .Inv(Slot).Object = 0
                    .Inv(Slot).Value = 0
                    .Inv(Slot).ItemPrefix = 0
                    .Inv(Slot).ItemSuffix = 0
                    If .EquippedObject(D).Object > 0 Then
                        C = FreeInvNum(Index)
                        .Inv(C).Object = .EquippedObject(D).Object
                        .Inv(C).Value = .EquippedObject(D).Value
                        .Inv(C).ItemPrefix = .EquippedObject(D).ItemPrefix
                        .Inv(C).ItemSuffix = .EquippedObject(D).ItemSuffix
                        .EquippedObject(D).Object = 0
                        .EquippedObject(D).Value = 0
                        .EquippedObject(D).ItemPrefix = 0
                        .EquippedObject(D).ItemSuffix = 0
                        SendSocket Index, Chr$(20) + Chr$(20 + D)    'Stop Using Object
                        SendSocket Index, Chr$(19) + Chr$(Slot)    'Use Object
                        SendSocket Index, Chr$(17) + Chr$(C) + DoubleChar$(CLng(.Inv(C).Object)) + QuadChar(.Inv(C).Value) + Chr$(.Inv(C).ItemPrefix) + Chr$(.Inv(C).ItemSuffix)    'New Inv Obj
                    Else
                        SendSocket Index, Chr$(19) + Chr$(Slot)    'Use Object
                    End If
                    .EquippedObject(D).Object = A
                    .EquippedObject(D).Value = B
                    .EquippedObject(D).ItemPrefix = E
                    .EquippedObject(D).ItemSuffix = F
                    CalculateStats Index
                Else
                    SendSocket Index, Chr$(16) & Chr$(35)    'not enough stats
                End If
            Case 11    'Ammo
                If Object(.Inv(Slot).Object).LevelReq <= .Level Then
                    If .EquippedObject(6).Object > 0 Then
                        SendSocket Index, Chr$(20) + Chr$(.EquippedObject(6).Object)
                        .EquippedObject(6).Object = Slot
                        SendSocket Index, Chr$(19) + Chr$(Slot)
                        CalculateStats Index
                    Else
                        .EquippedObject(6).Object = Slot
                        SendSocket Index, Chr$(19) + Chr$(Slot)
                        CalculateStats Index
                    End If
                Else
                    SendSocket Index, Chr$(16) & Chr$(35)    'not enough stats
                End If
            End Select
        Else
            'No such object
        End If
    End With
End Sub
Sub UnEquipObject(Index As Long, Slot As Long)
    Dim A As Long
    With Player(Index)
        If Not Slot = 6 Then
            A = FreeInvNum(Index)
            If A > 0 And Not .EquippedObject(Slot).Object = 0 Then    'There is room
                If Object(.EquippedObject(Slot).Object).Type = 10 Then    'projectile weapon
                    If .EquippedObject(6).Object > 0 Then
                        SendSocket Index, Chr$(20) + Chr$(.EquippedObject(6).Object)
                        .EquippedObject(6).Object = 0
                        .EquippedObject(6).Value = 0
                        .EquippedObject(6).ItemPrefix = 0
                        .EquippedObject(6).ItemSuffix = 0
                        CalculateStats Index
                    End If
                End If
                .Inv(A).Object = .EquippedObject(Slot).Object
                .Inv(A).Value = .EquippedObject(Slot).Value
                .Inv(A).ItemPrefix = .EquippedObject(Slot).ItemPrefix
                .Inv(A).ItemSuffix = .EquippedObject(Slot).ItemSuffix
                .EquippedObject(Slot).Object = 0
                .EquippedObject(Slot).Value = 0
                .EquippedObject(Slot).ItemPrefix = 0
                .EquippedObject(Slot).ItemSuffix = 0
                SendSocket Index, Chr$(17) + Chr$(A) + DoubleChar$(CLng(.Inv(A).Object)) + QuadChar(.Inv(A).Value) + Chr$(.Inv(A).ItemPrefix) + Chr$(.Inv(A).ItemSuffix)    'New Inv Obj
                SendSocket Index, Chr$(20) + Chr$(20 + Slot)    'Stop Using Object
                CalculateStats Index
            Else
                SendSocket Index, Chr$(16) + Chr$(1)    'Inventory Full
            End If
        Else
            SendSocket Index, Chr$(20) + Chr$(.EquippedObject(6).Object)
            .EquippedObject(6).Object = 0
            .EquippedObject(6).Value = 0
            .EquippedObject(6).ItemPrefix = 0
            .EquippedObject(6).ItemSuffix = 0
            CalculateStats Index
        End If
    End With
End Sub

Function FindUnEquipInvObject(Index As Long, ObjectNum As Long) As Long
    Dim A As Long
    With Player(Index)
        For A = 1 To 20
            If .Inv(A).Object = ObjectNum Then
                If .EquippedObject(6).Object = A Then GoTo TheNextOne
                FindUnEquipInvObject = A
                Exit Function
            End If
TheNextOne:
        Next A
    End With
End Function

Sub SendBankData(Index As Long)
    Dim A As Long, St1 As String
    With Player(Index)
        SendSocket Index, Chr$(89) & QuadChar(.Bank)
        For A = 0 To 29
            If .ItemBank(A).Object > 0 Then    'Something there
                St1 = St1 & DoubleChar(10) & Chr$(113) & Chr$(A) & DoubleChar$(CLng(.ItemBank(A).Object)) + QuadChar(.ItemBank(A).Value) + Chr$(.ItemBank(A).ItemPrefix) + Chr$(.ItemBank(A).ItemSuffix)
            End If
        Next A
        If Len(St1) > 0 Then SendRaw Index, St1
    End With
End Sub

Sub ProcessBankData(Index As Long, St As String)
    Dim A As Long, B As Long, C As Long, D As Long, E As Long
    With Player(Index)
        Select Case Asc(Mid$(St, 1, 1))
        Case 1    'Deposit Object
            A = Asc(Mid$(St, 2, 1))    'Slot
            If A >= 1 And A <= 20 Then
                If .Inv(A).Object > 0 Then    'Deposit it
                    D = -1
                    For E = 0 To 29
                        If .ItemBank(E).Object = 0 And D = -1 Then    'Open Slot
                            D = E
                        End If
                    Next E
                    If D >= 0 Then
                        Parameter(0) = Index
                        Parameter(1) = .Inv(A).Value
                        If RunScript("DROPOBJ" + CStr(.Inv(A).Object)) = 0 Then
                            If Not Object(.Inv(A).Object).Type = 6 Then
                                .ItemBank(D).Object = .Inv(A).Object
                                .ItemBank(D).Value = .Inv(A).Value
                                .ItemBank(D).ItemPrefix = .Inv(A).ItemPrefix
                                .ItemBank(D).ItemSuffix = .Inv(A).ItemSuffix
                                .Inv(A).Object = 0
                                .Inv(A).Value = 0
                                .Inv(A).ItemPrefix = 0
                                .Inv(A).ItemSuffix = 0
                                SendSocket Index, Chr$(18) + Chr$(A)    'Remove inv object
                                SendBankData Index
                            End If
                        End If
                    Else
                        SendSocket Index, Chr$(16) + Chr$(44)
                    End If
                End If
            End If
        Case 2    'Add Gold
            C = FindInvObject(Index, CLng(World.ObjMoney))    'Money Slot
            If C > 0 Then    'Has money
                If Not Asc(Mid$(St, 2, 1)) > 120 Then
                    D = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                    If D > 0 Then
                        If .Inv(C).Value >= D Then    'Got the cash
                            .Bank = .Bank + D
                            TakeObj Index, CLng(World.ObjMoney), D
                            SendBankData Index
                        End If
                    Else
                        Hacker Index, "Deposit:  Gold Dupe"
                    End If
                Else
                    Hacker Index, "Deposit:  Gold Dupe"
                End If
            End If
        Case 3    'Remove Object
            A = Asc(Mid$(St, 2, 1))    'Slot
            If A >= 0 And A < 30 Then
                If .ItemBank(A).Object > 0 Then
                    B = -1
                    Select Case Object(.ItemBank(A).Object).Type
                    Case 6, 11
                        For C = 1 To 20
                            If .Inv(C).Object = .ItemBank(A).Object And .Inv(C).Value > 0 Then
                                B = C
                                Exit For
                            End If
                        Next C
                        If B = -1 Then B = FreeInvNum(Index)
                    Case Else
                        B = FreeInvNum(Index)
                    End Select
                    If B > 0 Then    'Has Room
                        Parameter(0) = Index
                        Parameter(1) = .ItemBank(A).Value
                        If RunScript("GETOBJ" + CStr(.ItemBank(A).Object)) = 0 Then
                            Select Case Object(.ItemBank(A).Object).Type
                            Case 6, 11
                                If .Inv(B).Object > 0 And .Inv(B).Object = .ItemBank(A).Object And .Inv(B).Value > 0 Then
                                    .Inv(B).Value = .Inv(B).Value + .ItemBank(A).Value
                                Else
                                    .Inv(B).Object = .ItemBank(A).Object
                                    .Inv(B).Value = .ItemBank(A).Value
                                End If
                                .Inv(B).ItemPrefix = .ItemBank(A).ItemPrefix
                                .Inv(B).ItemSuffix = .ItemBank(A).ItemSuffix
                            Case Else
                                .Inv(B).Object = .ItemBank(A).Object
                                .Inv(B).Value = .ItemBank(A).Value
                                .Inv(B).ItemPrefix = .ItemBank(A).ItemPrefix
                                .Inv(B).ItemSuffix = .ItemBank(A).ItemSuffix
                            End Select

                            .ItemBank(A).Object = 0
                            .ItemBank(A).Value = 0
                            .ItemBank(A).ItemPrefix = 0
                            .ItemBank(A).ItemSuffix = 0

                            SendSocket Index, Chr$(114) & Chr$(A)
                            SendSocket Index, Chr$(17) + Chr$(B) + DoubleChar$(CLng(.Inv(B).Object)) + QuadChar(.Inv(B).Value) + Chr$(.Inv(B).ItemPrefix) + Chr$(.Inv(B).ItemSuffix)    'New Inv Obj
                        End If
                    Else
                        SendSocket Index, Chr$(16) + Chr$(1)
                    End If
                End If
            End If
        Case 4    'Remove Gold
            D = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
            If D > 0 And .Bank >= D Then    'Got the cash
                C = FindInvObject(Index, CLng(World.ObjMoney))
                If C > 0 Then
                    .Bank = .Bank - D
                    GiveObj Index, CLng(World.ObjMoney), D
                    SendBankData Index
                Else
                    C = FreeInvNum(Index)
                    If C > 0 Then
                        .Bank = .Bank - D
                        GiveObj Index, CLng(World.ObjMoney), D
                        SendBankData Index
                    Else
                        SendSocket Index, Chr$(16) + Chr$(1)
                    End If
                End If
            End If
        Case 5    'Deposit Value Item
            A = Asc(Mid$(St, 2, 1))    'Slot
            B = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
            If A >= 1 And A <= 20 Then
                If .Inv(A).Object > 0 And .Inv(A).Value > 0 Then    'Deposit it
                    D = -1
                    Select Case Object(.Inv(A).Object).Type
                    Case 6, 11    'Value Items only
                        For E = 0 To 29
                            If .ItemBank(E).Object = .Inv(A).Object Then    'Open Slot
                                D = E
                                Exit For
                            End If
                        Next E
                        If D = -1 Then
                            For E = 0 To 29
                                If .ItemBank(E).Object = 0 Then    'Open Slot
                                    D = E
                                    Exit For
                                End If
                            Next E
                        End If
                    End Select
                    If D >= 0 Then
                        If B > 0 And .Inv(A).Value >= B Then
                            Parameter(0) = Index
                            Parameter(1) = B
                            If RunScript("DROPOBJ" + CStr(.Inv(A).Object)) = 0 Then

                                If .ItemBank(D).Object > 0 And .ItemBank(D).Value > 0 Then
                                    .ItemBank(D).Value = .ItemBank(D).Value + B
                                Else
                                    .ItemBank(D).Object = .Inv(A).Object
                                    .ItemBank(D).Value = B
                                End If

                                .ItemBank(D).ItemPrefix = .Inv(A).ItemPrefix
                                .ItemBank(D).ItemSuffix = .Inv(A).ItemSuffix

                                If .Inv(A).Value - B > 0 Then
                                    .Inv(A).Value = .Inv(A).Value - B
                                    SendSocket Index, Chr$(17) + Chr$(A) + DoubleChar$(CLng(.Inv(A).Object)) + QuadChar(.Inv(A).Value) + Chr$(.Inv(A).ItemPrefix) + Chr$(.Inv(A).ItemSuffix)    'New Inv Obj
                                Else
                                    .Inv(A).Object = 0
                                    .Inv(A).Value = 0
                                    .Inv(A).ItemPrefix = 0
                                    .Inv(A).ItemSuffix = 0
                                    SendSocket Index, Chr$(18) + Chr$(A)    'Remove inv object
                                End If
                                SendBankData Index
                            End If
                        End If
                    Else
                        SendSocket Index, Chr$(16) + Chr$(44)
                    End If
                End If
            End If
        Case 6    'Withdraw Value Item
            A = Asc(Mid$(St, 2, 1))    'Slot
            B = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
            D = .ItemBank(A).Object
            If A >= 0 And A <= 29 Then
                If B > 0 And .ItemBank(A).Object > 0 And .ItemBank(A).Value >= B Then
                    If Object(D).Type = 6 Or Object(D).Type = 11 Then
                        C = FindInvObject(Index, D)
                        If C > 0 Then
                            Parameter(0) = Index
                            Parameter(1) = B
                            If RunScript("GETOBJ" + CStr(D)) = 0 Then
                                .ItemBank(A).Value = .ItemBank(A).Value - B
                                If .ItemBank(A).Value = 0 Then
                                    .ItemBank(A).Object = 0
                                    .ItemBank(A).ItemPrefix = 0
                                    .ItemBank(A).ItemSuffix = 0
                                    SendSocket Index, Chr$(114) & Chr$(A)
                                End If
                                GiveObj Index, D, B
                                SendBankData Index
                            End If
                        Else
                            C = FreeInvNum(Index)
                            If C > 0 Then
                                Parameter(0) = Index
                                Parameter(1) = B
                                If RunScript("GETOBJ" + CStr(D)) = 0 Then
                                    .ItemBank(A).Value = .ItemBank(A).Value - B
                                    If .ItemBank(A).Value = 0 Then
                                        .ItemBank(A).Object = 0
                                        .ItemBank(A).ItemPrefix = 0
                                        .ItemBank(A).ItemSuffix = 0
                                        SendSocket Index, Chr$(114) & Chr$(A)
                                    End If
                                    GiveObj Index, D, B
                                    SendBankData Index
                                End If
                            Else
                                SendSocket Index, Chr$(16) + Chr$(1)
                            End If
                        End If
                    End If
                End If
            End If
        End Select
    End With
End Sub

Sub RepairItem(Index As Long)
    Dim A As Long, B As Long, C As Long
    With Player(Index)
        If .CurrentRepairTar <= 20 Then
            B = .Inv(.CurrentRepairTar).Object    'Object
        Else
            B = .EquippedObject(.CurrentRepairTar - 20).Object    'Object
        End If
        If ExamineBit(Object(B).flags, 1) = 255 Then B = 0
        If ExamineBit(Object(B).flags, 0) = 255 Then B = 0
        If B > 0 Then    'Slot isn't empty
            A = GetRepairCost(Index, .CurrentRepairTar)    'Cost
            C = FindInvObject(Index, CLng(World.ObjMoney))    'Money Slot
            If C > 0 Then    'Has money
                If .Inv(C).Value >= A Then    'Has the Cash
                    TakeObj Index, CLng(World.ObjMoney), A    'Take Cash
                    If .CurrentRepairTar <= 20 Then
                        .Inv(.CurrentRepairTar).Object = B
                        Select Case Object(B).Type
                        Case 1, 2, 3, 4
                            .Inv(.CurrentRepairTar).Value = Object(B).Data(0) * 10
                        Case 8
                            .Inv(.CurrentRepairTar).Value = Object(B).Data(1) * 10
                        End Select
                        SendSocket Index, Chr$(17) + Chr$(.CurrentRepairTar) + DoubleChar$(B) + QuadChar(.Inv(.CurrentRepairTar).Value) + Chr$(.Inv(.CurrentRepairTar).ItemPrefix) + Chr$(.Inv(.CurrentRepairTar).ItemSuffix)
                    Else
                        .EquippedObject(.CurrentRepairTar - 20).Object = B
                        Select Case Object(B).Type
                        Case 1, 2, 3, 4
                            .EquippedObject(.CurrentRepairTar - 20).Value = Object(B).Data(0) * 10
                        Case 8
                            .EquippedObject(.CurrentRepairTar - 20).Value = Object(B).Data(1) * 10
                        End Select
                        SendSocket Index, Chr$(115) + DoubleChar$(B) + QuadChar(.EquippedObject(.CurrentRepairTar - 20).Value) + Chr$(.EquippedObject(.CurrentRepairTar - 20).ItemPrefix) + Chr$(.EquippedObject(.CurrentRepairTar - 20).ItemSuffix)
                    End If
                    SendSocket Index, Chr$(98) + Chr$(2) + DoubleChar$(B)
                Else
                    SendSocket Index, Chr$(16) + Chr$(33)
                End If
            Else
                SendSocket Index, Chr$(16) + Chr$(33)
            End If
        End If
    End With
End Sub

Sub RepairAll(Index As Long)
    Dim A As Long
    With Player(Index)
        For A = 1 To 25
            If A <= 20 Then
                If .Inv(A).Object > 0 Then
                    Select Case Object(.Inv(A).Object).Type
                    Case 1, 2, 3, 4, 8
                        If GetRepairCost(Index, CInt(A)) > 0 Then
                            .CurrentRepairTar = A
                            RepairItem Index
                        End If
                    End Select
                End If
            Else
                If .EquippedObject(A - 20).Object > 0 Then
                    Select Case Object(.EquippedObject(A - 20).Object).Type
                    Case 1, 2, 3, 4, 8
                        If GetRepairCost(Index, CInt(A)) > 0 Then
                            .CurrentRepairTar = A
                            RepairItem Index
                        End If
                    End Select
                End If
            End If
        Next A
    End With
End Sub

Function NoDirectionalWalls(TheMap As Long, X As Long, Y As Long, Direction As Long) As Boolean
    If X < 0 Or Y < 0 Or X > 11 Or Y > 11 Then
        NoDirectionalWalls = False
        Exit Function
    End If

    NoDirectionalWalls = True
    Select Case Direction
    Case 0    'Up
        If Y >= 0 Then
            If Y > 0 Then
                If Map(TheMap).Tile(X, Y - 1).Att = 17 Then
                    If ExamineBit(Map(TheMap).Tile(X, Y - 1).AttData(0), 3) Then
                        NoDirectionalWalls = False
                        Exit Function
                    End If
                End If
            End If
            If Map(TheMap).Tile(X, Y).Att = 17 Then
                If ExamineBit(Map(TheMap).Tile(X, Y).AttData(0), 1) Then
                    NoDirectionalWalls = False
                    Exit Function
                End If
            End If
        End If
    Case 1    'Down
        If Y < 12 Then
            If Y < 11 Then
                If Map(TheMap).Tile(X, Y + 1).Att = 17 Then
                    If ExamineBit(Map(TheMap).Tile(X, Y + 1).AttData(0), 0) Then
                        NoDirectionalWalls = False
                        Exit Function
                    End If
                End If
            End If
            If Map(TheMap).Tile(X, Y).Att = 17 Then
                If ExamineBit(Map(TheMap).Tile(X, Y).AttData(0), 2) Then
                    NoDirectionalWalls = False
                    Exit Function
                End If
            End If
        End If
    Case 2    'Left
        If X >= 0 Then
            If X > 0 Then
                If Map(TheMap).Tile(X - 1, Y).Att = 17 Then
                    If ExamineBit(Map(TheMap).Tile(X - 1, Y).AttData(0), 6) Then
                        NoDirectionalWalls = False
                        Exit Function
                    End If
                End If
            End If
            If Map(TheMap).Tile(X, Y).Att = 17 Then
                If ExamineBit(Map(TheMap).Tile(X, Y).AttData(0), 4) Then
                    NoDirectionalWalls = False
                    Exit Function
                End If
            End If
        End If
    Case 3    'Right
        If X < 12 Then
            If X < 11 Then
                If Map(TheMap).Tile(X + 1, Y).Att = 17 Then
                    If ExamineBit(Map(TheMap).Tile(X + 1, Y).AttData(0), 5) Then
                        NoDirectionalWalls = False
                        Exit Function
                    End If
                End If
            End If
            If Map(TheMap).Tile(X, Y).Att = 17 Then
                If ExamineBit(Map(TheMap).Tile(X, Y).AttData(0), 7) Then
                    NoDirectionalWalls = False
                    Exit Function
                End If
            End If
        End If
    End Select
End Function

Sub SendServerOptions(Index As Long)
    SendSocket Index, Chr$(139) + Chr$(World.StatStrength) + Chr$(World.StatEndurance) + Chr$(World.StatIntelligence) + Chr$(World.StatConcentration) + Chr$(World.StatConstitution) + Chr$(World.StatStamina) + Chr$(World.StatWisdom) + Chr$(World.ObjMoney) + DoubleChar$(CLng(World.Cost_Per_Durability)) + DoubleChar$(CLng(World.Cost_Per_Strength)) + DoubleChar$(CLng(World.Cost_Per_Modifier)) + Chr$(World.GuildJoinLevel) + Chr$(World.GuildNewLevel) + QuadChar(World.GuildJoinPrice) + QuadChar(World.GuildNewPrice)
End Sub

Sub WriteString(lpAppName, lpKeyName As String, A, Optional lpFileName As String = "odyssey.ini")
    Dim lpString As String, Valid As Long
    lpString = A
    Valid = WritePrivateProfileString&(lpAppName, lpKeyName, lpString, App.Path + "\" + lpFileName)
End Sub

Sub RemoveFromGuild(Name As String, TheGuild As Long)
    Dim A As Long
    A = FindGuildMember(Name, TheGuild)
    If A >= 0 Then
        With Guild(TheGuild).Member(A)
            .Name = ""
            .Rank = 0
            .JoinDate = 0
            .Kills = 0
            .Deaths = 0
        End With
        GuildRS.Bookmark = Guild(TheGuild).Bookmark
        GuildRS.Edit
        GuildRS("MemberName" + CStr(A)) = ""
        GuildRS("MemberRank" + CStr(A)) = 0
        GuildRS("MemberJoinDate" + CStr(A)) = 0
        GuildRS("MemberKills" + CStr(A)) = 0
        GuildRS("MemberDeaths" + CStr(A)) = 0
        GuildRS.Update
    End If
End Sub

Function FindGuild(Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long, B As Long
    For A = 1 To MaxGuilds
        With Guild(A)
            If .Name <> vbNullString Then
                For B = 0 To 19
                    If .Member(B).Name <> vbNullString Then
                        If UCase$(.Member(B).Name) = Name Then
                            FindGuild = A
                            Exit Function
                        End If
                    End If
                Next B
            End If
        End With
    Next A

    FindGuild = 0
End Function

Function getTime() As Currency
    'getTime = GetTickCount64
    getTime = timeGetTime
End Function

Function RandomPrefix() As Byte
    Dim A As Long
    Dim TotalPrefixes As Long
    
    Randomize
    
    If Int(Rnd * 100) <= PrefixSuffixChance Then
        For A = 1 To MaxModifications
            If Len(ItemPrefix(A).Name) = 0 Then
                TotalPrefixes = A - 1
                Exit For
            End If
        Next A
        
        If TotalPrefixes > 0 Then
            For A = 1 To 5
                RandomPrefix = Int(Rnd * TotalPrefixes) + 1
                If ItemPrefix(RandomPrefix).OccursNaturally Then
                    Exit For
                End If
            Next A
            
            If ItemPrefix(A).OccursNaturally Then
                
            Else
                RandomPrefix = 0
            End If
        End If
    Else
        RandomPrefix = 0
    End If
End Function

Function RandomSuffix() As Byte
    Dim A As Long
    Dim TotalSuffixes As Long
    
    Randomize
    
    If Int(Rnd * 100) <= PrefixSuffixChance Then
        For A = 1 To MaxModifications
            If Len(ItemSuffix(A).Name) = 0 Then
                TotalSuffixes = A - 1
                Exit For
            End If
        Next A
        
        If TotalSuffixes > 0 Then
            For A = 1 To 5
                RandomSuffix = Int(Rnd * TotalSuffixes) + 1
                If ItemSuffix(RandomSuffix).OccursNaturally Then
                    Exit For
                End If
            Next A
            
            If ItemSuffix(A).OccursNaturally Then
                
            Else
                RandomSuffix = 0
            End If
        End If
    Else
        RandomSuffix = 0
    End If
End Function
