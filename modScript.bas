Attribute VB_Name = "modScript"
Option Explicit

Declare Function RunASMScript Lib "script.dll" Alias "RunScript" (Script As Any, FunctionTable As Any, Parameters As Any) As Long
Declare Function SysFreeString Lib "oleaut32.dll" (ByVal StringPointer As Long) As Long
Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal St As String, ByVal Length As Long) As Long

Public FunctionTable(0 To 255) As Long
Public ScriptRunning As Boolean

Public Parameter(0 To 5) As Long
Public StringStack(0 To 1023) As Long
Public StringPointer As Long

Public LastScript As String
Public LastScript2 As String
Public LastScript3 As String

'Public LastParameter(0 To 5) As Long

Sub Boot_Player(ByVal Index As Long, ByVal Reason As String)
    BootPlayer Index, 0, StrConv(Reason, vbUnicode)
End Sub
Sub Ban_Player(ByVal Index As Long, ByVal NumDays As Long, ByVal Reason As String)
    BanPlayer Index, 0, NumDays, StrConv(Reason, vbUnicode), "Script Ban"
End Sub
Function Find_Player(ByVal Name As String) As Long
    Find_Player = FindPlayer(StrConv(Name, vbUnicode))
End Function

Function GetObjX(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And ObjIndex >= 0 And ObjIndex <= MaxMapObjects Then
        With Map(MapIndex).Object(ObjIndex)
            GetObjX = .X
        End With
    End If
End Function
Function GetObjY(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And ObjIndex >= 0 And ObjIndex <= MaxMapObjects Then
        With Map(MapIndex).Object(ObjIndex)
            GetObjY = .Y
        End With
    End If
End Function
Function GetObjNum(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And ObjIndex >= 0 And ObjIndex <= MaxMapObjects Then
        With Map(MapIndex).Object(ObjIndex)
            GetObjNum = .Object
        End With
    End If
End Function
Function GetTileAtt(ByVal MapIndex As Long, ByVal X As Long, ByVal Y As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        GetTileAtt = Map(MapIndex).Tile(X, Y).Att
    End If
End Function

Function GetTileAtt2(ByVal MapIndex As Long, ByVal X As Long, ByVal Y As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        GetTileAtt2 = Map(MapIndex).Tile(X, Y).Att2
    End If
End Function

Function GetTileIsVacant(ByVal MapIndex As Long, ByVal X As Long, ByVal Y As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        GetTileIsVacant = PlayerIsVacant(MapIndex, X, Y)
    End If
End Function

Function GetTileNoDirectionalWalls(ByVal MapIndex As Long, ByVal X As Long, ByVal Y As Long, ByVal Direction As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 And Direction >= 0 And Direction <= 3 Then
        GetTileNoDirectionalWalls = NoDirectionalWalls(MapIndex, X, Y, Direction)
    End If
End Function

Function DestroyObj(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And ObjIndex >= 0 And ObjIndex <= MaxMapObjects Then
        With Map(MapIndex).Object(ObjIndex)
            .Object = 0
            .Value = 0
            .ItemPrefix = 0
            .ItemSuffix = 0
            SendToMap MapIndex, Chr$(15) + Chr$(ObjIndex)    'Erase Map Obj
        End With
    End If
End Function
Function GetObjVal(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And ObjIndex >= 0 And ObjIndex <= MaxMapObjects Then
        With Map(MapIndex).Object(ObjIndex)
            GetObjVal = .Value
        End With
    End If
End Function


Function NewString(St As String) As Long
    Dim A As Long
    If StringPointer < 1024 Then
        A = SysAllocStringByteLen(St, Len(St))
        StringStack(StringPointer) = A
        StringPointer = StringPointer + 1
        NewString = A
    Else
        NewString = 0
    End If
End Function

Function AttackMonster(ByVal Index As Long, ByVal MonsterIndex As Long, ByVal Damage As Long) As Long
    Dim MapNum As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        If Player(Index).Mode = modePlaying Then
            MapNum = Player(Index).Map
            If Map(MapNum).Monster(MonsterIndex).Monster > 0 Then
                ScriptRunning = False
                Parameter(0) = Index
                If RunScript("ATTACKMONSTER" + CStr(Map(MapNum).Monster(MonsterIndex).Monster)) = 0 Then
                    If Damage < 0 Then Damage = 0
                    If Damage > 255 Then Damage = 255
                    With Map(MapNum).Monster(MonsterIndex)
                        Damage = Damage - Monster(.Monster).Armor
                        If Damage < 0 Then Damage = 0
                        SendToMap MapNum, Chr$(44) + Chr$(Index) + Chr$(0) + Chr$(MonsterIndex) + Chr$(Damage) + DoubleChar$(CLng(.HP))    'Hit Monster
                        .Target = Index
                        .TargetIsMonster = False
                        If .HP > Damage Then
                            .HP = .HP - Damage
                        Else
                            'Monster Died
                            SendToMapAllBut MapNum, Index, Chr$(39) + Chr$(MonsterIndex)    'Monster Died
                            If ExamineBit(Monster(.Monster).flags, 4) = False Then
                                GainExp Index, CLng(Monster(.Monster).Experience)
                            Else
                                GainEliteExp Index, CLng(Monster(.Monster).Experience)
                            End If
                            SendSocket Index, Chr$(51) + Chr$(MonsterIndex) + QuadChar(Player(Index).Experience)    'You killed monster
                            B = Int(Rnd * 3)
                            C = Monster(.Monster).Object(B)
                            If C > 0 Then
                                NewMapObject MapNum, C, Monster(.Monster).Value(B), CLng(.X), CLng(.Y), False
                            End If

                            Parameter(0) = Index
                            RunScript "MONSTERDIE" + CStr(.Monster)
                            
                            .Monster = 0
                            AttackMonster = True
                        End If
                    End With
                End If
                ScriptRunning = True
            End If
        End If
    End If
End Function
Function AttackPlayer(ByVal Index As Long, ByVal Target As Long, ByVal Damage As Long) As Long
    If Index >= 1 And Index <= MaxUsers And Target >= 1 And Target <= MaxUsers Then
        If Player(Index).Mode = modePlaying And Player(Target).Mode = modePlaying And Player(Target).IsDead = False Then
            If Damage < 0 Then Damage = 0
            If Damage > 255 Then Damage = 255
            ScriptRunning = False
            CombatAttackPlayer Index, Target, Damage
            ScriptRunning = True
            AttackPlayer = True
        End If
    End If
End Function
Function CanAttackMonster(ByVal Index As Long, ByVal MonsterIndex As Long) As Long
    Dim MapIndex As Long
    If Index >= 1 And Index <= MaxUsers And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        If Player(Index).Mode = modePlaying Then
            MapIndex = Player(Index).Map
            If ExamineBit(Map(MapIndex).flags, 5) = False And Map(MapIndex).Monster(MonsterIndex).Monster > 0 Then
                CanAttackMonster = True
            End If
        End If
    End If
End Function
Function CanAttackPlayer(ByVal Player1 As Long, ByVal Player2 As Long) As Long
    Dim PKMap As Boolean
    If Player1 >= 1 And Player1 <= MaxUsers And Player2 >= 1 And Player2 <= MaxUsers Then
        With Player(Player1)
            If ExamineBit(Map(.Map).flags, 0) = False Then
                If .Mode = modePlaying And Player(Player2).Mode = modePlaying Then
                    If .Map = Player(Player2).Map Then
                        If .IsDead = 0 And Player(Player2).IsDead = 0 Then
                            PKMap = ExamineBit(Map(.Map).flags, 6)
                            If .Guild > 0 Or PKMap = True Then
                                If Player(Player2).Guild > 0 Or PKMap = True Then
                                    CanAttackPlayer = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End If
End Function
Function GetAbs(ByVal Value As Long) As Long
    GetAbs = Abs(Value)
End Function
Function GetMaxUsers() As Long
    GetMaxUsers = MaxUsers
End Function
Function GetMonsterType(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        GetMonsterType = Map(MapIndex).Monster(MonsterIndex).Monster
    End If
End Function
Function GetMonsterHP(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        GetMonsterHP = Map(MapIndex).Monster(MonsterIndex).HP
    End If
End Function
Function GetMonsterTarget(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        GetMonsterTarget = Map(MapIndex).Monster(MonsterIndex).Target
    End If
End Function
Sub NPCSay(ByVal MapIndex As Long, ByVal St As String)
    If MapIndex >= 1 And MapIndex <= MaxMaps Then
        If Map(MapIndex).NPC > 0 Then
            SendToMap MapIndex, Chr$(88) + DoubleChar$(CLng(Map(MapIndex).NPC)) + StrConv(St, vbUnicode)
        End If
    End If
End Sub
Sub NPCTell(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                A = Map(Player(Index).Map).NPC
                If A > 0 Then
                    SendSocket Index, Chr$(88) + DoubleChar$(A) + StrConv(St, vbUnicode)
                End If
            End If
        End With
    End If
End Sub

Sub ScriptTimer(ByVal Index As Long, ByVal Seconds As Long, ByVal Script As String)
    Dim A As Long, Tick As Currency
    If Index >= 1 And Index <= MaxUsers Then
        If Seconds > 86400 Then Seconds = 86400
        If Seconds < 0 Then Seconds = 0
        With Player(Index)
            If .Mode = modePlaying Then
                Tick = getTime
                For A = 1 To MaxPlayerTimers
                    If .ScriptTimer(A) = 0 Then
                        .Script(A) = StrConv(Script, vbUnicode)
                        .ScriptTimer(A) = Tick + Seconds * 1000
                        If A > 15 Then
                            SendToGods Chr$(56) + Chr$(7) + "WARNING:  " + CStr(.Name) + " has " + CStr(Val(A)) + " timers running"
                        End If
                        Exit For
                        'Parameter(0) = Index
                        '.ScriptTimer = 0
                        'ScriptRunning = False
                        'RunScript .Script
                        'ScriptRunning = True
                    End If
                Next A
            End If
        End With
    End If
End Sub
Sub SetFlag(ByVal FlagNum As Long, ByVal Value As Long)
    If FlagNum >= 0 And FlagNum <= 255 Then
        If Value >= 0 Then
            World.Flag(FlagNum) = Value
        End If
    End If
End Sub
Function GetFlag(ByVal FlagNum As Long) As Long
    If FlagNum >= 0 And FlagNum <= 255 Then
        GetFlag = World.Flag(FlagNum)
    End If
End Function

Function SetMonsterTarget(ByVal MapIndex As Long, ByVal MonsterIndex As Long, ByVal Player As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters And Player >= 1 And Player <= MaxUsers Then
        With Map(MapIndex).Monster(MonsterIndex)
            .Target = Player
            .TargetIsMonster = False
        End With
    End If
End Function

Function SetMonsterHP(ByVal MapIndex As Long, ByVal MonsterIndex As Long, ByVal HP As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters And HP >= 1 And HP <= 32000 Then
        If Map(MapIndex).Monster(MonsterIndex).Monster > 0 Then
            With Map(MapIndex).Monster(MonsterIndex)
                .HP = HP
                If .HP > Monster(Map(MapIndex).Monster(MonsterIndex).Monster).HP Then .HP = Monster(Map(MapIndex).Monster(MonsterIndex).Monster).HP
                SendToMap MapIndex, Chr$(142) + Chr$(MonsterIndex) + Chr$(Monster(Map(MapIndex).Monster(MonsterIndex).Monster).HP)
            End With
        End If
    End If
End Function

Function GetMonsterX(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        GetMonsterX = Map(MapIndex).Monster(MonsterIndex).X
    End If
End Function
Function GetMonsterY(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        GetMonsterY = Map(MapIndex).Monster(MonsterIndex).Y
    End If
End Function

Function GetMonsterDirection(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        GetMonsterDirection = Map(MapIndex).Monster(MonsterIndex).D
    End If
End Function

Function GetSqr(ByVal Value As Long) As Long
    GetSqr = Sqr(Value)
End Function

Function HasObj(ByVal Index As Long, ByVal ObjIndex As Long) As Long
    Dim A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers And ObjIndex >= 1 And ObjIndex <= MaxObjects Then
        B = Object(ObjIndex).Type
        With Player(Index)
            For A = 1 To 20
                With .Inv(A)
                    If .Object = ObjIndex Then
                        If B = 6 Or B = 11 Then
                            C = .Value
                            Exit For
                        Else
                            C = C + 1
                        End If
                    End If
                End With
            Next A
            For A = 1 To 5
                With .EquippedObject(A)
                    If .Object = ObjIndex Then
                        C = C + 1
                    End If
                End With
            Next A
        End With
        HasObj = C
    End If
End Function
Function GetPlayerName(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            GetPlayerName = NewString(.Name)
        End With
    End If
End Function
Function GetPlayerIP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            GetPlayerIP = NewString(.IP)
        End With
    End If
End Function

Function GetPlayerUser(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            GetPlayerUser = NewString(.User)
        End With
    End If
End Function
Function GetPlayerDesc(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            GetPlayerDesc = NewString(.desc)
        End With
    End If
End Function
Function GetGuildName(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        With Guild(Index)
            GetGuildName = NewString(.Name)
        End With
    End If
End Function
Function GiveObj(ByVal Index As Long, ByVal ObjIndex As Long, ByVal Amount As Long) As Long
    Dim A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers And ObjIndex >= 1 And ObjIndex <= MaxObjects Then
        With Player(Index)
            If .Mode = modePlaying Then
                B = Object(ObjIndex).Type
                If B = 6 Or B = 11 Then
                    A = FindInvObject(Index, ObjIndex)
                    If A = 0 Then
                        A = FreeInvNum(Index)
                    Else
                        C = 1
                    End If
                Else
                    A = FreeInvNum(Index)
                End If
                If A > 0 Then
                    PrintItem .User + " - " + .Name + " (GiveObj) " + Object(ObjIndex).Name + " (" + CStr(Amount) + ") - Map: " + CStr(.Map)
                    With .Inv(A)
                        .Object = ObjIndex
                        Select Case B
                        Case 1, 2, 3, 4    'Weapon, Shield, Armor, Helmut
                            .Value = CLng(Object(ObjIndex).Data(0)) * 10
                        Case 6, 11    'Money
                            If C = 1 Then
                                .Value = .Value + Amount
                            Else
                                .Value = Amount
                            End If
                            If .Value = 0 Then
                                .Object = 0
                                .ItemPrefix = 0
                                .ItemSuffix = 0
                                Exit Function
                            End If
                        Case 8    'Ring
                            .Value = CLng(Object(ObjIndex).Data(1)) * 10
                        Case Else
                            .Value = 0
                        End Select

                        SendSocket Index, Chr$(17) + Chr$(A) + DoubleChar$(ObjIndex) + QuadChar(.Value) + Chr$(.ItemPrefix) + Chr$(.ItemSuffix)    'New Inv Obj
                    End With
                End If
            End If
        End With
    End If
End Function

Function TakeObj(ByVal Index As Long, ByVal ObjIndex As Long, ByVal Amount As Long) As Long
    Dim A As Long
    If Index >= 1 And Index <= MaxUsers And ObjIndex >= 1 And ObjIndex <= MaxObjects Then
        With Player(Index)
            If .Mode = modePlaying Then
                A = FindInvObject(Index, ObjIndex)
                If A > 0 Then
                    PrintItem .User + " - " + .Name + " (TakeObj) " + Object(ObjIndex).Name + " (" + CStr(Amount) + ") - Map: " + CStr(.Map)
                    With .Inv(A)
                        If Object(ObjIndex).Type = 6 Or Object(ObjIndex).Type = 11 Then
                            If .Value >= Amount Then
                                .Value = .Value - Amount
                                If .Value = 0 Then
                                    .Object = 0
                                    .Value = 0
                                    .ItemPrefix = 0
                                    .ItemSuffix = 0
                                    SendSocket Index, Chr$(18) + Chr$(A)
                                Else
                                    SendSocket Index, Chr$(17) + Chr$(A) + DoubleChar$(ObjIndex) + QuadChar(.Value) + Chr$(.ItemPrefix) + Chr$(.ItemSuffix)    'New Inv Obj
                                End If
                                TakeObj = Amount
                            End If
                        Else
                            .Object = 0
                            .Value = 0
                            .ItemPrefix = 0
                            .ItemSuffix = 0
                            TakeObj = 1
                            SendSocket Index, Chr$(18) + Chr$(A)
                        End If
                    End With
                End If
            End If
        End With
    End If
End Function
Sub GlobalMessage(ByVal Message As String, ByVal MsgColor As Long)
    MsgColor = MsgColor Mod 16
    SendAll Chr$(56) + Chr$(MsgColor) + StrConv(Message, vbUnicode)
End Sub

Function IsPlaying(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        IsPlaying = Player(Index).Mode = modePlaying
    End If
End Function

Sub MapMessage(ByVal Index As Long, ByVal Message As String, ByVal MsgColor As Long)
    If Index >= 1 And Index <= MaxMaps Then
        MsgColor = MsgColor Mod 16
        SendToMap Index, Chr$(56) + Chr$(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Sub MapMessageAllBut(ByVal MapIndex As Long, ByVal PlayerIndex As Long, ByVal Message As String, ByVal MsgColor As Long)
    If MapIndex >= 1 And MapIndex <= MaxMaps Then
        MsgColor = MsgColor Mod 16
        SendToMapAllBut MapIndex, PlayerIndex, Chr$(56) + Chr$(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Function OpenDoor(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    Dim A As Long
    If MapNum >= 1 And MapNum <= MaxMaps And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        A = FreeMapDoorNum(MapNum)
        If A >= 0 Then
            With Map(MapNum).Door(A)
                .Att = Map(MapNum).Tile(X, Y).Att
                .X = X
                .Y = Y
                .T = getTime
            End With
            Map(MapNum).Tile(X, Y).Att = 0
            SendToMap MapNum, Chr$(36) + Chr$(A) + Chr$(X) + Chr$(Y)
            OpenDoor = 1
        End If
    End If
End Function
Sub PlayerMessage(ByVal Index As Long, ByVal Message As String, ByVal MsgColor As Long)
    If Index >= 1 And Index <= MaxUsers Then
        MsgColor = MsgColor Mod 16
        SendSocket Index, Chr$(56) + Chr$(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Function RunScript0(ByVal Script As String) As Long
    ScriptRunning = False
    RunScript0 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function
Function RunScript1(ByVal Script As String, ByVal Parm1 As Long) As Long
    Parameter(0) = Parm1
    ScriptRunning = False
    RunScript1 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function
Function RunScript2(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    ScriptRunning = False
    RunScript2 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function
Function RunScript3(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    ScriptRunning = False
    RunScript3 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function
Function RunScript4(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long, ByVal Parm4 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    Parameter(3) = Parm4
    ScriptRunning = False
    RunScript4 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    If Index >= 1 And Index <= MaxUsers And Sprite >= 0 And Sprite <= MaxSprite Then
        With Player(Index)
            If Sprite = 0 Then
                If .Guild > 0 Then
                    If Guild(.Guild).Sprite > 0 Then
                        .Sprite = Guild(.Guild).Sprite
                    Else
                        .Sprite = .Class * 2 + .Gender - 1
                    End If
                Else
                    .Sprite = .Class * 2 + .Gender - 1
                End If
            Else
                .Sprite = Sprite
            End If
            SendToMap .Map, Chr$(63) + Chr$(Index) + DoubleChar$(CLng(.Sprite))
        End With
    End If
End Sub
Function SpawnMonster(ByVal MapIndex As Long, ByVal Monster As Long, ByVal X As Long, ByVal Y As Long) As Long
    Dim A As Long

    If MapIndex >= 1 And MapIndex <= MaxMaps And Monster >= 1 And Monster <= MaxTotalMonsters And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        With Map(MapIndex)
            For A = 0 To MaxMonsters
                With .Monster(A)
                    If .Monster = 0 Then
                        SendToMapRaw MapIndex, SpawnMapMonster(MapIndex, A, Monster, X, Y)
                        SpawnMonster = 1
                        Exit Function
                    End If
                End With
            Next A
        End With
    End If

    SpawnMonster = 0
End Function

Function DespawnMonster(ByVal MapIndex As Long, ByVal Monster As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And Monster >= 0 And Monster <= MaxMonsters Then
        With Map(MapIndex)
            With .Monster(Monster)
                If .Monster > 0 Then
                    .Monster = 0
                    SendToMap MapIndex, Chr$(39) + Chr(Monster)
                    DespawnMonster = 1
                    Exit Function
                End If
            End With
        End With
    End If

    DespawnMonster = 0
End Function

Function SpawnObject(ByVal MapIndex As Long, ByVal Object As Long, ByVal Value As Long, ByVal X As Long, ByVal Y As Long) As Long
    If MapIndex >= 1 And MapIndex <= MaxMaps And Object >= 1 And Object <= MaxObjects And Value >= 0 And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        SpawnObject = NewMapObject(MapIndex, Object, Value, X, Y, False)
    End If
End Function
Function Str(ByVal Value As Long) As Long
    Str = NewString(CStr(Value))
End Function
Function StrCat(ByVal String1 As String, ByVal String2 As String) As Long
    StrCat = NewString(StrConv(String1, vbUnicode) + StrConv(String2, vbUnicode))
End Function

Function StrCmp(ByVal String1 As String, ByVal String2 As String) As Long
    StrCmp = UCase$(StrConv(String1, vbUnicode)) = UCase$(StrConv(String2, vbUnicode))
End Function
Function GetInStr(ByVal String1 As String, ByVal String2 As String) As Long
    GetInStr = InStr(UCase$(StrConv(String1, vbUnicode)), UCase$(StrConv(String2, vbUnicode)))
End Function

Function StrFormat(ByVal String1 As String, ByVal String2 As String) As Long
    Dim St As String, St1 As String, St2 As String
    St1 = StrConv(String1, vbUnicode)
    St2 = StrConv(String2, vbUnicode)

    Dim A As Long, B As Byte
    For A = 1 To Len(St1)
        B = Asc(Mid$(St1, A, 1))
        If B = 42 Then
            St = St + St2
        Else
            St = St + Chr$(B)
        End If
    Next A

    StrFormat = NewString(St)
End Function
Function GetGuildHall(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildHall = Guild(Index).Hall
    End If
End Function
Function GetGuildBank(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildBank = Guild(Index).Bank
    End If
End Function
Function GetGuildSprite(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildSprite = Guild(Index).Sprite
    End If
End Function

Function GetGuildMemberCount(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildMemberCount = CountGuildMembers(Index)
    End If
End Function
Function GetMapPlayerCount(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxMaps Then
        GetMapPlayerCount = Map(Index).NumPlayers
    End If
End Function

Function GetPlayerAccess(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerAccess = Player(Index).Access
    End If
End Function
Function GetPlayerAgility(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerAgility = statPlayerAgility
    End If
End Function

Function GetPlayerBank(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBank = Player(Index).Bank
    End If
End Function
Function GetPlayerClass(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerClass = Player(Index).Class
    End If
End Function
Function GetPlayerEndurance(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerEndurance = World.StatEndurance
    End If
End Function
Function GetPlayerEnergy(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerEnergy = Player(Index).Energy
    End If
End Function
Function GetPlayerEquipped(ByVal Index As Long, ByVal EquippedIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers And EquippedIndex >= 1 And EquippedIndex <= 6 Then
        GetPlayerEquipped = Player(Index).EquippedObject(EquippedIndex).Object
    End If
End Function

Function GetPlayerExperience(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerExperience = Player(Index).Experience
    End If
End Function
Function GetPlayerGender(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerGender = Player(Index).Gender
    End If
End Function
Function GetPlayerGuild(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerGuild = Player(Index).Guild
    End If
End Function
Function GetPlayerHP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerHP = Player(Index).HP
    End If
End Function
Function GetPlayerIntelligence(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerIntelligence = World.StatIntelligence
    End If
End Function
Function GetPlayerInvObject(ByVal Index As Long, ByVal InvIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers And InvIndex >= 1 And InvIndex <= 20 Then
        GetPlayerInvObject = Player(Index).Inv(InvIndex).Object
    End If
End Function
Function GetPlayerInvValue(ByVal Index As Long, ByVal InvIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers And InvIndex >= 1 And InvIndex <= 20 Then
        GetPlayerInvValue = Player(Index).Inv(InvIndex).Value
    End If
End Function
Function GetPlayerLevel(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerLevel = Player(Index).Level
    End If
End Function
Function GetPlayerMana(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMana = Player(Index).Mana
    End If
End Function
Function GetPlayerMap(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                GetPlayerMap = .Map
            End If
        End With
    End If
End Function

Function GetPlayerMaxEnergy(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxEnergy = Player(Index).MaxEnergy
    End If
End Function
Function GetPlayerMaxHP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxHP = Player(Index).MaxHP
    End If
End Function
Function GetPlayerMaxMana(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxMana = Player(Index).MaxMana
    End If
End Function
Function GetPlayerSprite(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerSprite = Player(Index).Sprite
    End If
End Function
Function GetPlayerFlag(ByVal Index As Long, ByVal FlagNum As Long) As Long
    If Index >= 1 And Index <= MaxUsers And FlagNum >= 0 And FlagNum <= MaxPlayerFlags Then
        GetPlayerFlag = Player(Index).Flag(FlagNum)
    End If
End Function
Sub SetPlayerFlag(ByVal Index As Long, ByVal FlagNum As Long, ByVal Value As Long)
    If Index >= 1 And Index <= MaxUsers And FlagNum >= 0 And FlagNum <= MaxPlayerFlags Then
        If Value < 0 Then
            Player(Index).Flag(FlagNum) = 0
            Exit Sub
        End If
        Player(Index).Flag(FlagNum) = Value
    End If
End Sub

Function GetPlayerStatus(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerStatus = Player(Index).Status
    End If
End Function
Function GetPlayerDirection(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerDirection = Player(Index).D
    End If
End Function
Function GetPlayerStrength(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerStrength = World.StatStrength
    End If
End Function
Function GetPlayerX(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerX = Player(Index).X
    End If
End Function
Function GetPlayerY(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerY = Player(Index).Y
    End If
End Function
Function GetValue(Value As Long) As Long
    GetValue = Value
End Function

Function Random(ByVal Max As Long) As Long
    Randomize
    Random = Int(Rnd * Max)
End Function

Function RunScript(Name As String) As Long
    Dim Tick As Currency
    On Error GoTo ScriptCrash
    
    If frmMain.mnuLogScripts.Checked = True Then
        PrintScript Name
        Tick = getTime
    End If

    'If ScriptRunning = False Then
        ScriptRS.Seek "=", Name
        If ScriptRS.NoMatch = False Then
            Dim StringFree As Long
            Dim StringCount As Long
            StringCount = StringPointer
            Dim MCode() As Byte
            MCode() = StrConv(ScriptRS!Data, vbFromUnicode)
            ScriptRunning = True
            LastScript2 = Name
            RunScript = RunASMScript(MCode(0), FunctionTable(0), Parameter(0))
            ScriptRunning = False
            For StringFree = StringCount To StringPointer - 1
                SysFreeString StringStack(StringFree)
            Next StringFree
            StringPointer = StringCount
        End If
    'End If
    
    If frmMain.mnuLogScripts.Checked = True Then
        Tick = (getTime - Tick)
        If Tick > 0 Then
            PrintScript Name + " (" + CStr(Tick) + ")"
        End If
    End If

    Exit Function

ScriptCrash:
    SendToGods Chr$(56) + Chr$(7) + "WARNING:  Server Crashed on Script " + Name
    PrintLog "WARNING:  Server Crashed on Script " + Name
    PrintDebug "WARNING:  Server Crashed on Script " + Name
    ScriptRunning = False
End Function
Sub SetPlayerEnergy(ByVal Index As Long, ByVal Energy As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If Energy > 255 Then Energy = 255
                If Energy < 0 Then Energy = 0
                .Energy = Energy
                SendSocket Index, Chr$(47) + Chr$(Energy)
            End If
        End With
    End If
End Sub
Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                .Name = StrConv(Name, vbUnicode)
                SendAll Chr$(64) + Chr$(Index) + .Name
            End If
        End With
    End If
End Sub

Sub SetPlayerMana(ByVal Index As Long, ByVal Mana As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If Mana > 255 Then Mana = 255
                If Mana < 0 Then Mana = 0
                .Mana = Mana
                SendSocket Index, Chr$(48) + Chr$(Mana)
            End If
        End With
    End If
End Sub


Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If HP > 255 Then HP = 255
                If HP < 1 Then HP = 1
                .HP = HP
                SendSocket Index, Chr$(46) + Chr$(HP)
            End If
        End With
    End If
End Sub
Sub SetPlayerBank(ByVal Index As Long, ByVal Bank As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                .Bank = Bank
            End If
        End With
    End If
End Sub

Sub SetPlayerStatus(ByVal Index As Long, ByVal Status As Long)
    If Index >= 1 And Index <= MaxUsers And Status >= 0 And Status <= 100 Then
        With Player(Index)
            If .Mode = modePlaying Then
                .Status = Status
                SendToMap .Map, Chr$(91) + Chr$(Index) + Chr$(Status)
            End If
        End With
    End If
End Sub
Sub SetPlayerGuild(ByVal Index As Long, ByVal GuildIndex As Long)
    If Index >= 1 And Index <= MaxUsers And GuildIndex >= 0 And GuildIndex <= 255 Then
        With Player(Index)
            If GuildIndex > 0 Then
                If .Guild <> GuildIndex Then
                    .Guild = GuildIndex
                    If Guild(GuildIndex).Sprite > 0 Then
                        .Sprite = Guild(GuildIndex).Sprite
                        SendToMap .Map, Chr$(63) + Chr$(Index) + DoubleChar$(CLng(.Sprite))
                    End If
                    SendSocket Index, Chr$(72) + Chr$(GuildIndex)    'Change guild
                    SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(GuildIndex)    'Player changed guild
                End If
            Else
                If .Guild > 0 Then
                    If Guild(.Guild).Sprite > 0 Then
                        .Sprite = .Class * 2 + .Gender - 1
                        SendToMap .Map, Chr$(63) + Chr$(Index) + DoubleChar$(CLng(.Sprite))
                    End If
                    .Guild = 0
                    SendSocket Index, Chr$(72) + Chr$(0)    'Change guild
                    SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(0)    'Player changed guild
                End If
            End If
        End With
    End If
End Sub

Sub SetGuildBank(ByVal Index As Long, ByVal Bank As Long)
    If Index >= 1 And Index <= 255 Then
        With Guild(Index)
            If .Name <> "" Then
                .Bank = Bank
                GuildRS.Bookmark = .Bookmark
                GuildRS.Edit
                GuildRS!Bank = Bank
                GuildRS.Update
            End If
        End With
    End If
End Sub
Sub PlayerWarp(ByVal Index As Long, ByVal Map As Long, ByVal X As Long, ByVal Y As Long)
    If Index >= 1 And Index <= MaxUsers And Map >= 1 And Map <= MaxMaps And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        With Player(Index)
            If .Mode = modePlaying Then
                If Not .Map = Map Then
                    ScriptRunning = False
                    Partmap (Index)
                    ScriptRunning = True
                    .Map = Map
                    .X = X
                    .Y = Y
                    ScriptRunning = False
                    JoinMap (Index)
                    ScriptRunning = True
                Else
                    .X = X
                    .Y = Y
                    MapWarp (Index)
                End If
            End If
        End With
    End If
End Sub
Sub DeleteString(ByVal StPointer As Long)
End Sub
Function StrVal(ByVal String1 As String) As Long
    On Error Resume Next
    StrVal = Int(Val(StrConv(String1, vbUnicode)))
    On Error GoTo 0
End Function

Sub GivePlayerExp(ByVal Index As Long, ByVal Experience As Long)
    If Index >= 1 And Index <= MaxUsers And Experience <= 50000 Then
        GainExp Index, Experience
        SendSocket Index, Chr$(60) + QuadChar(Player(Index).Experience)
    End If
End Sub

Sub GiveEliteExp(ByVal Index As Long, ByVal Experience As Long)
    If Index >= 1 And Index <= MaxUsers And Experience <= 50000 Then
        GainEliteExp Index, Experience
        SendSocket Index, Chr$(60) + QuadChar(Player(Index).Experience)
    End If
End Sub
Function GetObjectName(ByVal ObjectNum As Long) As Long
    If ObjectNum >= 1 And ObjectNum <= MaxObjects Then
        GetObjectName = NewString(Object(ObjectNum).Name)
    End If
End Function

Function GetObjectData(ByVal ObjectNum As Long, ByVal DataNum As Long) As Long
    If ObjectNum >= 1 And ObjectNum <= MaxObjects Then
        If DataNum >= 0 And DataNum <= 3 Then
            GetObjectData = Object(ObjectNum).Data(DataNum)
        End If
    End If
End Function

Function GetObjectType(ByVal ObjectNum As Long) As Long
    If ObjectNum >= 1 And ObjectNum <= MaxObjects Then
        GetObjectType = Object(ObjectNum).Type
    End If
End Function

Sub DisplayObjDur(ByVal Index As Long, ByVal ObjectNum As Long)
    Dim Percent As Single, St As String, MsgColor As Long
    Dim Display As Boolean
    Select Case Object(Player(Index).Inv(ObjectNum).Object).Type
    Case 1, 2, 3, 4, 8
        Display = True
    Case Else
        Display = False
    End Select
    If Display = True Then
        Percent = Player(Index).Inv(ObjectNum).Value / (Object(Player(Index).Inv(ObjectNum).Object).Data(0) * 10)
        Percent = Int(Percent * 100)
        If Percent > 100 Then Percent = 100
        If Percent <= 5 Then
            St = "Your " + Object(Player(Index).Inv(ObjectNum).Object).Name + " is about to break!"
            MsgColor = 2
        Else
            St = "Your " + Object(Player(Index).Inv(ObjectNum).Object).Name + " is at " + CStr(Percent) + "% durability."
            MsgColor = 14
        End If
        SendSocket Index, Chr$(56) + Chr$(MsgColor Mod 16) + St
    Else
        St = "This is an invalid object or no object."
        MsgColor = 2
        SendSocket Index, Chr$(56) + Chr$(MsgColor Mod 16) + St
    End If
End Sub

Sub SetInvObjectVal(ByVal Index As Long, ByVal InvSlot As Long, ByVal NewVal As Long)
    If Index >= 1 And Index <= MaxUsers And InvSlot >= 1 And InvSlot <= 20 Then
        Player(Index).Inv(InvSlot).Value = NewVal
    End If
End Sub

Sub PlayCustomWav(ByVal Index As Long, ByVal SoundNum As Long)
    If Index >= 1 And Index <= MaxUsers And SoundNum <= 255 And SoundNum >= 1 Then
        SendSocket Index, Chr$(96) + Chr$(SoundNum)
    End If
End Sub

Function GetPlayerArmor(ByVal Index As Long, ByVal Damage As Long) As Long
    If Index >= 1 And Index <= MaxUsers And Damage <= 255 And Damage >= 1 Then
        GetPlayerArmor = PlayerArmor(Index, Damage)
    End If
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerDamage = PlayerDamage(Index)
    End If
End Function
Sub CreateTileEffect(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Sprite As Long, ByVal Speed As Long, ByVal TotalFrames As Long, ByVal LoopCount As Long, ByVal EndSound As Long)
    If Index >= 1 And Index <= MaxMaps Then
        If X < 0 Then X = 0
        If X > 11 Then X = 11
        If Y < 0 Then Y = 0
        If Y > 11 Then Y = 11
        SendToMap Index, Chr$(99) + Chr$(1) + Chr$(X) + Chr$(Y) + Chr$(Sprite) + DoubleChar(Speed) + Chr$(TotalFrames) + Chr$(LoopCount) + Chr$(EndSound)
    End If
End Sub
Sub CreateCharacterEffect(ByVal Index As Long, ByVal Player As Long, ByVal Sprite As Long, ByVal Speed As Long, ByVal TotalFrames As Long, ByVal LoopCount As Long, ByVal EndSound As Long)
    If Index >= 1 And Index <= MaxMaps Then
        SendToMap Index, Chr$(99) + Chr$(2) + Chr$(Player) + Chr$(Sprite) + DoubleChar(Speed) + Chr$(TotalFrames) + Chr$(LoopCount) + Chr$(EndSound)
    End If
End Sub
Sub CreateMonsterEffect(ByVal Index As Long, ByVal Player As Long, ByVal Monster As Long, ByVal Sprite As Long, ByVal Speed As Long, ByVal TotalFrames As Long, ByVal EndSound As Long)
    If Index >= 1 And Index <= MaxMaps Then
        SendToMap Index, Chr$(99) + Chr$(3) + Chr$(Player) + Chr$(Monster) + Chr$(Sprite) + DoubleChar(Speed) + Chr$(TotalFrames) + Chr$(EndSound)
    End If
End Sub
Sub CreatePlayerEffect(ByVal Index As Long, ByVal SourcePlayer As Long, ByVal TargetPlayer As Long, ByVal Sprite As Long, ByVal Speed As Long, ByVal TotalFrames As Long, ByVal EndSound As Long)
    If Index >= 1 And Index <= MaxMaps Then
        SendToMap Index, Chr$(99) + Chr$(4) + Chr$(SourcePlayer) + Chr$(TargetPlayer) + Chr$(Sprite) + DoubleChar(Speed) + Chr$(TotalFrames) + Chr$(EndSound)
    End If
End Sub
Function GetPlayerMagicArmor(ByVal Index As Long, ByVal Damage As Long) As Long
    If Index >= 1 And Index <= MaxUsers And Damage <= 255 And Damage >= 1 Then
        GetPlayerMagicArmor = MagicArmor(Index, Damage)
    End If
End Function
Function ScriptMagicAttackPlayer(ByVal Index As Long, ByVal Target As Long, ByVal Damage As Long) As Long
    If Index >= 1 And Index <= MaxUsers And Target >= 1 And Target <= MaxUsers Then
        If Player(Index).Mode = modePlaying And Player(Target).Mode = modePlaying And Player(Target).IsDead = False Then
            If Damage < 0 Then Damage = 0
            If Damage > 255 Then Damage = 255
            ScriptRunning = False
            MagicAttackPlayer Index, Target, Damage
            ScriptRunning = True
            ScriptMagicAttackPlayer = True
        End If
    End If
End Function
 
Function ScriptMagicAttackMonster(ByVal Index As Long, ByVal MonsterIndex As Long, ByVal Damage As Long) As Long
    If Index >= 1 And Index <= MaxUsers And MonsterIndex >= 0 And MonsterIndex <= MaxMonsters Then
        If Player(Index).Mode = modePlaying Then
            Dim MapNum As Long
            MapNum = Player(Index).Map
            If Map(MapNum).Monster(MonsterIndex).Monster > 0 Then
                If Damage < 0 Then Damage = 0
                If Damage > 255 Then Damage = 255
                ScriptRunning = False
                MagicAttackMonster Index, MonsterIndex, Damage
                ScriptRunning = True
                ScriptMagicAttackMonster = True
            End If
        End If
    End If
End Function

Sub ScriptResetMap(ByVal Map As Long)
    If Map >= 1 And Map <= MaxMaps Then
        ResetMap Map
    End If
End Sub

Sub CreateMapFloatText(ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal Message As String, ByVal MsgColor As Long)
    MsgColor = MsgColor Mod 16
    If X >= 0 And Y >= 0 And X <= 11 And Y <= 11 Then
        SendToMap Map, Chr$(112) + Chr$(MsgColor) + Chr$(X) + Chr$(Y) + StrConv(Message, vbUnicode)
    End If
End Sub

Sub CreateMapStaticText(ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal Message As String, ByVal MsgColor As Long)
    MsgColor = MsgColor Mod 16
    If X >= 0 And Y >= 0 And X <= 11 And Y <= 11 Then
        SendToMap Map, Chr$(148) + Chr$(MsgColor) + Chr$(X) + Chr$(Y) + StrConv(Message, vbUnicode)
    End If
End Sub

Sub CreatePlayerFloatText(ByVal Index As Long, ByVal Message As String, ByVal MsgColor As Long)
    MsgColor = MsgColor Mod 16
    SendToMap Player(Index).Map, Chr$(112) + Chr$(MsgColor) + Chr$(Player(Index).X) + Chr$(Player(Index).Y) + StrConv(Message, vbUnicode)
End Sub

Function GetPlayerGuildRank(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerGuildRank = Player(Index).GuildRank
    End If
End Function

Sub CreatePlayerProjectile(ByVal Index As Long, ByVal Direction As Long, ByVal ProjectileType As Long, ByVal Damage As Long)
    If Index >= 1 And Index <= MaxUsers Then
        If Direction >= 0 And Direction <= 3 Then
        
            If Damage < 0 Then Damage = 0
            If Damage > 255 Then Damage = 255
            
            Dim DamageArray As Long
            DamageArray = FindProjectileDamageSlot(Index)
            Player(Index).ProjectileDamage(DamageArray).Live = True
            Player(Index).ProjectileDamage(DamageArray).Damage = Damage
            Player(Index).ProjectileDamage(DamageArray).ShootTime = getTime
            
            SendToMap Player(Index).Map, Chr$(99) + Chr$(6) + Chr$(Direction) + Chr$(Player(Index).X) + Chr$(Player(Index).Y) + Chr$(ProjectileType) + Chr$(Index) + Chr$(DamageArray)
        End If
    End If
End Sub

Sub CreatePlayerMagicProjectile(ByVal Index As Long, ByVal Direction As Long, ByVal ProjectileType As Long, ByVal Damage As Long)
    If Index >= 1 And Index <= MaxUsers Then
        If Direction >= 0 And Direction <= 3 Then
        
            If Damage < 0 Then Damage = 0
            If Damage > 255 Then Damage = 255
            
            Dim DamageArray As Long
            DamageArray = FindProjectileDamageSlot(Index)
            Player(Index).ProjectileDamage(DamageArray).Live = True
            Player(Index).ProjectileDamage(DamageArray).Damage = Damage
            Player(Index).ProjectileDamage(DamageArray).ShootTime = getTime
            
            SendToMap Player(Index).Map, Chr$(99) + Chr$(7) + Chr$(Direction) + Chr$(Player(Index).X) + Chr$(Player(Index).Y) + Chr$(ProjectileType) + Chr$(Index) + Chr$(DamageArray)
        End If
    End If
End Sub

Function GetPlayerIsDead(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerIsDead = Player(Index).IsDead
    End If
End Function

Sub SetPlayerIsDead(ByVal Index As Long, ByVal IsDead As Long)
    If Index >= 1 And Index <= MaxUsers Then
        If IsDead > 0 Then
            Player(Index).IsDead = 1
            Player(Index).DeadTick = getTime + 15000
            SendAll Chr$(120) + Chr$(Index) + Chr$(1)
        Else
            Player(Index).IsDead = 0
            SendAll Chr$(120) + Chr$(Index)
        End If
    End If
End Sub

Sub SetItemSuffix(ByVal Index As Long, ByVal Slot As Long, ByVal Suffix As Long)
    If Slot >= 1 And Slot <= 20 Then
        If Player(Index).Mode = modePlaying Then
            If Player(Index).Inv(Slot).Object > 0 Then
                Select Case Object(Player(Index).Inv(Slot).Object).Type
                Case 1, 2, 3, 4, 7, 8, 10
                    Player(Index).Inv(Slot).ItemSuffix = Suffix
                    SendSocket Index, Chr$(17) + Chr$(Slot) + DoubleChar$(CInt(Player(Index).Inv(Slot).Object)) + QuadChar(Player(Index).Inv(Slot).Value) + Chr$(Player(Index).Inv(Slot).ItemPrefix) + Chr$(Player(Index).Inv(Slot).ItemSuffix)    'New Inv Obj
                End Select
            End If
        End If
    End If
End Sub

Function GetItemSuffix(ByVal Index As Long, ByVal Slot As Long) As Long
    GetItemSuffix = 0
    If Slot >= 1 And Slot <= 20 Then
        If Player(Index).Mode = modePlaying Then
            If Player(Index).Inv(Slot).Object > 0 Then
                Select Case Object(Player(Index).Inv(Slot).Object).Type
                Case 1, 2, 3, 4, 7, 8, 10
                    GetItemSuffix = Player(Index).Inv(Slot).ItemSuffix
                End Select
            End If
        End If
    End If
End Function

Sub SetEquippedItemSuffix(ByVal Index As Long, ByVal Slot As Long, ByVal Suffix As Long)
    If Slot >= 1 And Slot <= 5 Then
        If Player(Index).Mode = modePlaying Then
            If Player(Index).EquippedObject(Slot).Object > 0 Then
                Select Case Object(Player(Index).EquippedObject(Slot).Object).Type
                Case 1, 2, 3, 4, 7, 8, 10
                    Player(Index).EquippedObject(Slot).ItemSuffix = Suffix
                    SendSocket Index, Chr$(115) + DoubleChar$(CInt(Player(Index).EquippedObject(Slot).Object)) + QuadChar(Player(Index).EquippedObject(Slot).Value) + Chr$(Player(Index).EquippedObject(Slot).ItemPrefix) + Chr$(Player(Index).EquippedObject(Slot).ItemSuffix)    'New Inv Obj
                    CalculateStats Index
                End Select
            End If
        End If
    End If
End Sub

Function GetEquippedItemSuffix(ByVal Index As Long, ByVal Slot As Long) As Long
    GetEquippedItemSuffix = 0
    If Slot >= 1 And Slot <= 5 Then
        If Player(Index).Mode = modePlaying Then
            If Player(Index).EquippedObject(Slot).Object > 0 Then
                Select Case Object(Player(Index).EquippedObject(Slot).Object).Type
                Case 1, 2, 3, 4, 7, 8, 10
                    GetEquippedItemSuffix = Player(Index).EquippedObject(Slot).ItemSuffix
                End Select
            End If
        End If
    End If
End Function

Function MonsterAttackPlayer(ByVal TheMap As Long, ByVal Monster As Long, ByVal Index As Long, ByVal Damage As Long) As Long
    Dim C As Long
    If Index >= 1 And Index <= MaxUsers And Monster >= 0 And Monster <= MaxMonsters Then
        If Player(Index).Mode = modePlaying And Player(Index).IsDead = False And Player(Index).Map = TheMap Then
            If TheMap >= 1 And TheMap <= MaxMaps Then
                If Map(TheMap).Monster(Monster).Monster > 0 Then
                    If Damage < 0 Then Damage = 0
                    If Damage > 255 Then Damage = 255
                    C = PlayerArmor(Index, Damage)
                    If C < 0 Then C = 0
                    If C > 255 Then C = 255
                    SendSocket Index, Chr$(50) + Chr$(0) + Chr$(Monster) + Chr$(C)
                    SendToMap TheMap, Chr$(111) + Chr$(12) + Chr$(C) + Chr$(Player(Index).X) + Chr$(Player(Index).Y)
                    SendToMap TheMap, Chr$(41) + Chr$(Monster)
                    With Player(Index)
                        If C >= .HP Then
                            Parameter(0) = Map(.Map).Monster(Monster).Monster
                            Parameter(1) = Monster
                            Parameter(2) = Index
                            ScriptRunning = False
                            If RunScript("MONSTERKILL" + CStr(Map(.Map).Monster(Monster).Monster)) = 0 Then
                                'Player Died
                                Map(.Map).Monster(Monster).Target = 0
                                SendSocket Index, Chr$(53) + DoubleChar$(CLng(Map(.Map).Monster(Monster).Monster))    'Monster Killed You
                                SendAllBut Index, Chr$(62) + Chr$(Index) + DoubleChar$(CLng(Map(.Map).Monster(Monster).Monster))    'Player was killed by monster
                                PlayerDied Index, -1
                            End If
                            ScriptRunning = True
                        Else
                            .HP = .HP - C
                            SendHPUpdate Index
                        End If
                    End With
                End If
            End If
        End If
    End If
End Function

Function MonsterMagicAttackPlayer(ByVal TheMap As Long, ByVal Monster As Long, ByVal Index As Long, ByVal Damage As Long) As Long
    Dim C As Long
    If Index >= 1 And Index <= MaxUsers And Monster >= 0 And Monster <= MaxMonsters Then
        If Player(Index).Mode = modePlaying And Player(Index).IsDead = False And Player(Index).Map = TheMap Then
            If TheMap >= 1 And TheMap <= MaxMaps Then
                If Map(TheMap).Monster(Monster).Monster > 0 Then
                    If Damage < 0 Then Damage = 0
                    If Damage > 255 Then Damage = 255
                    C = MagicArmor(Index, Damage)
                    If C < 0 Then C = 0
                    If C > 255 Then C = 255
                    SendSocket Index, Chr$(50) + Chr$(0) + Chr$(Monster) + Chr$(C)
                    SendToMap TheMap, Chr$(111) + Chr$(12) + Chr$(C) + Chr$(Player(Index).X) + Chr$(Player(Index).Y)
                    SendToMap TheMap, Chr$(41) + Chr$(Monster)
                    With Player(Index)
                        If C >= .HP Then
                            Parameter(0) = Map(.Map).Monster(Monster).Monster
                            Parameter(1) = Monster
                            Parameter(2) = Index
                            ScriptRunning = False
                            If RunScript("MONSTERKILL" + CStr(Map(.Map).Monster(Monster).Monster)) = 0 Then
                                'Player Died
                                Map(.Map).Monster(Monster).Target = 0
                                SendSocket Index, Chr$(53) + DoubleChar$(CLng(Map(.Map).Monster(Monster).Monster))    'Monster Killed You
                                SendAllBut Index, Chr$(62) + Chr$(Index) + DoubleChar$(CLng(Map(.Map).Monster(Monster).Monster))    'Player was killed by monster
                                PlayerDied Index, -1
                            End If
                            ScriptRunning = True
                        Else
                            .HP = .HP - C
                            SendHPUpdate Index
                        End If
                    End With
                End If
            End If
        End If
    End If
End Function

Sub SetItemPrefix(ByVal Index As Long, ByVal Slot As Long, ByVal Prefix As Long)
    If Slot >= 1 And Slot <= 20 Then
        If Player(Index).Mode = modePlaying Then
            If Player(Index).Inv(Slot).Object > 0 Then
                Select Case Object(Player(Index).Inv(Slot).Object).Type
                Case 1, 2, 3, 4, 7, 8, 10
                    Player(Index).Inv(Slot).ItemPrefix = Prefix
                    SendSocket Index, Chr$(17) + Chr$(Slot) + DoubleChar$(CInt(Player(Index).Inv(Slot).Object)) + QuadChar(Player(Index).Inv(Slot).Value) + Chr$(Player(Index).Inv(Slot).ItemPrefix) + Chr$(Player(Index).Inv(Slot).ItemSuffix)    'New Inv Obj
                End Select
            End If
        End If
    End If
End Sub

Function GetItemPrefix(ByVal Index As Long, ByVal Slot As Long) As Long
    GetItemPrefix = 0
    If Slot >= 1 And Slot <= 20 Then
        If Player(Index).Mode = modePlaying Then
            If Player(Index).Inv(Slot).Object > 0 Then
                Select Case Object(Player(Index).Inv(Slot).Object).Type
                Case 1, 2, 3, 4, 7, 8, 10
                    GetItemPrefix = Player(Index).Inv(Slot).ItemPrefix
                End Select
            End If
        End If
    End If
End Function

Sub SetEquippedItemPrefix(ByVal Index As Long, ByVal Slot As Long, ByVal Prefix As Long)
    If Slot >= 1 And Slot <= 5 Then
        If Player(Index).Mode = modePlaying Then
            If Player(Index).EquippedObject(Slot).Object > 0 Then
                Select Case Object(Player(Index).EquippedObject(Slot).Object).Type
                Case 1, 2, 3, 4, 7, 8, 10
                    Player(Index).EquippedObject(Slot).ItemPrefix = Prefix
                    SendSocket Index, Chr$(115) + DoubleChar$(CInt(Player(Index).EquippedObject(Slot).Object)) + QuadChar(Player(Index).EquippedObject(Slot).Value) + Chr$(Player(Index).EquippedObject(Slot).ItemPrefix) + Chr$(Player(Index).EquippedObject(Slot).ItemSuffix)    'New Inv Obj
                    CalculateStats Index
                End Select
            End If
        End If
    End If
End Sub

Function GetEquippedItemPrefix(ByVal Index As Long, ByVal Slot As Long) As Long
    GetEquippedItemPrefix = 0
    If Slot >= 1 And Slot <= 5 Then
        If Player(Index).Mode = modePlaying Then
            If Player(Index).EquippedObject(Slot).Object > 0 Then
                Select Case Object(Player(Index).EquippedObject(Slot).Object).Type
                Case 1, 2, 3, 4, 7, 8, 10
                    GetEquippedItemPrefix = Player(Index).EquippedObject(Slot).ItemPrefix
                End Select
            End If
        End If
    End If
End Function

Function GetPrefixName(ByVal PrefixNum As Long) As Long
    If PrefixNum >= 1 And PrefixNum <= 255 Then
        GetPrefixName = NewString(ItemPrefix(PrefixNum).Name)
    End If
End Function

Function GetSuffixName(ByVal SuffixNum As Long) As Long
    If SuffixNum >= 1 And SuffixNum <= 255 Then
        GetSuffixName = NewString(ItemSuffix(SuffixNum).Name)
    End If
End Function

Sub SetPlayerMaxHP(ByVal Index As Long, ByVal MaxHP As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If MaxHP > 255 Then MaxHP = 255
                If MaxHP < 1 Then MaxHP = 1
                .MaxHP = MaxHP
            End If
        End With
    End If
End Sub

Sub SetPlayerMaxEnergy(ByVal Index As Long, ByVal MaxEnergy As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If MaxEnergy > 255 Then MaxEnergy = 255
                If MaxEnergy < 1 Then MaxEnergy = 1
                .MaxEnergy = MaxEnergy
            End If
        End With
    End If
End Sub

Sub SetPlayerMaxMana(ByVal Index As Long, ByVal MaxMana As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If MaxMana > 255 Then MaxMana = 255
                If MaxMana < 1 Then MaxMana = 1
                .MaxMana = MaxMana
            End If
        End With
    End If
End Sub

Function GetPlayerConcentration(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerConcentration = World.StatConcentration
    End If
End Function


Function GetPlayerConstitution(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerConstitution = World.StatConstitution
    End If
End Function

Function GetPlayerStamina(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerStamina = World.StatStamina
    End If
End Function

Function GetPlayerWisdom(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerWisdom = World.StatWisdom
    End If
End Function

Function GetPlayerBaseStrength(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBaseStrength = World.StatStrength
    End If
End Function

Function GetPlayerBaseEndurance(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBaseEndurance = World.StatEndurance
    End If
End Function

Function GetPlayerBaseAgility(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBaseAgility = statPlayerAgility
    End If
End Function

Function GetPlayerBaseIntelligence(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBaseIntelligence = World.StatIntelligence
    End If
End Function

Function GetPlayerBaseConcentration(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBaseConcentration = World.StatConcentration
    End If
End Function

Function GetPlayerBaseConstitution(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBaseConstitution = World.StatConstitution
    End If
End Function

Function GetPlayerBaseStamina(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBaseStamina = World.StatStamina
    End If
End Function

Function GetPlayerBaseWisdom(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBaseWisdom = World.StatWisdom
    End If
End Function

Sub SetPlayerStrength(ByVal Index As Long, ByVal Strength As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then

            End If
        End With
    End If
End Sub

Sub SetPlayerEndurance(ByVal Index As Long, ByVal Endurance As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then

            End If
        End With
    End If
End Sub

Sub SetPlayerAgility(ByVal Index As Long, ByVal Agility As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then

            End If
        End With
    End If
End Sub

Sub SetPlayerIntelligence(ByVal Index As Long, ByVal Intelligence As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
            
            End If
        End With
    End If
End Sub

Sub SetPlayerConcentration(ByVal Index As Long, ByVal Concentration As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then

            End If
        End With
    End If
End Sub

Sub SetPlayerConstitution(ByVal Index As Long, ByVal Constitution As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then

            End If
        End With
    End If
End Sub

Sub SetPlayerStamina(ByVal Index As Long, ByVal Stamina As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then

            End If
        End With
    End If
End Sub

Sub SetPlayerWisdom(ByVal Index As Long, ByVal Wisdom As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then

            End If
        End With
    End If
End Sub

Sub SetPlayerClass(ByVal Index As Long, ByVal Class As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If Class >= 1 And Class <= NumClasses Then
                    .Class = Class
                    CalculateStats Index
                    SendSocket Index, Chr$(145) + Chr$(.Class)
                    SetPlayerSprite Index, 0
                End If
            End If
        End With
    End If
End Sub

Sub SetPlayerDirection(ByVal Index As Long, ByVal Direction As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If Direction >= 0 And Direction <= 3 Then
                    .D = Direction
                    SendToMap .Map, Chr$(146) + Chr$(Index) + Chr$(.D)
                End If
            End If
        End With
    End If
End Sub

Sub ScriptCalculateStats(ByVal Index As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                ScriptRunning = False
                CalculateStats Index
                ScriptRunning = True
            End If
        End With
    End If
End Sub

Sub SetOutdoorLight(ByVal Light As Long)
    If Light < 0 Or Light > 255 Then Light = 0

    OutdoorLight = CByte(Light)
    SendAll Chr$(143) + Chr$(OutdoorLight)
End Sub

Function ReadStrScript(ByVal Filename As String, ByVal Heading As String, ByVal Name As String, Optional ByVal Default As String = "0") As String
    Dim filename2 As String
    filename2 = StrConv(Filename, vbUnicode)
    If Len(filename2) < 16 Then
        Dim R As Integer, buffer As String, tempname As String
        buffer = String$(255, 0)
        tempname = StrConv(Name, vbUnicode)
        R = GetPrivateProfileString(StrConv(Heading, vbUnicode), tempname, StrConv(Default, vbUnicode), buffer, 255, App.Path + "\scriptini\" + StrConv(Filename, vbUnicode) + ".ini")
        ReadStrScript = StrConv(Left$(buffer, R), vbFromUnicode)
    End If
End Function

Function ReadIntScript(ByVal Filename As String, ByVal Heading As String, ByVal Name As String, Optional ByVal Default As Long = 0) As Long
    Dim filename2 As String
    filename2 = StrConv(Filename, vbUnicode)
    If Len(filename2) < 16 Then
        ReadIntScript = GetPrivateProfileInt&(StrConv(Heading, vbUnicode), StrConv(Name, vbUnicode), Default, App.Path + "\scriptini\" + StrConv(Filename, vbUnicode) + ".ini")
    End If
End Function

Sub WriteStrScript(ByVal Filename As String, ByVal Heading As String, ByVal Name As String, ByVal Data As String)
    Dim filename2 As String
    filename2 = StrConv(Filename, vbUnicode)
    If Len(filename2) < 16 Then
        WriteString StrConv(Heading, vbUnicode), StrConv(Name, vbUnicode), StrConv(Data, vbUnicode), "\scriptini\" + filename2 + ".ini"
    End If
End Sub

Function Divide(ByVal Numerator As Long, ByVal Denominator As Long) As Long
    If Not Denominator = 0 Then
        Divide = Numerator / Denominator
    End If
End Function

Function GetMapName(ByVal TheMap As Long) As Long
    If TheMap >= 1 And TheMap <= MaxMaps Then
        With Map(TheMap)
            GetMapName = NewString(.Name)
        End With
    End If
End Function

Function GetBootLocationMap(ByVal TheMap As Long) As Long
    If TheMap >= 1 And TheMap <= MaxMaps Then
        With Map(TheMap)
            GetBootLocationMap = .BootLocation.Map
        End With
    End If
End Function

Function GetBootLocationX(ByVal TheMap As Long) As Long
    If TheMap >= 1 And TheMap <= MaxMaps Then
        With Map(TheMap)
            GetBootLocationX = .BootLocation.X
        End With
    End If
End Function

Function GetBootLocationY(ByVal TheMap As Long) As Long
    If TheMap >= 1 And TheMap <= MaxMaps Then
        With Map(TheMap)
            GetBootLocationY = .BootLocation.Y
        End With
    End If
End Function

Function GetMapIsFriendly(ByVal TheMap As Long) As Long
    If TheMap >= 1 And TheMap <= MaxMaps Then
        With Map(TheMap)
            If ExamineBit(.flags, 0) = True Then
                GetMapIsFriendly = True
            Else
                GetMapIsFriendly = False
            End If
        End With
    End If
End Function

Function GetMapIsPK(ByVal TheMap As Long) As Long
    If TheMap >= 1 And TheMap <= MaxMaps Then
        With Map(TheMap)
            If ExamineBit(.flags, 6) = True Then
                GetMapIsPK = True
            Else
                GetMapIsPK = False
            End If
        End With
    End If
End Function

Function GetMapIsArena(ByVal TheMap As Long) As Long
    If TheMap >= 1 And TheMap <= MaxMaps Then
        With Map(TheMap)
            If ExamineBit(.flags, 7) = True Then
                GetMapIsArena = True
            Else
                GetMapIsArena = False
            End If
        End With
    End If
End Function

Function GetMapObjVal(ByVal TheMap As Long, ByVal TheObject As Long) As Long
    If TheMap >= 1 And TheMap <= MaxMaps Then
        If TheObject >= 0 And TheObject <= MaxMapObjects Then
            With Map(TheMap)
                If .Object(TheObject).Object > 0 Then
                    GetMapObjVal = .Object(TheObject).Value
                End If
            End With
        End If
    End If
End Function

Sub SetMapObjVal(ByVal TheMap As Long, ByVal TheObject As Long, ByVal TheValue As Long)
    If TheMap >= 1 And TheMap <= MaxMaps Then
        If TheObject >= 0 And TheObject <= MaxMapObjects Then
            With Map(TheMap)
                If .Object(TheObject).Object > 0 Then
                    .Object(TheObject).Value = TheValue
                End If
            End With
        End If
    End If
End Sub

Function StrLen(ByVal TheString As String) As Long
    Dim ConvertString As String
    ConvertString = StrConv(TheString, vbUnicode)
    StrLen = Len(ConvertString)
End Function

Function GetPlayerSkillLevel(ByVal Index As Long, ByVal Skill As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        If Skill > 0 And Skill <= MaxSkill Then
            GetPlayerSkillLevel = Player(Index).Skill(Skill).Level
        End If
    End If
End Function

Sub SetPlayerSkillLevel(ByVal Index As Long, ByVal Skill As Long, ByVal Level As Long)
    If Index >= 1 And Index <= MaxUsers Then
        If Skill <= MaxSkill Then
            If Level >= 0 And Level <= 255 Then
                With Player(Index)
                    .Skill(Skill).Level = Level
                    .Skill(Skill).Experience = 0
                    SendSocket Index, Chr$(119) + Chr$(2) + Chr$(Skill) + Chr$(.Skill(Skill).Level)
                End With
            End If
        End If
    End If
End Sub

Function GetPlayerMagicLevel(ByVal Index As Long, ByVal Magic As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        If Magic > 0 And Magic <= 255 Then
            GetPlayerMagicLevel = Player(Index).MagicLevel(Magic).Level
        End If
    End If
End Function

Sub SetPlayerMagicLevel(ByVal Index As Long, ByVal Magic As Long, ByVal Level As Long)
    If Index >= 1 And Index <= MaxUsers Then
        If Magic > 0 And Magic <= 255 Then
            If Level >= 0 And Level <= 255 Then
                With Player(Index)
                    .MagicLevel(Magic).Level = Level
                    .MagicLevel(Magic).Experience = 0
                    SendSocket Index, Chr$(153) + Chr$(2) + DoubleChar$(Magic) + Chr$(.MagicLevel(Magic).Level)
                End With
            End If
        End If
    End If
End Sub

Sub GivePlayerSkillExp(ByVal Index As Long, ByVal Skill As Long, ByVal Exp As Long)
    If Index >= 1 And Index <= MaxUsers Then
        If Player(Index).InUse = True Then
            With Player(Index)
                If Skill > 0 And Skill <= MaxSkill Then
                    If .Skill(Skill).Level < 255 Then
                        .Skill(Skill).Experience = .Skill(Skill).Experience + Exp
                        If .Skill(Skill).Experience >= Int(5 * .Skill(Skill).Level ^ 1.3) Then    'Skill Level
                            .Skill(Skill).Level = .Skill(Skill).Level + 1
                            .Skill(Skill).Experience = 0
                            SendSocket Index, Chr$(119) + Chr$(2) + Chr$(Skill) + Chr$(.Skill(Skill).Level)
                        Else
                            SendSocket Index, Chr$(119) + Chr$(3) + Chr$(Skill) + Chr$(.Skill(Skill).Level) + QuadChar$(.Skill(Skill).Experience)
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Sub GivePlayerMagicExp(ByVal Index As Long, ByVal Magic As Long, ByVal Exp As Long)
    If Index >= 1 And Index <= MaxUsers Then
        If Player(Index).InUse = True Then
            With Player(Index)
                If Magic > 0 And Magic <= 255 Then
                    If .MagicLevel(Magic).Level < 255 Then
                        .MagicLevel(Magic).Experience = .MagicLevel(Magic).Experience + Exp
                        If .MagicLevel(Magic).Experience >= Int(5 * .MagicLevel(Magic).Level ^ 1.3) Then    'Skill Level
                            .MagicLevel(Magic).Level = .MagicLevel(Magic).Level + 1
                            .MagicLevel(Magic).Experience = 0
                            SendSocket Index, Chr$(153) + Chr$(2) + DoubleChar$(Magic) + Chr$(.MagicLevel(Magic).Level)
                        Else
                            SendSocket Index, Chr$(153) + Chr$(3) + DoubleChar$(Magic) + Chr$(.MagicLevel(Magic).Level) + QuadChar$(.MagicLevel(Magic).Experience)
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Sub InitFunctionTable()
    FunctionTable(0) = GetValue(AddressOf DeleteString)
    FunctionTable(1) = GetValue(AddressOf StrCat)
    FunctionTable(2) = GetValue(AddressOf StrCmp)
    FunctionTable(3) = GetValue(AddressOf StrFormat)

    FunctionTable(4) = GetValue(AddressOf Random)

    FunctionTable(5) = GetValue(AddressOf GetPlayerAccess)
    FunctionTable(6) = GetValue(AddressOf GetPlayerMap)
    FunctionTable(7) = GetValue(AddressOf GetPlayerX)
    FunctionTable(8) = GetValue(AddressOf GetPlayerY)
    FunctionTable(9) = GetValue(AddressOf GetPlayerSprite)
    FunctionTable(10) = GetValue(AddressOf GetPlayerClass)
    FunctionTable(11) = GetValue(AddressOf GetPlayerGender)
    FunctionTable(12) = GetValue(AddressOf GetPlayerHP)
    FunctionTable(13) = GetValue(AddressOf GetPlayerEnergy)
    FunctionTable(14) = GetValue(AddressOf GetPlayerMana)
    FunctionTable(15) = GetValue(AddressOf GetPlayerMaxHP)
    FunctionTable(16) = GetValue(AddressOf GetPlayerMaxEnergy)
    FunctionTable(17) = GetValue(AddressOf GetPlayerMaxMana)
    FunctionTable(18) = GetValue(AddressOf GetPlayerStrength)
    FunctionTable(19) = GetValue(AddressOf GetPlayerEndurance)
    FunctionTable(20) = GetValue(AddressOf GetPlayerIntelligence)
    FunctionTable(21) = GetValue(AddressOf GetPlayerAgility)
    FunctionTable(22) = GetValue(AddressOf GetPlayerBank)
    FunctionTable(23) = GetValue(AddressOf GetPlayerExperience)
    FunctionTable(24) = GetValue(AddressOf GetPlayerLevel)
    FunctionTable(25) = GetValue(AddressOf GetPlayerStatus)
    FunctionTable(26) = GetValue(AddressOf GetPlayerGuild)

    FunctionTable(27) = GetValue(AddressOf GetPlayerInvObject)
    FunctionTable(28) = GetValue(AddressOf GetPlayerInvValue)
    FunctionTable(29) = GetValue(AddressOf GetPlayerEquipped)

    FunctionTable(30) = GetValue(AddressOf GetPlayerName)
    FunctionTable(31) = GetValue(AddressOf GetPlayerUser)
    FunctionTable(32) = GetValue(AddressOf GetPlayerDesc)

    FunctionTable(33) = GetValue(AddressOf SetPlayerHP)
    FunctionTable(34) = GetValue(AddressOf SetPlayerEnergy)
    FunctionTable(35) = GetValue(AddressOf SetPlayerMana)


    FunctionTable(36) = GetValue(AddressOf PlayerMessage)
    FunctionTable(37) = GetValue(AddressOf PlayerWarp)

    FunctionTable(38) = GetValue(AddressOf MapMessage)

    FunctionTable(39) = GetValue(AddressOf GlobalMessage)

    FunctionTable(40) = GetValue(AddressOf GetGuildHall)
    FunctionTable(41) = GetValue(AddressOf GetGuildBank)
    FunctionTable(42) = GetValue(AddressOf GetGuildMemberCount)
    FunctionTable(43) = GetValue(AddressOf GetGuildName)

    FunctionTable(44) = GetValue(AddressOf GetMapPlayerCount)

    FunctionTable(45) = GetValue(AddressOf MapMessageAllBut)

    FunctionTable(46) = GetValue(AddressOf HasObj)
    FunctionTable(47) = GetValue(AddressOf TakeObj)
    FunctionTable(48) = GetValue(AddressOf GiveObj)

    FunctionTable(50) = GetValue(AddressOf GetMaxUsers)

    FunctionTable(51) = GetValue(AddressOf RunScript0)
    FunctionTable(52) = GetValue(AddressOf RunScript1)
    FunctionTable(53) = GetValue(AddressOf RunScript2)
    FunctionTable(54) = GetValue(AddressOf RunScript3)

    FunctionTable(55) = GetValue(AddressOf OpenDoor)

    FunctionTable(56) = GetValue(AddressOf Str)

    FunctionTable(57) = GetValue(AddressOf SetPlayerSprite)

    FunctionTable(58) = GetValue(AddressOf GetAbs)
    FunctionTable(59) = GetValue(AddressOf GetSqr)

    FunctionTable(60) = GetValue(AddressOf CanAttackPlayer)
    FunctionTable(61) = GetValue(AddressOf IsPlaying)

    FunctionTable(62) = GetValue(AddressOf AttackPlayer)
    FunctionTable(63) = GetValue(AddressOf AttackMonster)
    FunctionTable(64) = GetValue(AddressOf CanAttackMonster)

    FunctionTable(65) = GetValue(AddressOf GetMonsterType)
    FunctionTable(66) = GetValue(AddressOf GetMonsterX)
    FunctionTable(67) = GetValue(AddressOf GetMonsterY)
    FunctionTable(68) = GetValue(AddressOf GetMonsterTarget)
    FunctionTable(69) = GetValue(AddressOf SetMonsterTarget)

    FunctionTable(70) = GetValue(AddressOf GetInStr)
    FunctionTable(71) = GetValue(AddressOf SpawnObject)

    FunctionTable(72) = GetValue(AddressOf NPCSay)
    FunctionTable(73) = GetValue(AddressOf NPCTell)

    FunctionTable(74) = GetValue(AddressOf GetGuildSprite)

    FunctionTable(75) = GetValue(AddressOf ScriptTimer)

    FunctionTable(76) = GetValue(AddressOf SetPlayerGuild)

    FunctionTable(77) = GetValue(AddressOf GetFlag)
    FunctionTable(78) = GetValue(AddressOf SetFlag)

    FunctionTable(79) = GetValue(AddressOf GetPlayerFlag)
    FunctionTable(80) = GetValue(AddressOf SetPlayerFlag)

    FunctionTable(83) = GetValue(AddressOf GetObjX)
    FunctionTable(84) = GetValue(AddressOf GetObjY)
    FunctionTable(85) = GetValue(AddressOf GetObjNum)
    FunctionTable(86) = GetValue(AddressOf GetObjVal)
    FunctionTable(87) = GetValue(AddressOf DestroyObj)
    FunctionTable(88) = GetValue(AddressOf Boot_Player)
    FunctionTable(89) = GetValue(AddressOf Ban_Player)
    FunctionTable(90) = GetValue(AddressOf SetPlayerName)
    FunctionTable(91) = GetValue(AddressOf SetPlayerBank)
    FunctionTable(92) = GetValue(AddressOf SetGuildBank)
    FunctionTable(93) = GetValue(AddressOf Find_Player)
    FunctionTable(94) = GetValue(AddressOf StrVal)
    FunctionTable(95) = GetValue(AddressOf GetTileAtt)
    FunctionTable(96) = GetValue(AddressOf GetPlayerIP)
    FunctionTable(97) = GetValue(AddressOf RunScript4)
    FunctionTable(98) = GetValue(AddressOf SpawnMonster)
    FunctionTable(99) = GetValue(AddressOf SetPlayerStatus)
    FunctionTable(100) = GetValue(AddressOf GivePlayerExp)
    FunctionTable(101) = GetValue(AddressOf GetPlayerDirection)

    FunctionTable(103) = GetValue(AddressOf GetObjectName)
    FunctionTable(104) = GetValue(AddressOf GetObjectData)
    FunctionTable(105) = GetValue(AddressOf GetObjectType)
    FunctionTable(106) = GetValue(AddressOf DisplayObjDur)
    FunctionTable(107) = GetValue(AddressOf SetInvObjectVal)
    FunctionTable(108) = GetValue(AddressOf PlayCustomWav)
    FunctionTable(109) = GetValue(AddressOf GetPlayerArmor)
    FunctionTable(110) = GetValue(AddressOf CreateTileEffect)
    FunctionTable(111) = GetValue(AddressOf CreateCharacterEffect)
    FunctionTable(112) = GetValue(AddressOf CreateMonsterEffect)
    FunctionTable(113) = GetValue(AddressOf CreatePlayerEffect)
    FunctionTable(114) = GetValue(AddressOf GetPlayerMagicArmor)
    FunctionTable(115) = GetValue(AddressOf ScriptMagicAttackPlayer)
    FunctionTable(116) = GetValue(AddressOf ScriptMagicAttackMonster)
    FunctionTable(117) = GetValue(AddressOf ScriptResetMap)
    FunctionTable(118) = GetValue(AddressOf CreateMapFloatText)
    FunctionTable(119) = GetValue(AddressOf CreatePlayerFloatText)
    FunctionTable(120) = GetValue(AddressOf GetPlayerGuildRank)
    FunctionTable(121) = GetValue(AddressOf CreatePlayerProjectile)
    FunctionTable(122) = GetValue(AddressOf CreatePlayerMagicProjectile)
    FunctionTable(123) = GetValue(AddressOf GetPlayerIsDead)
    FunctionTable(124) = GetValue(AddressOf SetPlayerIsDead)
    FunctionTable(125) = GetValue(AddressOf GetMonsterDirection)
    FunctionTable(126) = GetValue(AddressOf GetPlayerDamage)

    FunctionTable(127) = GetValue(AddressOf SetItemSuffix)
    FunctionTable(128) = GetValue(AddressOf GetItemSuffix)
    FunctionTable(129) = GetValue(AddressOf SetEquippedItemSuffix)
    FunctionTable(130) = GetValue(AddressOf GetEquippedItemSuffix)
    FunctionTable(131) = GetValue(AddressOf SetItemPrefix)
    FunctionTable(132) = GetValue(AddressOf GetItemPrefix)
    FunctionTable(133) = GetValue(AddressOf SetEquippedItemPrefix)
    FunctionTable(134) = GetValue(AddressOf GetEquippedItemPrefix)
    FunctionTable(135) = GetValue(AddressOf GetPrefixName)
    FunctionTable(136) = GetValue(AddressOf GetSuffixName)

    FunctionTable(137) = GetValue(AddressOf SetPlayerMaxHP)
    FunctionTable(138) = GetValue(AddressOf SetPlayerMaxEnergy)
    FunctionTable(139) = GetValue(AddressOf SetPlayerMaxMana)
    FunctionTable(140) = GetValue(AddressOf GetPlayerConcentration)
    FunctionTable(141) = GetValue(AddressOf GetPlayerConstitution)
    FunctionTable(142) = GetValue(AddressOf GetPlayerStamina)
    FunctionTable(143) = GetValue(AddressOf GetPlayerWisdom)
    FunctionTable(144) = GetValue(AddressOf GetPlayerBaseStrength)
    FunctionTable(145) = GetValue(AddressOf GetPlayerBaseEndurance)
    FunctionTable(146) = GetValue(AddressOf GetPlayerBaseIntelligence)
    FunctionTable(147) = GetValue(AddressOf GetPlayerBaseAgility)
    FunctionTable(148) = GetValue(AddressOf GetPlayerBaseConcentration)
    FunctionTable(149) = GetValue(AddressOf GetPlayerBaseConstitution)
    FunctionTable(150) = GetValue(AddressOf GetPlayerBaseStamina)
    FunctionTable(151) = GetValue(AddressOf GetPlayerBaseWisdom)
    FunctionTable(152) = GetValue(AddressOf SetPlayerStrength)
    FunctionTable(153) = GetValue(AddressOf SetPlayerEndurance)
    FunctionTable(154) = GetValue(AddressOf SetPlayerAgility)
    FunctionTable(155) = GetValue(AddressOf SetPlayerIntelligence)
    FunctionTable(156) = GetValue(AddressOf SetPlayerConcentration)
    FunctionTable(157) = GetValue(AddressOf SetPlayerConstitution)
    FunctionTable(158) = GetValue(AddressOf SetPlayerStamina)
    FunctionTable(159) = GetValue(AddressOf SetPlayerWisdom)

    FunctionTable(160) = GetValue(AddressOf ScriptCalculateStats)

    FunctionTable(161) = GetValue(AddressOf ReadIntScript)
    FunctionTable(162) = GetValue(AddressOf ReadStrScript)
    FunctionTable(163) = GetValue(AddressOf WriteStrScript)

    FunctionTable(164) = GetValue(AddressOf SetOutdoorLight)
    '165 free
    '166 free

    FunctionTable(167) = GetValue(AddressOf SetPlayerClass)
    FunctionTable(168) = GetValue(AddressOf SetPlayerDirection)

    FunctionTable(169) = GetValue(AddressOf CreateMapStaticText)

    FunctionTable(170) = GetValue(AddressOf GetTileAtt2)
    FunctionTable(171) = GetValue(AddressOf GetTileIsVacant)
    FunctionTable(172) = GetValue(AddressOf GetTileNoDirectionalWalls)
    FunctionTable(173) = GetValue(AddressOf GetMapName)

    FunctionTable(174) = GetValue(AddressOf StrLen)
    FunctionTable(175) = GetValue(AddressOf Divide)

    FunctionTable(176) = GetValue(AddressOf GetMonsterHP)
    FunctionTable(177) = GetValue(AddressOf SetMonsterHP)
    FunctionTable(178) = GetValue(AddressOf MonsterAttackPlayer)
    FunctionTable(179) = GetValue(AddressOf MonsterMagicAttackPlayer)
    FunctionTable(180) = GetValue(AddressOf GetMapIsFriendly)
    FunctionTable(181) = GetValue(AddressOf GetMapIsPK)
    FunctionTable(182) = GetValue(AddressOf GetMapIsArena)
    FunctionTable(183) = GetValue(AddressOf GetMapObjVal)
    FunctionTable(184) = GetValue(AddressOf SetMapObjVal)
    
    FunctionTable(185) = GetValue(AddressOf GetPlayerSkillLevel)
    FunctionTable(186) = GetValue(AddressOf GivePlayerSkillExp)
    FunctionTable(187) = GetValue(AddressOf GetPlayerMagicLevel)
    FunctionTable(188) = GetValue(AddressOf GivePlayerMagicExp)
    
    FunctionTable(189) = GetValue(AddressOf GetBootLocationMap)
    FunctionTable(190) = GetValue(AddressOf GetBootLocationX)
    FunctionTable(191) = GetValue(AddressOf GetBootLocationY)
    
    FunctionTable(192) = GetValue(AddressOf SetPlayerSkillLevel)
    FunctionTable(193) = GetValue(AddressOf SetPlayerMagicLevel)
    
    FunctionTable(194) = GetValue(AddressOf DespawnMonster)
    
    FunctionTable(195) = GetValue(AddressOf GiveEliteExp)
End Sub
