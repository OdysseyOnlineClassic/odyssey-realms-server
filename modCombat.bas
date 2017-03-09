Attribute VB_Name = "modCombat"
Option Explicit
Function ProjectileDamage(Index As Long) As Long
    Dim A As Long, Damage As Long, Weapon As Long, Modifier As Long
    With Player(Index)
        If .EquippedObject(1).Object > 0 Then
            'Uses Weapon
            Weapon = .EquippedObject(1).Object
            If .EquippedObject(6).Object > 0 Then
                Damage = Int(World.StatStrength) + Object(.EquippedObject(1).Object).Data(0) + Object(.Inv(.EquippedObject(6).Object).Object).Data(0) + 1
            End If
        End If

        If .EquippedObject(5).Object > 0 Then
            'Has Ring
            If Object(.EquippedObject(5).Object).Data(0) = 0 Then
                Modifier = Object(.EquippedObject(5).Object).Data(2)
                If Not ExamineBit(Object(.EquippedObject(5).Object).flags, 1) = 255 Then A = .EquippedObject(5).Value - 1 Else A = 1
                Damage = Damage + Modifier
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(5)
                    .EquippedObject(5).Object = 0
                    .EquippedObject(5).ItemPrefix = 0
                    .EquippedObject(5).ItemSuffix = 0
                Else
                    .EquippedObject(5).Value = A
                End If
            End If
        End If
    End With
    ProjectileDamage = Damage
End Function

Sub ProjectileAttackPlayer(Index As Long, A As Long)
    Dim B As Long, C As Long, F As Long, St As String, X As Long
    With Player(Index)
        If .IsDead = False Then
            If ExamineBit(Map(.Map).flags, 0) = False And Not Map(.Map).Tile(Player(A).X, Player(A).Y).Att = 6 And Not Map(.Map).Tile(.X, .Y).Att = 6 Then
                If A >= 1 And A <= MaxUsers And Player(A).IsDead = False Then
                    If Player(A).Mode = modePlaying And Player(A).Map = .Map Then
                        If .Guild > 0 Or ExamineBit(Map(.Map).flags, 6) = True Then
                            If Player(A).Guild > 0 Or ExamineBit(Map(.Map).flags, 6) = True Then
                                Parameter(0) = Index
                                Parameter(1) = A
                                If RunScript("ATTACKPLAYER") = 0 Then
                                    With Player(A)
                                        B = 0
                                        C = PlayerArmor(A, ProjectileDamage(Index))
                                        If C < 0 Then C = 0
                                        If C > 255 Then C = 255
                                        If .HP > C Then
                                            .HP = .HP - C
                                        Else
                                            .HP = 0
                                        End If
                                    End With
                                    SendSocket A, Chr$(49) + Chr$(B) + Chr$(Index) + Chr$(C)
                                    If B = 1 Then
                                        St = DoubleChar$(4) + Chr$(117) + Chr$(Player(A).X) + Chr$(Player(A).Y) + Chr$(1)
                                    Else
                                        St = DoubleChar$(5) + Chr$(111) + Chr$(12) + Chr$(C) + Chr$(Player(A).X) + Chr$(Player(A).Y)
                                    End If
                                    SendToMapRaw CLng(.Map), St
                                    If Player(A).HP = 0 Then
                                        Parameter(0) = Index
                                        Parameter(1) = A
                                        If RunScript("KILLPLAYER") = 0 Then
    
                                            If ExamineBit(Map(Player(Index).Map).flags, 7) = False Then
                                                If Player(Index).Guild > 0 And Player(A).Guild > 0 Then
                                                    Guild(.Guild).Kills = Guild(.Guild).Kills + 1
                                                    Guild(.Guild).Member(.GuildSlot).Kills = Guild(.Guild).Member(.GuildSlot).Kills + 1
                                                    For X = 0 To DeclarationCount
                                                        If Guild(.Guild).Declaration(X).Guild = Player(A).Guild And Guild(.Guild).Declaration(X).Type = 1 Then
                                                            Guild(.Guild).Declaration(X).Kills = Guild(.Guild).Declaration(X).Kills + 1
                                                        End If
                                                    Next X
                                                    Guild(.Guild).UpdateFlag = True
    
                                                    Guild(Player(A).Guild).Deaths = Guild(Player(A).Guild).Deaths + 1
                                                    Guild(Player(A).Guild).Member(Player(A).GuildSlot).Deaths = Guild(Player(A).Guild).Member(Player(A).GuildSlot).Deaths + 1
                                                    For X = 0 To DeclarationCount
                                                        If Guild(Player(A).Guild).Declaration(X).Guild = .Guild And Guild(Player(A).Guild).Declaration(X).Type = 1 Then
                                                            Guild(Player(A).Guild).Declaration(X).Deaths = Guild(Player(A).Guild).Declaration(X).Deaths + 1
                                                        End If
                                                    Next X
                                                    Guild(Player(A).Guild).UpdateFlag = True
                                                End If
                                            End If
    
                                            'Player Died
                                            SendSocket A, Chr$(52) + Chr$(Index)    'Player Killed You
                                            SendAllButBut Index, A, Chr$(61) + Chr$(A) + Chr$(Index)    'Player was killed by player
                                            SendSocket Index, Chr$(45) + Chr$(A) + QuadChar(.Experience)    'You Killed Player
    
                                            F = Player(A).Experience
                                            B = Player(A).Status
                                            If PlayerDied(A, Index) = True Then
                                                If Not A = Index Then
                                                    If B <> 1 Then
                                                        .Status = 1
                                                    End If
                                                    F = F - Player(A).Experience
                                                    If Player(A).Level > 80 And Player(Index).Level > 80 Then
                                                        GainEliteExp Index, F
                                                    Else
                                                        GainExp Index, F
                                                    End If
                                                End If
                                            End If
                                            SetPlayerStatus Index, .Status
                                        End If
                                    End If
                                Else
                                    SendHPUpdate A
                                End If
                            Else
                                SendSocket Index, Chr$(16) + Chr$(19)    'Player not in guild
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(20)    'You are not in guild
                        End If
                    End If
                End If
            Else
                SendSocket Index, Chr$(16) + Chr$(9)    'Friendly Zone
            End If
        End If
    End With
End Sub

Sub ProjectileAttackMonster(Index As Long, A As Long)
    Dim B As Long, C As Long, D As Long, E As Long
    With Player(Index)
        If .IsDead = False Then
            If ExamineBit(Map(.Map).flags, 5) = False Then
                If A <= MaxMonsters Then
                    If Map(.Map).Monster(A).Monster > 0 Then
                        Parameter(0) = Index
                        If RunScript("ATTACKMONSTER" + CStr(Map(.Map).Monster(A).Monster)) = 0 Then
                            With Monster(Map(.Map).Monster(A).Monster)
                                'Hit Target
                                B = 0
                                C = ProjectileDamage(Index) - .Armor
                                If C < 0 Then C = 0
                                If C > 255 Then C = 255
                            End With

                            With Map(.Map).Monster(A)
                                .Target = Index
                                .TargetIsMonster = False
                                If .HP > C Then
                                    .HP = .HP - C

                                    SendToMap Player(Index).Map, Chr$(44) + Chr$(Index) + Chr$(B) + Chr$(A) + Chr$(C) + DoubleChar$(CLng(.HP))
                                Else
                                    SendToMap Player(Index).Map, Chr$(44) + Chr$(Index) + Chr$(B) + Chr$(A) + Chr$(C) + DoubleChar$(CLng(.HP))

                                    'Monster Died
                                    SendToMapAllBut Player(Index).Map, Index, Chr$(39) + Chr$(A)    'Monster Died

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
                                        NewMapObject CLng(Player(Index).Map), E, Monster(.Monster).Value(D), CLng(.X), CLng(.Y), False
                                    End If

                                    Parameter(0) = Index
                                    RunScript "MONSTERDIE" + CStr(.Monster)

                                    .Monster = 0
                                End If
                            End With
                        End If
                    Else
                        SendSocket Index, Chr$(16) + Chr$(5)    'No such monster
                    End If
                End If
            Else
                SendSocket Index, Chr$(16) + Chr$(12)    'Can't attack monsters here
            End If
        End If
    End With
End Sub

Sub MagicAttackPlayer(Index As Long, A As Long, MagicDamage As Long)
    Dim B As Long, C As Long, F As Long, St As String, X As Long
    With Player(Index)
        If .IsDead = False Then
            If ExamineBit(Map(.Map).flags, 0) = False And Not Map(.Map).Tile(Player(A).X, Player(A).Y).Att = 6 And Not Map(.Map).Tile(.X, .Y).Att = 6 Then
                If A >= 1 And A <= MaxUsers And Player(A).IsDead = False Then
                    If Player(A).Mode = modePlaying And Player(A).Map = .Map Then
                        If .Guild > 0 Or ExamineBit(Map(.Map).flags, 6) = True Then
                            If Player(A).Guild > 0 Or ExamineBit(Map(.Map).flags, 6) = True Then
                                Parameter(0) = Index
                                Parameter(1) = A
                                If RunScript("ATTACKPLAYER") = 0 Then
                                    With Player(A)
                                        B = 0
                                        MagicDamage = MagicDamage + (World.StatConcentration)
                                        C = MagicArmor(A, MagicDamage)
                                        If C < 0 Then C = 0
                                        If C > 255 Then C = 255
                                        If .HP > C Then
                                            .HP = .HP - C
                                        Else
                                            .HP = 0
                                        End If
                                    End With
                                    SendSocket A, Chr$(49) + Chr$(B) + Chr$(Index) + Chr$(C)
                                    If B = 1 Then
                                        St = DoubleChar$(4) + Chr$(117) + Chr$(Player(A).X) + Chr$(Player(A).Y) + Chr$(1)
                                    Else
                                        St = DoubleChar$(5) + Chr$(111) + Chr$(12) + Chr$(C) + Chr$(Player(A).X) + Chr$(Player(A).Y)
                                    End If
                                    SendToMapRaw CLng(.Map), St
                                    If Player(A).HP = 0 Then
                                        Parameter(0) = Index
                                        Parameter(1) = A
                                        If RunScript("KILLPLAYER") = 0 Then
    
                                            If ExamineBit(Map(Player(Index).Map).flags, 7) = False Then
                                                If Player(Index).Guild > 0 And Player(A).Guild > 0 Then
                                                    Guild(.Guild).Kills = Guild(.Guild).Kills + 1
                                                    Guild(.Guild).Member(.GuildSlot).Kills = Guild(.Guild).Member(.GuildSlot).Kills + 1
                                                    For X = 0 To DeclarationCount
                                                        If Guild(.Guild).Declaration(X).Guild = Player(A).Guild And Guild(.Guild).Declaration(X).Type = 1 Then
                                                            Guild(.Guild).Declaration(X).Kills = Guild(.Guild).Declaration(X).Kills + 1
                                                        End If
                                                    Next X
                                                    Guild(.Guild).UpdateFlag = True
    
                                                    Guild(Player(A).Guild).Deaths = Guild(Player(A).Guild).Deaths + 1
                                                    Guild(Player(A).Guild).Member(Player(A).GuildSlot).Deaths = Guild(Player(A).Guild).Member(Player(A).GuildSlot).Deaths + 1
                                                    For X = 0 To DeclarationCount
                                                        If Guild(Player(A).Guild).Declaration(X).Guild = .Guild And Guild(Player(A).Guild).Declaration(X).Type = 1 Then
                                                            Guild(Player(A).Guild).Declaration(X).Deaths = Guild(Player(A).Guild).Declaration(X).Deaths + 1
                                                        End If
                                                    Next X
                                                    Guild(Player(A).Guild).UpdateFlag = True
                                                End If
                                            End If
    
                                            'Player Died
                                            SendSocket A, Chr$(52) + Chr$(Index)    'Player Killed You
                                            SendSocket Index, Chr$(45) + Chr$(A) + QuadChar(.Experience)    'You Killed Player
                                            SendAllButBut Index, A, Chr$(61) + Chr$(A) + Chr$(Index)    'Player was killed by player
    
                                            F = Player(A).Experience
                                            B = Player(A).Status
                                            If PlayerDied(A, Index) = True Then
                                                If B <> 1 Then
                                                    .Status = 1
                                                End If
                                                F = F - Player(A).Experience
                                                If Player(A).Level > 80 And Player(Index).Level > 80 Then
                                                    GainEliteExp Index, F
                                                Else
                                                    GainExp Index, F
                                                End If
                                            End If
                                           
                                            SetPlayerStatus Index, .Status
                                        End If
                                    Else
                                        SendHPUpdate A
                                    End If
                                End If
                            Else
                                SendSocket Index, Chr$(16) + Chr$(19)    'Player not in guild
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(20)    'You are not in guild
                        End If
                    End If
                End If
            Else
                SendSocket Index, Chr$(16) + Chr$(9)    'Friendly Zone
            End If
        End If
    End With
End Sub

Sub MagicAttackMonster(Index As Long, A As Long, MagicDamage As Long)
    Dim B As Long, C As Long, D As Long, E As Long
    With Player(Index)
        If .IsDead = False Then
            If ExamineBit(Map(.Map).flags, 5) = False Then
                If A <= MaxMonsters Then
                    If Map(.Map).Monster(A).Monster > 0 Then
                        Parameter(0) = Index
                        If RunScript("ATTACKMONSTER" + CStr(Map(.Map).Monster(A).Monster)) = 0 Then
                            With Monster(Map(.Map).Monster(A).Monster)
                                'Hit Target
                                B = 0
                                MagicDamage = MagicDamage + (World.StatConcentration)
                                C = MagicDamage - .MagicDefense
                                If C < 0 Then C = 0
                                If C > 255 Then C = 255
                            End With

                            With Map(.Map).Monster(A)
                                .Target = Index
                                .TargetIsMonster = False
                                If .HP > C Then
                                    .HP = .HP - C
                                    SendToMap Player(Index).Map, Chr$(44) + Chr$(Index) + Chr$(B) + Chr$(A) + Chr$(C) + DoubleChar$(CLng(.HP))
                                Else
                                    SendToMap Player(Index).Map, Chr$(44) + Chr$(Index) + Chr$(B) + Chr$(A) + Chr$(C) + DoubleChar$(CLng(.HP))

                                    'Monster Died
                                    SendToMapAllBut Player(Index).Map, Index, Chr$(39) + Chr$(A)    'Monster Died

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
                                        NewMapObject CLng(Player(Index).Map), E, Monster(.Monster).Value(D), CLng(.X), CLng(.Y), False
                                    End If

                                    Parameter(0) = Index
                                    RunScript "MONSTERDIE" + CStr(.Monster)

                                    .Monster = 0
                                End If
                            End With
                        End If
                    Else
                        SendSocket Index, Chr$(16) + Chr$(5)    'No such monster
                    End If
                End If
            Else
                SendSocket Index, Chr$(16) + Chr$(12)    'Can't attack monsters here
            End If
        End If
    End With
End Sub


Function PlayerArmor(Index As Long, ByVal Damage As Long) As Long
    Dim A As Long
    With Player(Index)
        If .EquippedObject(2).Object > 0 Then
            Randomize
            If Int(Rnd * 100) < statPlayerAgility Then
                'Uses shield
                If .EquippedObject(2).Object > 0 Then
                    If Not ExamineBit(Object(.EquippedObject(2).Object).flags, 1) = 255 Then A = .EquippedObject(2).Value - 1 Else A = 1
                    Damage = Damage - Object(.EquippedObject(2).Object).Data(1)
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr$(57) + Chr$(2)
                        .EquippedObject(2).Value = 0
                        .EquippedObject(2).Object = 0
                        .EquippedObject(2).ItemPrefix = 0
                        .EquippedObject(2).ItemSuffix = 0
                        CalculateStats Index
                    Else
                        .EquippedObject(2).Value = A
                    End If
                End If
            End If
        End If
        If .EquippedObject(5).Object > 0 Then
            If Object(.EquippedObject(5).Object).Data(0) = 1 Then    'Defensive
                'Uses ring
                If Not ExamineBit(Object(.EquippedObject(5).Object).flags, 1) = 255 Then A = .EquippedObject(5).Value - 1 Else A = 1
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(5)
                    .Inv(5).Object = 0
                    .Inv(5).ItemPrefix = 0
                    .Inv(5).ItemSuffix = 0
                    .EquippedObject(5).Object = 0
                    .EquippedObject(5).ItemPrefix = 0
                    .EquippedObject(5).ItemSuffix = 0
                    CalculateStats Index
                Else
                    .EquippedObject(5).Value = A
                End If
            End If
        End If
        If .EquippedObject(3).Object > 0 Then
            'Uses armor
            If .EquippedObject(3).Object > 0 Then
                If Not ExamineBit(Object(.EquippedObject(3).Object).flags, 1) = 255 Then A = .EquippedObject(3).Value - 1 Else A = 1
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(3)
                    .EquippedObject(3).Value = 0
                    .EquippedObject(3).Object = 0
                    .EquippedObject(3).ItemPrefix = 0
                    .EquippedObject(3).ItemSuffix = 0
                    CalculateStats Index
                Else
                    .EquippedObject(3).Value = A
                End If
            End If
        End If
        'Helm
        If .EquippedObject(4).Object > 0 Then
            'Uses helm
            If .EquippedObject(4).Object > 0 Then
                If Not ExamineBit(Object(.EquippedObject(4).Object).flags, 1) = 255 Then A = .EquippedObject(4).Value - 1 Else A = 1
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(4)
                    .EquippedObject(4).Value = 0
                    .EquippedObject(4).Object = 0
                    .EquippedObject(4).ItemPrefix = 0
                    .EquippedObject(4).ItemSuffix = 0
                    CalculateStats Index
                Else
                    .EquippedObject(4).Value = A
                End If
            End If
        End If
        Damage = Damage - .TotalDefense
    End With
    PlayerArmor = Damage
End Function

Function MagicArmor(Index As Long, ByVal Damage As Long) As Long
    Dim A As Long, ObjNum As Long
    With Player(Index)
        If .EquippedObject(2).Object > 0 Then
            Randomize
            If Int(Rnd * 100) < statPlayerAgility Then
                'Uses shield
                If .EquippedObject(2).Object > 0 Then
                    If Not ExamineBit(Object(.EquippedObject(2).Object).flags, 1) = 255 Then A = .EquippedObject(2).Value - 1 Else A = 1
                    Damage = Damage - Object(.EquippedObject(2).Object).Data(2)
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr$(57) + Chr$(2)
                        .EquippedObject(2).Value = 0
                        .EquippedObject(2).Object = 0
                        .EquippedObject(2).ItemPrefix = 0
                        .EquippedObject(2).ItemSuffix = 0
                        CalculateStats Index
                    Else
                        .EquippedObject(2).Value = A
                    End If
                End If
            End If
        End If
        If .EquippedObject(5).Object > 0 Then
            If Object(.EquippedObject(5).Object).Data(0) = 2 Then    'Magic Defensive
                'Uses ring
                If Not ExamineBit(Object(.EquippedObject(5).Object).flags, 1) = 255 Then A = .EquippedObject(5).Value - 1 Else A = 1
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(ObjNum)
                    .Inv(ObjNum).Object = 0
                    .EquippedObject(5).Object = 0
                    .EquippedObject(5).ItemPrefix = 0
                    .EquippedObject(5).ItemSuffix = 0
                    CalculateStats Index
                Else
                    .EquippedObject(5).Value = A
                End If
            End If
        End If
        If .EquippedObject(3).Object > 0 Then
            'Uses armor
            If .EquippedObject(3).Object > 0 Then
                If Not ExamineBit(Object(.EquippedObject(3).Object).flags, 1) = 255 Then A = .EquippedObject(3).Value - 1 Else A = 1
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(3)
                    .EquippedObject(3).Value = 0
                    .EquippedObject(3).Object = 0
                    .EquippedObject(3).ItemPrefix = 0
                    .EquippedObject(3).ItemSuffix = 0
                    CalculateStats Index
                Else
                    .EquippedObject(3).Value = A
                End If
            End If
        End If
        'Helm
        If .EquippedObject(4).Object > 0 Then
            'Uses helm
            If .EquippedObject(4).Object > 0 Then
                If Not ExamineBit(Object(.EquippedObject(4).Object).flags, 1) = 255 Then A = .EquippedObject(4).Value - 1 Else A = 1
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(4)
                    .EquippedObject(4).Value = 0
                    .EquippedObject(4).Object = 0
                    .EquippedObject(4).ItemPrefix = 0
                    .EquippedObject(4).ItemSuffix = 0
                    CalculateStats Index
                Else
                    .EquippedObject(4).Value = A
                End If
            End If
        End If
        Damage = Damage - .MagicDefense
    End With
    MagicArmor = Damage
End Function
Function PlayerDamage(Index As Long) As Long
    Dim A As Long, Modifier As Long
    With Player(Index)
        If .EquippedObject(1).Object > 0 Then
            If Object(.EquippedObject(1).Object).Type = 1 Then
                'Uses Weapon
                If Not ExamineBit(Object(.EquippedObject(1).Object).flags, 1) = 255 Then A = .EquippedObject(1).Value - 1 Else A = 1
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(1)
                    .EquippedObject(1).Object = 0
                    .EquippedObject(1).ItemPrefix = 0
                    .EquippedObject(1).ItemSuffix = 0
                    .EquippedObject(1).Value = 0
                    CalculateStats Index
                Else
                    .EquippedObject(1).Value = A
                End If
            End If
        Else
        End If

        If .EquippedObject(5).Object > 0 Then
            'Has Ring
            If Object(.EquippedObject(5).Object).Data(0) = 0 Then
                Modifier = Object(.EquippedObject(5).Object).Data(2)
                If Not ExamineBit(Object(.EquippedObject(5).Object).flags, 1) = 255 Then A = .EquippedObject(5).Value - 1 Else A = 1
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(5)
                    .EquippedObject(5).Object = 0
                    .EquippedObject(5).ItemPrefix = 0
                    .EquippedObject(5).ItemSuffix = 0
                    .EquippedObject(5).Value = 0
                    CalculateStats Index
                Else
                    .EquippedObject(5).Value = A
                End If
            End If
        End If

        PlayerDamage = .PhysicalAttack
    End With
End Function

Sub CombatAttackPlayer(Index As Long, A As Long, Damage As Long)
    Dim B As Long, C As Long, F As Long, St As String, X As Long
    With Player(Index)
        If .IsDead = False Then
            If ExamineBit(Map(.Map).flags, 0) = False And Not Map(.Map).Tile(Player(A).X, Player(A).Y).Att = 6 And Not Map(.Map).Tile(.X, .Y).Att = 6 Then
                If A >= 1 And A <= MaxUsers And Player(A).IsDead = False Then
                    If Player(A).Mode = modePlaying And Player(A).Map = .Map Then
                        If .Guild > 0 Or ExamineBit(Map(.Map).flags, 6) = True Then
                            If Sqr((CSng(Player(A).X) - CSng(.X)) ^ 2 + (CSng(Player(A).Y) - CSng(.Y)) ^ 2) <= LagHitDistance Then
                                If Player(A).Guild > 0 Or ExamineBit(Map(.Map).flags, 6) = True Then
                                    Parameter(0) = Index
                                    Parameter(1) = A
                                    If RunScript("AttackPlayer") = 0 Then
                                        With Player(A)
                                            C = PlayerArmor(A, Damage)
                                            If C < 0 Then C = 0
                                            If C > 255 Then C = 255
    
                                            If .HP > C Then
                                                .HP = .HP - C
                                            Else
                                                .HP = 0
                                            End If
                                        End With
                                        SendSocket A, Chr$(49) + Chr$(Index) + Chr$(C)
                                        SendSocket Index, Chr$(43) + Chr$(B) + Chr$(A) + Chr$(C)
                                        If B = 1 Then
                                            St = DoubleChar$(4) + Chr$(117) + Chr$(Player(A).X) + Chr$(Player(A).Y) + Chr$(1)
                                        Else
                                            St = DoubleChar$(5) + Chr$(111) + Chr$(12) + Chr$(C) + Chr$(Player(A).X) + Chr$(Player(A).Y)
                                        End If
                                        St = St + vbNullChar + Chr$(2) + Chr$(45) + Chr$(Index)
                                        SendToMapRaw CLng(.Map), St
                                        If Player(A).HP = 0 Then
                                            Parameter(0) = Index
                                            Parameter(1) = A
                                            If RunScript("KILLPLAYER") = 0 Then
    
                                                If ExamineBit(Map(Player(Index).Map).flags, 7) = False Then
                                                    If Player(Index).Guild > 0 And Player(A).Guild > 0 Then
                                                        Guild(.Guild).Kills = Guild(.Guild).Kills + 1
                                                        Guild(.Guild).Member(.GuildSlot).Kills = Guild(.Guild).Member(.GuildSlot).Kills + 1
                                                        For X = 0 To DeclarationCount
                                                            If Guild(.Guild).Declaration(X).Guild = Player(A).Guild And Guild(.Guild).Declaration(X).Type = 1 Then
                                                                Guild(.Guild).Declaration(X).Kills = Guild(.Guild).Declaration(X).Kills + 1
                                                            End If
                                                        Next X
                                                        Guild(.Guild).UpdateFlag = True
    
                                                        Guild(Player(A).Guild).Deaths = Guild(Player(A).Guild).Deaths + 1
                                                        Guild(Player(A).Guild).Member(Player(A).GuildSlot).Deaths = Guild(Player(A).Guild).Member(Player(A).GuildSlot).Deaths + 1
                                                        For X = 0 To DeclarationCount
                                                            If Guild(Player(A).Guild).Declaration(X).Guild = .Guild And Guild(Player(A).Guild).Declaration(X).Type = 1 Then
                                                                Guild(Player(A).Guild).Declaration(X).Deaths = Guild(Player(A).Guild).Declaration(X).Deaths + 1
                                                            End If
                                                        Next X
                                                        Guild(Player(A).Guild).UpdateFlag = True
                                                    End If
                                                End If
    
                                                'Player Died
                                                SendSocket A, Chr$(52) + Chr$(Index)    'Player Killed You
                                                SendSocket Index, Chr$(45) + Chr$(A) + QuadChar(.Experience)    'You Killed Player
                                                SendAllButBut Index, A, Chr$(61) + Chr$(A) + Chr$(Index)    'Player was killed by player
    
                                                F = Player(A).Experience
                                                B = Player(A).Status
                                                If PlayerDied(A, Index) = True Then
                                                    If B <> 1 Then
                                                        .Status = 1
                                                    End If
                                                    F = F - Player(A).Experience
                                                    If Player(A).Level > 80 And Player(Index).Level > 80 Then
                                                        GainEliteExp Index, F
                                                    Else
                                                        GainExp Index, F
                                                    End If
                                                End If
                                                SetPlayerStatus Index, .Status
                                            End If
                                        Else
                                            SendHPUpdate A
                                        End If
                                    End If
                                Else
                                    SendSocket Index, Chr$(16) + Chr$(19)    'Player not in guild
                                End If
                            Else
                                SendSocket Index, Chr$(16) + Chr$(6)    'Too far away
                            End If
                        Else
                            SendSocket Index, Chr$(16) + Chr$(20)    'You are not in guild
                        End If
                    End If
                End If
            Else
                SendSocket Index, Chr$(16) + Chr$(9)    'Friendly Zone
            End If
        End If
    End With
End Sub

Sub SendHPUpdate(Index As Long)
    Dim A As Long
    With Player(Index)
        If .Guild > 0 Then
            For A = 1 To MaxUsers
                If Not A = Index Then
                    If Player(A).InUse = True Then
                        If Player(A).Map = .Map Then
                            If Player(A).Guild > 0 Then
                                If (Player(A).Guild = .Guild Or IsGuildAlly(.Guild, Player(A).Guild) = True) Then
                                    SendSocket A, Chr$(150) + Chr$(Index) + Chr$(.HP)
                                End If
                            End If

                            If Player(A).Access > 0 Then
                                SendSocket A, Chr$(150) + Chr$(Index) + Chr$(.HP)
                            End If
                        End If
                    End If
                End If
            Next A
        End If
    End With
End Sub

Function IsGuildAlly(GuildIndex As Byte, Ally As Byte) As Boolean
    Dim A As Long
    For A = 0 To DeclarationCount
        If Guild(GuildIndex).Declaration(A).Type = 0 Then
            If Guild(GuildIndex).Declaration(A).Guild = Ally Then
                IsGuildAlly = True
                Exit Function
            End If
        End If
    Next A
End Function

Function FindProjectileDamageSlot(Index As Long) As Long
    Dim A As Long
    For A = 1 To 20
        If Player(Index).ProjectileDamage(A).Live = False Then
            FindProjectileDamageSlot = A
            Exit Function
        Else
            If Player(Index).ProjectileDamage(A).ShootTime + 10000 < getTime Then
                FindProjectileDamageSlot = A
                Exit Function
            End If
        End If
    Next A
    
    FindProjectileDamageSlot = 1
End Function
