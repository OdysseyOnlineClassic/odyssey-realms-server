Attribute VB_Name = "modMonsters"
Option Explicit

Public Function MoveMonsters(MapNum As Long, Tick As Currency) As String
    Dim St1 As String
    Dim A As Long
    
    With Map(MapNum)
        For A = 0 To MaxMonsters
            With .Monster(A)
                If .Monster > 0 Then
                    If Tick > .MoveTimer Then
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then
                            .MoveTimer = Tick + 440
                        Else
                            .MoveTimer = Tick + 240
                        End If
                        If .Target = 0 And .TargetIsMonster = False Then
                            St1 = St1 + MonsterHasNoTarget(MapNum, Tick, A)
                        Else
                            St1 = St1 + MonsterHasTarget(MapNum, Tick, A)
                        End If
                    End If
                Else
                    If A <= MaxMonsters Then
                        With Map(MapNum).MonsterSpawn(Int(A / 2))
                            If .Monster > 0 Then
                                If Tick > .Timer Then
                                    If Int(Rnd * .Rate) = 0 Then
                                        St1 = St1 + NewMapMonster(MapNum, A)
                                    End If
                                    .Timer = Tick + 500
                                End If
                            End If
                        End With
                    End If
                End If
            End With
        Next A
    End With
    
    MoveMonsters = St1
End Function

Private Function MonsterHasNoTarget(MapNum As Long, Tick As Currency, MonsterIndex As Long) As String
    Dim St1 As String
    Dim B As Long, C As Long, D As Long, E As Long
    With Map(MapNum).Monster(MonsterIndex)
        'Supersight looks for closest target
        If ExamineBit(Monster(.Monster).flags, 1) = True Then
            B = .Target
            D = 1000
            For C = 1 To MaxUsers
                If Player(C).InUse = True And Not B = C Then
                    If Player(C).Map = MapNum And Player(C).IsDead = False Then
                        If ExamineBit(Monster(.Monster).flags, 3) = False Then
                            'Isn't Friendly
                            If ExamineBit(Monster(.Monster).flags, 0) = False Or Player(C).Status = 1 Then
                                If Player(C).Access = 0 Then
                                    E = Sqr((CSng(.X) - CSng(Player(C).X)) ^ 2 + (CSng(.Y) - CSng(Player(C).Y)) ^ 2)
                                    If E <= Monster(.Monster).Sight Then
                                        If E < D Then
                                            Parameter(0) = C
                                            If RunScript("MONSTERSEE" + CStr(.Monster)) = 0 Then
                                                .Target = C
                                                B = .Target
                                                D = Sqr((CSng(.X) - CSng(Player(C).X)) ^ 2 + (CSng(.Y) - CSng(Player(C).Y)) ^ 2)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next C
        End If
        
        'Random Movement
        If Rnd < 0.3 Then
            .D = Int(Rnd * 4)
            Select Case .D
                Case 0    'Up
                    If .Y > 0 Then
                        If IsVacant(MapNum, CLng(.X), CLng(.Y - 1)) Then
                            If NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 0) Then
                                .Y = .Y - 1
                                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(MonsterIndex) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                                CheckIfMonstersNotice MapNum, MonsterIndex
                            End If
                        End If
                    End If
                Case 1    'Down
                    If .Y < 11 Then
                        If IsVacant(MapNum, CLng(.X), CLng(.Y + 1)) Then
                            If NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 1) Then
                                .Y = .Y + 1
                                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(MonsterIndex) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                                CheckIfMonstersNotice MapNum, MonsterIndex
                            End If
                        End If
                    End If
                Case 2    'Left
                    If .X > 0 Then
                        If IsVacant(MapNum, CLng(.X - 1), CLng(.Y)) Then
                            If NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 2) Then
                                .X = .X - 1
                                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(MonsterIndex) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                                CheckIfMonstersNotice MapNum, MonsterIndex
                            End If
                        End If
                    End If
                Case 3    'Right
                    If .X < 11 Then
                        If IsVacant(MapNum, CLng(.X + 1), CLng(.Y)) Then
                            If NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 3) Then
                                .X = .X + 1
                                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(MonsterIndex) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                                CheckIfMonstersNotice MapNum, MonsterIndex
                            End If
                        End If
                    End If
            End Select
        End If
    End With
    
    MonsterHasNoTarget = St1
End Function

Private Function MonsterHasTarget(MapNum As Long, Tick As Currency, MonsterIndex As Long) As String
    Dim St1 As String
    Dim C As Long, D As Long, E As Long, F As Long, G As Long, H As Long, I As Long, J As Long
    
    With Map(MapNum).Monster(MonsterIndex)
        If .TargetIsMonster = False Then
            D = Sqr((CSng(.X) - CSng(Player(.Target).X)) ^ 2 + (CSng(.Y) - CSng(Player(.Target).Y)) ^ 2)
            
            'Supersight looks for closest target
            If ExamineBit(Monster(.Monster).flags, 1) = True Then
                For C = 1 To MaxUsers
                    If Player(C).InUse = True And Not .Target = C Then
                        If Player(C).Map = MapNum And Player(C).IsDead = False Then
                            If ExamineBit(Monster(.Monster).flags, 3) = False Then
                                'Isn't Friendly
                                If ExamineBit(Monster(.Monster).flags, 0) = False Or Player(C).Status = 1 Then
                                    If Player(C).Access = 0 Then
                                        E = Sqr((CSng(.X) - CSng(Player(C).X)) ^ 2 + (CSng(.Y) - CSng(Player(C).Y)) ^ 2)
                                        If E <= Monster(.Monster).Sight Then
                                            If E < D Then
                                                Parameter(0) = C
                                                If RunScript("MONSTERSEE" + CStr(.Monster)) = 0 Then
                                                    .Target = C
                                                    D = Sqr((CSng(.X) - CSng(Player(C).X)) ^ 2 + (CSng(.Y) - CSng(Player(C).Y)) ^ 2)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next C
            End If
        End If
        Dim ValidTarget As Boolean
        If .TargetIsMonster = False Then
            If Player(.Target).Mode = modePlaying And Player(.Target).Map = MapNum And Player(.Target).IsDead = False Then
                ValidTarget = True
            End If
        Else
            If Map(MapNum).Monster(.Target).Monster > 0 Then
                ValidTarget = True
            End If
        End If
        
        'Move Toward Target
        If ValidTarget = True Then
            C = .X
            D = .Y
            E = .D

            Dim TryToMove As Boolean
            TryToMove = False
            
            If .TargetIsMonster = False Then
                F = CLng((Player(.Target).X) - C) ^ 2
                G = CLng((Player(.Target).Y) - D) ^ 2
                I = Player(.Target).X
                J = Player(.Target).Y
            Else
                F = CLng((Map(MapNum).Monster(.Target).X) - C) ^ 2
                G = CLng((Map(MapNum).Monster(.Target).Y) - D) ^ 2
                I = Map(MapNum).Monster(.Target).X
                J = Map(MapNum).Monster(.Target).Y
            End If

            If Sqr(CSng(F + G)) > 1 Then
                If .X < C And .D <> 3 Then
                    .D = 3
                ElseIf .X > C And .D <> 2 Then
                    .D = 2
                ElseIf .Y < D And .D <> 1 Then
                    .D = 1
                ElseIf .Y > D And .D <> 0 Then
                    .D = 0
                End If

                If F > G Then
                    H = 0
                ElseIf G > F Then
                    H = 1
                Else
                    If Rnd < 0.5 Then
                        H = 1
                    Else
                        H = 0
                    End If
                End If

                If H = 0 Then
                    If D < J Then
                        If IsVacant(MapNum, C, D + 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 1) Then
                            H = 1
                        End If
                    ElseIf D > J Then
                        If IsVacant(MapNum, C, D - 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 0) Then
                            H = 1
                        End If
                    End If
                ElseIf H = 1 Then
                    If C < I Then
                        If IsVacant(MapNum, C + 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 3) Then
                            H = 0
                        End If
                    ElseIf C > I Then
                        If IsVacant(MapNum, C - 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 2) Then
                            H = 0
                        End If
                    End If
                End If

                TryToMove = True
            End If

            If TryToMove = False Then
                If .TargetIsMonster = False Then
                    C = Player(.Target).X
                    D = Player(.Target).Y
                Else
                    C = Map(MapNum).Monster(.Target).X
                    D = Map(MapNum).Monster(.Target).Y
                End If
                If .X < C And .D <> 3 Then
                    .D = 3
                ElseIf .X > C And .D <> 2 Then
                    .D = 2
                ElseIf .Y < D And .D <> 1 Then
                    .D = 1
                ElseIf .Y > D And .D <> 0 Then
                    .D = 0
                End If
                If Not .D = E Then
                    St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(MonsterIndex) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                End If
                If NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), CLng(.D)) = False Then
                    TryToMove = True
                End If
            End If

            C = .X
            D = .Y
            E = .D

            If TryToMove = True Then
                If H = 1 Then
                    If C < I Then
                        If IsVacant(MapNum, C + 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 3) Then
                            C = C + 1
                            E = 3
                        End If
                    ElseIf C > I Then
                        If IsVacant(MapNum, C - 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 2) Then
                            C = C - 1
                            E = 2
                        End If
                    End If
                    If C = .X And D = .Y Then
                        If D < J Then
                            If IsVacant(MapNum, C, D + 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 1) Then
                                D = D + 1
                                E = 1
                            Else
                                If Rnd < 0.5 Then
                                    If C > 0 And IsVacant(MapNum, C - 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 2) Then
                                        C = C - 1
                                        E = 2
                                    ElseIf C < 11 And IsVacant(MapNum, C + 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 3) Then
                                        C = C + 1
                                        E = 3
                                    End If
                                Else
                                    If C < 11 And IsVacant(MapNum, C + 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 3) Then
                                        C = C + 1
                                        E = 3
                                    ElseIf C > 0 And IsVacant(MapNum, C - 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 2) Then
                                        C = C - 1
                                        E = 2
                                    End If
                                End If
                            End If
                        ElseIf D > J Then
                            If IsVacant(MapNum, C, D - 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 0) Then
                                D = D - 1
                                E = 0
                            Else
                                If Rnd < 0.5 Then
                                    If C > 0 And IsVacant(MapNum, C - 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 2) Then
                                        C = C - 1
                                        E = 2
                                    ElseIf C < 11 And IsVacant(MapNum, C + 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 3) Then
                                        C = C + 1
                                        E = 3
                                    End If
                                Else
                                    If C < 11 And IsVacant(MapNum, C + 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 3) Then
                                        C = C + 1
                                        E = 3
                                    ElseIf C > 0 And IsVacant(MapNum, C - 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 2) Then
                                        C = C - 1
                                        E = 2
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If D < J Then
                        If IsVacant(MapNum, C, D + 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 1) Then
                            D = D + 1
                            E = 1
                        End If
                    ElseIf D > J Then
                        If IsVacant(MapNum, C, D - 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 0) Then
                            D = D - 1
                            E = 0
                        End If
                    End If
                    If C = .X And D = .Y Then
                        If C < I Then
                            If IsVacant(MapNum, C + 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 3) Then
                                C = C + 1
                                E = 3
                            Else
                                If Rnd < 0.5 Then
                                    If D > 0 And IsVacant(MapNum, C, D - 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 0) Then
                                        D = D - 1
                                        E = 0
                                    ElseIf D < 11 And IsVacant(MapNum, C, D + 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 1) Then
                                        D = D + 1
                                        E = 1
                                    End If
                                Else
                                    If D < 11 And IsVacant(MapNum, C, D + 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 1) Then
                                        D = D + 1
                                        E = 1
                                    ElseIf D > 0 And IsVacant(MapNum, C, D - 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 0) Then
                                        D = D - 1
                                        E = 0
                                    End If
                                End If
                            End If
                        ElseIf C > I Then
                            If IsVacant(MapNum, C - 1, D) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 2) Then
                                C = C - 1
                                E = 2
                            Else
                                If Rnd < 0.5 Then
                                    If D > 0 And IsVacant(MapNum, C, D - 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 0) Then
                                        D = D - 1
                                        E = 0
                                    ElseIf D < 11 And IsVacant(MapNum, C, D + 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 1) Then
                                        D = D + 1
                                        E = 1
                                    Else
                                    End If
                                Else
                                    If D < 11 And IsVacant(MapNum, C, D + 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 1) Then
                                        D = D + 1
                                        E = 1
                                    ElseIf D > 0 And IsVacant(MapNum, C, D - 1) And NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), 0) Then
                                        D = D - 1
                                        E = 0
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If C <> .X Or D <> .Y Or E <> .D Then
                    .X = C
                    .Y = D
                    .D = E
                    
                    CheckIfMonstersNotice MapNum, MonsterIndex
                    St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(MonsterIndex) + Chr$(C) + Chr$(D) + Chr$(.D)
                End If
            Else
                St1 = St1 + MonsterAttack(MapNum, Tick, MonsterIndex, E)
            End If
        Else
            .Target = 0
            .TargetIsMonster = False
        End If
    End With
    
    MonsterHasTarget = St1
End Function

Private Function MonsterAttack(MapNum As Long, Tick As Currency, MonsterIndex As Long, MonsterDirection As Long) As String
    Dim St1 As String
    Dim C As Long, D As Long, E As Long
    Dim AgilityChance As Integer
    
    With Map(MapNum).Monster(MonsterIndex)
        If Tick > .AttackTimer Then
            'Attack
            
            If .TargetIsMonster = False Then
                C = Player(.Target).X
                D = Player(.Target).Y
            Else
                C = Map(MapNum).Monster(.Target).X
                D = Map(MapNum).Monster(.Target).Y
            End If
            
            If .X < C And .D <> 3 Then
                .D = 3
            ElseIf .X > C And .D <> 2 Then
                .D = 2
            ElseIf .Y < D And .D <> 1 Then
                .D = 1
            ElseIf .Y > D And .D <> 0 Then
                .D = 0
            End If
            If Not .D = MonsterDirection Then
                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(MonsterIndex) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
            End If
            If NoDirectionalWalls(MapNum, CLng(.X), CLng(.Y), CLng(.D)) Then
                If ExamineBit(Monster(.Monster).flags, 5) = True Then
                    .AttackTimer = Tick + 440
                Else
                    .AttackTimer = Tick + 800
                End If
            
                If .TargetIsMonster = False Then
                    If CInt(statPlayerAgility) - CInt(Monster(.Monster).Agility) > 0 Then AgilityChance = CInt(statPlayerAgility) - CInt(Monster(.Monster).Agility) Else AgilityChance = 0
                    If Int(Rnd * 100) > AgilityChance Then
                        C = PlayerArmor(CLng(.Target), Monster(.Monster).Strength)
                        If C < 0 Then C = 0
                        If C > 255 Then C = 255
                        SendSocket .Target, Chr$(50) + vbNullChar + Chr$(MonsterIndex) + Chr$(C)
                        St1 = St1 + DoubleChar$(5) + Chr$(111) + Chr$(12) + Chr$(C) + Chr$(Player(.Target).X) + Chr$(Player(.Target).Y)
                        St1 = St1 + DoubleChar(2) + Chr$(41) + Chr$(MonsterIndex)
                        With Player(.Target)
                            If C >= .HP Then
                                Parameter(0) = Map(MapNum).Monster(MonsterIndex).Monster
                                Parameter(1) = MonsterIndex
                                Parameter(2) = Map(MapNum).Monster(MonsterIndex).Target
                                If RunScript("MONSTERKILL" + CStr(Map(MapNum).Monster(MonsterIndex).Monster)) = 0 Then
                                    'Player Died
                                    SendSocket Map(MapNum).Monster(MonsterIndex).Target, Chr$(53) + DoubleChar$(CLng(Map(MapNum).Monster(MonsterIndex).Monster))    'Monster Killed You
                                    SendAllBut Map(MapNum).Monster(MonsterIndex).Target, Chr$(62) + Chr$(Map(MapNum).Monster(MonsterIndex).Target) + DoubleChar$(CLng(Map(MapNum).Monster(MonsterIndex).Monster))    'Player was killed by monster
                                    PlayerDied CLng(Map(MapNum).Monster(MonsterIndex).Target), -1
                                    Map(MapNum).Monster(MonsterIndex).Target = 0
                                    Map(MapNum).Monster(MonsterIndex).TargetIsMonster = False
                                End If
                            Else
                                .HP = .HP - C
                                SendHPUpdate CLng(Map(MapNum).Monster(MonsterIndex).Target)
                            End If
                        End With
                    Else
                        SendSocket .Target, Chr$(50) + Chr$(1) + Chr$(MonsterIndex) + vbNullChar
                        St1 = St1 + DoubleChar$(4) + Chr$(117) + Chr$(Player(.Target).X) + Chr$(Player(.Target).Y) + Chr$(1)
                    End If
                Else
                    'Attack Monster
                    
                    If .Monster > 0 And Map(MapNum).Monster(.Target).Monster > 0 Then
                        Parameter(0) = MapNum
                        Parameter(1) = MonsterIndex
                        Parameter(2) = .Target
                        If RunScript("MVMATTACK" + CStr(Map(MapNum).Monster(.Target).Monster)) = 0 Then

                            'Hit Target
                            C = CLng(Monster(.Monster).Strength) - CLng(Monster(Map(MapNum).Monster(.Target).Monster).Armor)
                            If C < 0 Then C = 0
                            If C > 255 Then C = 255

                            With Map(MapNum).Monster(.Target)
                                .Target = MonsterIndex
                                .TargetIsMonster = True
                                If .HP > C Then
                                    .HP = .HP - C

                                    SendToMap MapNum, Chr$(155) + Chr$(MonsterIndex) + Chr$(Map(MapNum).Monster(MonsterIndex).Target) + Chr$(C) + DoubleChar$(CLng(.HP))
                                Else
                                    SendToMap MapNum, Chr$(155) + Chr$(MonsterIndex) + Chr$(Map(MapNum).Monster(MonsterIndex).Target) + Chr$(C) + DoubleChar$(CLng(.HP))
                                    
                                    'Monster Died
                                    SendToMapAllBut MapNum, -1, Chr$(39) + Chr$(Map(MapNum).Monster(MonsterIndex).Target)    'Monster Died
                                    
                                    'D = Int(Rnd * 3)
                                    'E = Monster(.Monster).Object(D)
                                    'If E > 0 Then
                                    '    NewMapObject MapNum, E, Monster(.Monster).Value(D), CLng(.X), CLng(.Y), False
                                    'End If
                                    Parameter(0) = MapNum
                                    Parameter(1) = MonsterIndex
                                    Parameter(2) = Map(MapNum).Monster(MonsterIndex).Target
                                    RunScript "MVMDIE" + CStr(.Monster)
                                    .Monster = 0
                                    
                                    Map(MapNum).Monster(MonsterIndex).Target = 0
                                    Map(MapNum).Monster(MonsterIndex).TargetIsMonster = False
                                End If
                            End With
                        End If
                    End If
                    
                    
                End If
            End If
        End If
    End With
    
    MonsterAttack = St1
End Function

Private Sub CheckIfMonstersNotice(MapNum As Long, MonsterIndex As Long)
    Dim H As Long, I As Long
    
    With Map(MapNum).Monster(MonsterIndex)
        'Check if monsters notice (MvM)
        For H = 0 To MaxMonsters
            If Not H = MonsterIndex Then
                If Map(MapNum).Monster(H).Monster > 0 Then
                    If Map(MapNum).Monster(H).Target = 0 Then
                        If ExamineBit(Monster(Map(MapNum).Monster(H).Monster).flags, 6) Then 'MvM flag
                            If Not .Monster = Map(MapNum).Monster(H).Monster Then
                                I = Sqr((CLng(.X) - CLng(Map(MapNum).Monster(H).X)) ^ 2 + (CLng(.Y) - CLng(Map(MapNum).Monster(H).Y)) ^ 2)
                                If I <= Monster(Map(MapNum).Monster(H).Monster).Sight Then
                                    Parameter(0) = MapNum
                                    Parameter(1) = MonsterIndex
                                    Parameter(2) = H
                                    If RunScript("MVMSEE" + CStr(Map(MapNum).Monster(MonsterIndex).Monster)) = 0 Then
                                        Map(MapNum).Monster(H).Target = MonsterIndex
                                        Map(MapNum).Monster(H).TargetIsMonster = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next H
    End With
End Sub
