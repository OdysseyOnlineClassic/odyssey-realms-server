Attribute VB_Name = "modMovement"
Option Explicit

Sub ProcessMovement(Index As Long, St As String, MapNum As Long)
    Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long, G As Long, H As Long, I As Long, J As Long
    With Player(Index)
        I = .X
        J = .Y
        A = Asc(Mid$(St, 1, 1))
        B = Asc(Mid$(St, 2, 1))
        If .IsDead = True Then Exit Sub
        If A > 11 Or B > 11 Then
            .X = I
            .Y = J
            PlayerWarp Index, .Map, .X, .Y
            Hacker Index, "Walk Around"
            Exit Sub
        End If

        If Abs(A - CLng(.X)) + Abs(B - CLng(.Y)) <= 1 Then
            If .X <> A Or .Y <> B Then
                .X = A
                .Y = B
                D = 1

                If Asc(Mid$(St, 4, 1)) = 4 Then
                    'Use Energy
                    If .Energy > 0 Then .Energy = .Energy - 1
                End If
            Else
                D = 0
            End If
            .D = Asc(Mid$(St, 3, 1))

            If D = 1 Then
                Select Case .D
                Case 0:
                    If .X = I And .Y = J - 1 Then

                    Else
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                Case 1:
                    If .X = I And .Y = J + 1 Then

                    Else
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                Case 2:
                    If .X = I - 1 And .Y = J Then

                    Else
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                Case 3:
                    If .X = I + 1 And .Y = J Then

                    Else
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                End Select
            End If

            'Check if monsters notice
            If .Access = 0 Then
                For C = 0 To MaxMonsters
                    If Map(MapNum).Monster(C).Monster > 0 Then
                        If ExamineBit(Monster(Map(MapNum).Monster(C).Monster).flags, 3) = False And ExamineBit(Monster(Map(MapNum).Monster(C).Monster).flags, 1) = False Then
                            'Isn't Friendly
                            If ExamineBit(Monster(Map(MapNum).Monster(C).Monster).flags, 0) = False Or .Status = 1 Then
                                With Map(MapNum).Monster(C)
                                    E = .X
                                    F = .Y
                                    G = Monster(.Monster).Sight
                                End With
                                H = Sqr((CLng(.X) - E) ^ 2 + (CLng(.Y) - F) ^ 2)
                                If H <= G Then
                                    Parameter(0) = Index
                                    If RunScript("MONSTERSEE" + CStr(Map(MapNum).Monster(C).Monster)) = 0 Then
                                        With Map(MapNum).Monster(C)
                                            If Index <> .Target Then
                                                .Target = Index
                                                .TargetIsMonster = False
                                            End If
                                        End With
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next C
            End If
            SendToMapAllBut MapNum, Index, Chr$(10) + Chr$(Index) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + Mid$(St, 4, 1)

            Select Case Map(MapNum).Tile(I, J).Att
            Case 17    'Directional Wall
                If D = 1 Then
                    If .Access = 0 Then
                        If NoDirectionalWalls2(CLng(.Map), CLng(.X), CLng(.Y), CLng(.D)) = False Then
                            .X = I
                            .Y = J
                            PlayerWarp Index, .Map, .X, .Y
                            Exit Sub
                        End If
                    End If
                End If
            End Select

            Select Case Map(MapNum).Tile(.X, .Y).Att
            Case 1, 13, 14, 15, 16    'Wall
                If D = 1 Then
                    If .Access = 0 Then
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                End If
            Case 2    'Warp
                A = Map(MapNum).Tile(.X, .Y).AttData(2)
                B = Map(MapNum).Tile(.X, .Y).AttData(3)
                C = CLng(Map(MapNum).Tile(.X, .Y).AttData(0)) * 256 + CLng(Map(MapNum).Tile(.X, .Y).AttData(1))
                If A <= 11 And B <= 11 And C >= 1 And C <= MaxMaps Then
                    PlayerWarp Index, C, A, B
                    Exit Sub
                Else
                    Hacker Index, "God warp"
                End If
            Case 3    'Key Door
                If D = 1 Then
                    If .Access = 0 Then
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                End If
            Case 4    'Door
                C = FreeMapDoorNum(MapNum)
                If C >= 0 Then
                    With Map(MapNum).Door(C)
                        .Att = 4
                        .X = A
                        .Y = B
                        .T = Player(Index).LastMsg
                    End With
                    Map(MapNum).Tile(A, B).Att = 0
                    SendToMap MapNum, Chr$(36) + Chr$(C) + Chr$(A) + Chr$(B)
                End If
            Case 8    'Touch Plate
                F = Map(MapNum).Tile(A, B).AttData(2)
                If F > 0 Then
                    If .Guild > 0 Then
                        If .GuildRank >= 1 And Guild(.Guild).Hall = F Then
                            G = 1
                        Else
                            G = 0
                        End If
                    Else
                        G = 0
                    End If
                Else
                    G = 1
                End If
                If G = 1 Then
                    D = Map(MapNum).Tile(A, B).AttData(0)
                    E = Map(MapNum).Tile(A, B).AttData(1)
                    If D <= 11 And E <= 11 Then
                        If Map(MapNum).Tile(D, E).Att > 0 Then
                            C = FreeMapDoorNum(MapNum)
                            If C >= 0 Then
                                With Map(MapNum).Door(C)
                                    .Att = Map(MapNum).Tile(D, E).Att
                                    .X = D
                                    .Y = E
                                    .T = Player(Index).LastMsg
                                End With
                                Map(MapNum).Tile(D, E).Att = 0
                                SendToMap MapNum, Chr$(36) + Chr$(C) + Chr$(D) + Chr$(E)
                            End If
                        End If
                    End If
                End If
            Case 9    'Damage Tile
                If D = 1 Then
                    A = Map(MapNum).Tile(.X, .Y).AttData(0)
                    SendSocket Index, Chr$(109) & Chr$(A)
                    If A >= .HP Then
                        'Player Died
                        SendAll Chr$(110) & Chr$(Index)
                        PlayerDied Index, -1
                    Else
                        .HP = .HP - A
                        'Floating text
                        SendToMap .Map, Chr$(112) + Chr$(12) + Chr$(.X) + Chr$(.Y) + CStr(A)
                    End If
                End If
            Case 11    'Script
                If D = 1 Then
                    Parameter(0) = Index
                    RunScript "MAP" + CStr(MapNum) + "_" + CStr(A) + "_" + CStr(B)
                End If
            Case 17    'Directional Wall
                If D = 1 Then
                    If .Access = 0 Then
                        If NoDirectionalWalls(CLng(.Map), I, J, CLng(.D)) = False Then
                            .X = I
                            .Y = J
                            PlayerWarp Index, .Map, .X, .Y
                            Exit Sub
                        End If
                    End If
                End If
            Case 19    ' Light
                If D = 1 And .Access = 0 Then
                    If ExamineBit(Map(MapNum).Tile(.X, .Y).AttData(2), 0) Then
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                End If
            Case 20    ' Light Dampening
                If D = 1 And .Access = 0 Then
                    If ExamineBit(Map(MapNum).Tile(.X, .Y).AttData(3), 0) Then
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                End If
            End Select
            Select Case Map(MapNum).Tile(.X, .Y).Att2
            Case 1, 13, 14, 15, 16    'Wall
                If D = 1 Then
                    If .Access = 0 Then
                        .X = I
                        .Y = J
                        PlayerWarp Index, .Map, .X, .Y
                        Exit Sub
                    End If
                End If
            Case 4    'Door
                C = FreeMapDoorNum(MapNum)
                If C >= 0 Then
                    With Map(MapNum).Door(C)
                        .Att = 4
                        .X = A
                        .Y = B
                        .T = Player(Index).LastMsg
                    End With
                    Map(MapNum).Tile(A, B).Att = 0
                    SendToMap MapNum, Chr$(36) + Chr$(C) + Chr$(A) + Chr$(B)
                End If
            Case 11    'Script
                If D = 1 Then
                    Parameter(0) = Index
                    RunScript "MAP" + CStr(MapNum) + "_" + CStr(A) + "_" + CStr(B)
                End If
            End Select
        Else
            PlayerWarp Index, .Map, I, J
            Exit Sub
        End If
    End With
End Sub

Function ReverseDirection(Direction As Byte) As Byte
    Select Case Direction
    Case 0:    'Up
        ReverseDirection = 1
        Exit Function
    Case 1:
        ReverseDirection = 0
        Exit Function
    Case 2:
        ReverseDirection = 3
        Exit Function
    Case 3:
        ReverseDirection = 2
        Exit Function
    End Select

    ReverseDirection = 0
End Function

