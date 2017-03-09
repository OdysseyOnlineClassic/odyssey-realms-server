Attribute VB_Name = "modMagic"
Sub ProcessMagicData(Index As Long, St As String)
Dim A As Long, B As Long
With Player(Index)
    A = Asc(Mid$(St, 1, 1)) 'The Skill they used
    Select Case .Class
        Case 2 'Mage
            Select Case A
                Case 1 'Seethe (Level 1)
                    If .Mana >= 7 Then
                        SendToMap .Map, Chr$(99) + Chr$(6) + Chr$(.D) + Chr$(.X) + Chr$(.Y) + Chr$(13) + Chr$(Index) + Chr$(A)
                        .Mana = .Mana - 7
                        SendSocket Index, Chr$(48) + Chr$(.Mana)
                    End If
                Case 2 'Explosion (Level 1)
                        If .Mana >= 12 Then
                            .Mana = .Mana - 12
                        End If
                Case 3 'Disrupt (Level 3)
                    If .level >= 3 Then
                        If .Mana >= 12 Then
                            SendToMap .Map, Chr$(99) + Chr$(6) + Chr$(.D) + Chr$(.X) + Chr$(.Y) + Chr$(12) + Chr$(Index) + Chr$(A)
                            .Mana = .Mana - 12
                            SendSocket Index, Chr$(48) + Chr$(.Mana)
                        End If
                    End If
                Case 4 'Conflagration (Level 5)
                    If .level >= 5 Then
                    
                    End If
            End Select
    End Select
End With
End Sub

Sub SpellDamagePlayer(Index As Long, St As String)
Dim A As Long, B As Long, C As Long
With Player(Index)
    A = Asc(Mid$(St, 1, 1))
    B = Asc(Mid$(St, 2, 1))
    Select Case .Class
        Case 2 'Mage
            Select Case A
                Case 1 'Seethe (Level 1)
                    If .X = Player(B).X Or .Y = Player(B).Y Then
                        C = Int(Rnd * 6) + 1
                        C = C + 8
                        MagicAttackPlayer Index, B, C
                    End If
                'Case 2 'Explosion (Level 1)
                Case 3 'Disrupt (Level 3)
                    If .level >= 3 Then
                        If .X = Player(B).X Or .Y = Player(B).Y Then
                            C = Int(Rnd * 4) + 1
                            C = C + 12
                            MagicAttackPlayer Index, B, C
                        End If
                    End If
                'Case 4 'Conflagration (Level 3)
            End Select
    End Select
End With
End Sub

Sub SpellDamageMonster(Index As Long, St As String)
Dim A As Long, B As Long, C As Long
With Player(Index)
    A = Asc(Mid$(St, 1, 1))
    B = Asc(Mid$(St, 2, 1))
    Select Case .Class
        Case 2 'Mage
            Select Case A
                Case 1 'Seethe (Level 1)
                    If .X = Map(.Map).Monster(B).X Or .Y = Map(.Map).Monster(B).Y Then
                        C = Int(Rnd * 6) + 1
                        C = C + 8
                        MagicAttackMonster Index, B, C
                    End If
                'Case 2 'Explosion (Level 1)
                Case 3 'Disrupt (Level 3)
                    If .level >= 3 Then
                        If .X = Map(.Map).Monster(B).X Or .Y = Map(.Map).Monster(B).Y Then
                            C = Int(Rnd * 4) + 1
                            C = C + 12
                            MagicAttackMonster Index, B, C
                        End If
                    End If
                'Case 4 'Conflagration (Level 3)
            End Select
    End Select
End With
End Sub
