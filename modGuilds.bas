Attribute VB_Name = "modGuilds"
Sub DeleteGuild(Index As Long, Reason As Byte)
    Dim A As Long, B As Long, C As Long

    With Guild(Index)
        If .Name <> "" Then
            .Name = ""
            GuildRS.Bookmark = .Bookmark
            GuildRS.Delete
        End If

        UserRS.Index = "Name"
        For A = 0 To 19
            With .Member(A)
                If .Name <> "" Then
                    B = FindPlayer(.Name)
                    If B > 0 Then
                        With Player(B)
                            .Guild = 0
                            .GuildRank = 0
                            If Guild(Index).Sprite > 0 Then
                                .Sprite = .Class * 2 + .Gender - 1
                                SendToMap .Map, Chr$(63) + Chr$(B) + DoubleChar$(CLng(.Sprite))
                            End If
                            SendSocket B, Chr$(75) + Chr$(Reason)
                            SendAllBut B, Chr$(73) + Chr$(B) + vbNullChar
                        End With
                    ElseIf Guild(Index).Sprite > 0 Then
                        UserRS.Seek "=", .Name
                        If UserRS.NoMatch = False Then
                            B = UserRS!Class * 2 + UserRS!Gender - 1
                            If B >= 1 And B <= MaxSprite Then
                                UserRS.Edit
                                UserRS!Sprite = B
                                UserRS.Update
                            End If
                        End If
                    End If
                End If
            End With
        Next A
    End With

    'Check if other guilds have declarations
    For A = 1 To MaxGuilds
        With Guild(A)
            If .Name <> "" Then
                C = 0
                For B = 0 To DeclarationCount
                    With .Declaration(B)
                        If .Guild = Index Then
                            .Guild = 0
                            SendToGuild A, Chr$(71) + Chr$(B) + vbNullChar + vbNullChar    'Declaration Data
                            C = 1
                        End If
                    End With
                Next B
                If C = 1 Then
                    GuildRS.Bookmark = .Bookmark
                    GuildRS.Edit
                    For B = 0 To DeclarationCount
                        With .Declaration(B)
                            GuildRS("DeclarationGuild" + CStr(B)) = .Guild
                            GuildRS("DeclarationType" + CStr(B)) = .Type
                        End With
                    Next B
                    GuildRS.Update
                End If
            End If
        End With
    Next A

    'Erase Join Requests
    For A = 1 To MaxUsers
        With Player(A)
            If .JoinRequest = Index Then .JoinRequest = 0
        End With
    Next A

    SendAll Chr$(70) + Chr$(Index)    'Erase Guild
End Sub

Function FindGuildMember(ByVal Name As String, GuildNum As Long) As Long
    Name = UCase$(Name)
    Dim A As Long
    With Guild(GuildNum)
        For A = 0 To 19
            If UCase$(.Member(A).Name) = Name Then
                FindGuildMember = A
                Exit Function
            End If
        Next A
    End With
    FindGuildMember = -1
End Function

Function FreeGuildDeclarationNum(GuildNum As Long) As Long
    Dim A As Long
    With Guild(GuildNum)
        For A = 0 To DeclarationCount
            If .Declaration(A).Guild = 0 Then
                FreeGuildDeclarationNum = A
                Exit Function
            End If
        Next A
    End With
    FreeGuildDeclarationNum = -1
End Function

Function FreeGuildMemberNum(GuildNum As Long)
    Dim A As Long
    
    If GuildMemberCount(GuildNum) > World.GuildMaxMembers Then
        FreeGuildMemberNum = -1
        Exit Function
    End If
    
    With Guild(GuildNum)
        For A = 0 To World.GuildMaxMembers - 1
            If .Member(A).Name = "" Then
                FreeGuildMemberNum = A
                Exit Function
            End If
        Next A
    End With
    
    FreeGuildMemberNum = -1
End Function

Function GuildMemberCount(GuildNum As Long)
    Dim A As Long, B As Long
    With Guild(GuildNum)
        For A = 0 To 19
            If Not .Member(A).Name = "" Then
                B = B + 1
            End If
        Next A
    End With

    GuildMemberCount = B
End Function

Function FreeGuildNum() As Long
    Dim A As Long
    For A = 1 To MaxGuilds
        If Guild(A).Name = "" Then
            FreeGuildNum = A
            Exit Function
        End If
    Next A
End Function

Sub SendToGuild(GuildNum As Long, St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Guild = GuildNum Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub
Sub SendToGuildAllBut(Index As Long, GuildNum As Long, St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Guild = GuildNum And Index <> A Then
                SendSocket A, St
            End If
        End With
    Next A
End Sub

Function GetGuildUpkeep(Index As Long) As Long
    Dim A As Long, B As Long
    A = 0
    With Guild(Index)
        For B = 0 To 19
            If .Member(B).Name <> "" Then
                A = A + World.GuildUpkeepMembers
            End If
        Next B
        If .Hall > 0 Then
            A = A + Hall(.Hall).Upkeep
        End If
        If .Sprite > 0 Then
            A = A + World.GuildUpkeepSprite
        End If
    End With

    GetGuildUpkeep = A
End Function
