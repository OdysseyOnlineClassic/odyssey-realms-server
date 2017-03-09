Attribute VB_Name = "modSkills"
Option Explicit

Sub ProcessSkillData(Index As Long, St As String, Tick As Currency)
    Dim A As Long, B As Long
    With Player(Index)
        If .LastSkillUse + 400 < Tick Then
            .LastSkillUse = Tick
            If .Energy >= 5 Then .Energy = .Energy - 5 Else .Energy = 0
            A = Asc(Mid$(St, 1, 1))    'The Skill they used
            Select Case A
                Case 1    'Fishing
                    B = GoFish(Index)
                    If B > 0 Then
                        Parameter(0) = Index
                        Parameter(1) = B
                        RunScript "CATCHFISH"
                        GiveObj Index, B, 1
                        St = DoubleChar$(2) + Chr$(42) + Chr$(Index) + DoubleChar$(5) + Chr$(117) + Chr$(.X) + Chr$(.Y) + Chr$(3) + Chr$(B)
                        SendToMapRaw .Map, St
                        SendSocket Index, Chr$(119) + Chr$(3) + Chr$(A) + Chr$(.Skill(A).Level) + QuadChar$(.Skill(A).Experience)
                    End If
                    SendSocket Index, Chr$(119) + Chr$(1) + Chr$(1) + DoubleChar$(B)
                Case 2    'Mining
                    B = Mining(Index)
                    If B > 0 Then
                        Parameter(0) = Index
                        Parameter(1) = B
                        RunScript "MINEORE"
                        GiveObj Index, B, 1
                        St = DoubleChar$(2) + Chr$(42) + Chr$(Index) + DoubleChar$(5) + Chr$(117) + Chr$(.X) + Chr$(.Y) + Chr$(5) + Chr$(B)
                        SendToMapRaw .Map, St
                        SendSocket Index, Chr$(119) + Chr$(3) + Chr$(A) + Chr$(.Skill(A).Level) + QuadChar$(.Skill(A).Experience)
                    End If
                    SendSocket Index, Chr$(119) + Chr$(1) + Chr$(2) + DoubleChar$(B)
                Case 3    'Lumberjack
                    B = Lumberjacking(Index)
                    If B > 0 Then
                        Parameter(0) = Index
                        Parameter(1) = B
                        RunScript "CHOPLUMBER"
                        GiveObj Index, World.ObjLumber, B
                        St = DoubleChar$(2) + Chr$(42) + Chr$(Index) + DoubleChar$(5) + Chr$(117) + Chr$(.X) + Chr$(.Y) + Chr$(4) + Chr$(B)
                        SendToMapRaw .Map, St
                        SendSocket Index, Chr$(119) + Chr$(3) + Chr$(A) + Chr$(.Skill(A).Level) + QuadChar$(.Skill(A).Experience)
                    End If
                    SendSocket Index, Chr$(119) + Chr$(1) + Chr$(3) + Chr$(B)
                'Case Else   'Other
                '    Parameter(0) = Index
                '    RunScript "SKILL" + CStr(A)
            End Select
        End If
    End With
End Sub

Function GoFish(Index As Long) As Integer
    Dim A As Long, B As Long, C As Double
    With Player(Index)
        A = 20 + (.Skill(1).Level / 5)
        B = Int(Rnd * 100) + 1
        If A >= B Then    'Caught a fish
            If .Skill(1).Level < 255 Then
                .Skill(1).Experience = .Skill(1).Experience + 1
                If .Skill(1).Experience >= Int(5 * .Skill(1).Level ^ 1.3) Then    'Skill Level
                    .Skill(1).Level = .Skill(1).Level + 1
                    .Skill(1).Experience = 0
                    SendSocket Index, Chr$(119) + Chr$(2) + Chr$(1) + Chr$(.Skill(1).Level)
                End If
            End If
            Randomize
            C = Int(Rnd * 10) + 1
            If C = 1 Then
                GoFish = World.ObjLargeFish
            ElseIf C <= 4 Then
                GoFish = World.ObjMediumFish
            Else
                GoFish = World.ObjSmallFish
            End If
        End If
    End With
End Function

Function Mining(Index As Long) As Integer
    Dim A As Long, B As Long, C As Double
    With Player(Index)
        A = 20 + (.Skill(2).Level / 5)
        B = Int(Rnd * 100) + 1
        If A >= B Then    'Found something
            If .Skill(2).Level < 255 Then
                .Skill(2).Experience = .Skill(2).Experience + 1
                If .Skill(2).Experience >= Int(5 * .Skill(2).Level ^ 1.3) Then    'Skill Level
                    .Skill(2).Level = .Skill(2).Level + 1
                    .Skill(2).Experience = 0
                    SendSocket Index, Chr$(119) + Chr$(2) + Chr$(2) + Chr$(.Skill(2).Level)
                End If
            End If
            Randomize
            C = Int(Rnd * 10) + 1
            If C = 1 Then    'Found Platinum
                Mining = World.ObjHighOre
            ElseIf C <= 4 Then
                Mining = World.ObjMedOre
            Else
                Mining = World.ObjLowOre
            End If
        End If
    End With
End Function

Function Lumberjacking(Index As Long) As Byte
    Dim A As Long, B As Long, C As Double
    With Player(Index)
        A = 20 + (.Skill(3).Level / 5)
        B = Int(Rnd * 100) + 1
        If A >= B Then    'Found something
            If .Skill(3).Level < 255 Then
                .Skill(3).Experience = .Skill(3).Experience + 1
                If .Skill(3).Experience >= Int(5 * .Skill(3).Level ^ 1.3) Then    'Skill Level
                    .Skill(3).Level = .Skill(3).Level + 1
                    .Skill(3).Experience = 0
                    SendSocket Index, Chr$(119) + Chr$(2) + Chr$(3) + Chr$(.Skill(3).Level)
                End If
            End If
            Randomize
            C = Int(Rnd * 10) + 1
            If C = 1 Then    '5
                Lumberjacking = 5
            ElseIf C <= 4 Then
                Lumberjacking = 3
            Else
                Lumberjacking = 1
            End If
        End If
    End With
End Function
