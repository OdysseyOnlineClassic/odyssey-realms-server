Attribute VB_Name = "modSecurity"
Option Explicit

Function GetFreeBugSlot() As Long
Dim A As Long
    For A = 1 To 500
        If Bug(A).Status = 0 Then
            GetFreeBugSlot = A
            Exit Function
        End If
    Next A
End Function

Sub CheckPingSpeed(Index As Long)
Dim A As Long, B As Long
    With Player(Index)
        For A = 1 To 5
            If .Ping(A) = 0 Then Exit Sub 'Only check if all of the ping arrays has been filled with values to compare.
        Next A
        For A = 1 To 5
            B = B + .Ping(A)
        Next A
        B = 5000 - (B / 5) 'B/5 = average interval frequency. Near to or above 5000 is in sync, if significantly lower frequency then possible speed hack
        If B >= 50 Then
            .SpeedStrikes = .SpeedStrikes + 1
            If .SpeedStrikes >= 2 Then
                SendToGods Chr$(16) + Chr$(0) + "Warning: Possible speed hack detected from player - " + .Name + " " + CStr(B) + "ms faster than the required average ping frequency!"
            ElseIf .SpeedStrikes >= 5 Then
                BootPlayer Index, 0, "Speed Hack Detected!"
            End If
        Else
            If .SpeedStrikes > 0 Then .SpeedStrikes = .SpeedStrikes - 1
        End If
    End With
End Sub

Function CheckBan(Index As Long, PlayerName As String, ComputerID As String, IPAddress As String) As Boolean
    Dim BanNum As Long
    BanNum = FindBan(PlayerName, ComputerID, IPAddress)
    If BanNum > 0 Then
        With Ban(BanNum)
            If CLng(Date) >= .UnbanDate Then
                'Unban
                .Name = ""
                .ComputerID = ""
                .IPAddress = ""
                .InUse = False
                .Banner = ""
                .Name = ""
                .Reason = ""
                .UnbanDate = 0
                BanRS.Seek "=", BanNum
                If BanRS.NoMatch = False Then
                    BanRS.Delete
                End If
                CheckBan = False
                Exit Function
            Else
                SendSocket Index, Chr$(0) + Chr$(3) + QuadChar(Ban(BanNum).UnbanDate) + Ban(BanNum).Reason
                SendToGods Chr$(16) + Chr$(0) + "Banned::" + PlayerName + "::" + .Name
                PrintCheat "Banned::" + PlayerName + "::" + .Name
                Player(Index).Mode = modeBanned
                CheckBan = True
                Exit Function
            End If
        End With
    End If

    CheckBan = False
End Function

Function FindBan(PlayerName As String, ComputerID As String, IPAddress) As Long
    Dim A As Long

    For A = 1 To 50
        If Ban(A).InUse = True Then
            If UCase$(Ban(A).Name) = UCase$(PlayerName) Then
                FindBan = A
                Exit Function
            ElseIf UCase$(Ban(A).ComputerID) = UCase$(ComputerID) Then
                FindBan = A
                Exit Function
            ElseIf UCase$(Ban(A).IPAddress) = UCase$(IPAddress) Then
                FindBan = A
                Exit Function
            End If
        End If
    Next A
End Function

Function BanPlayer(A As Long, Index As Long, NumDays As Long, Reason As String, Banner As String) As Boolean
    Dim C As Long

    With Player(A)
        If Not .Access = 4 Then
            C = FreeBanNum
            If C >= 1 Then
                If CheckBan(A, Player(A).Name, Player(A).ComputerID, Player(A).IP) = False Then
                    With Ban(C)
                        .Name = Player(A).Name
                        If Len(.Name) < 2 Then .Name = "null2523"
                        .Reason = Reason
                        .Banner = Banner
                        .ComputerID = Player(A).ComputerID
                        .IPAddress = Player(A).IP
                        .InUse = True
                        .UnbanDate = CLng(Date) + NumDays
                        BanRS.Seek "=", C
                        If BanRS.NoMatch = True Then
                            BanRS.AddNew
                            BanRS!number = C
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
                        SendSocket A, Chr$(67) + Chr$(Index) + .Reason
                        If Player(A).Mode = modePlaying Then
                            SendAllBut A, Chr$(66) + Chr$(A) + Chr$(Index) + .Reason
                        End If
                        AddSocketQue A
                        BanPlayer = True
                    End With
                End If
            End If
        End If
    End With
End Function

Sub BootPlayer(A As Long, Index As Long, Reason As String)
    Dim D As Long
    For D = 1 To 80
        If CloseSocketQue(D) = A Then Exit Sub
    Next D

    With Player(A)
        If .InUse = True And Not .Access = 4 Then
            If Reason <> "" Then
                SendSocket A, Chr$(67) + Chr$(Index) + Reason
                If .Mode = modePlaying Then
                    SendAllBut A, Chr$(68) + Chr$(A) + Chr$(Index) + Reason
                Else
                    SendToGods Chr$(56) + Chr$(15) + "User " + Chr$(34) + .User + Chr$(34) + " with name " + Chr$(34) + .Name + Chr$(34) + " has been booted: " + Reason
                End If
                AddSocketQue A
            Else
                SendSocket A, Chr$(67) + Chr$(Index)
                If .Mode = modePlaying Then
                    SendAllBut A, Chr$(68) + Chr$(A) + Chr$(Index)
                Else
                    SendAllBut A, Chr$(56) + Chr$(15) + "User " + Chr$(34) + .User + Chr$(34) + " with name " + Chr$(34) + .Name + Chr$(34) + " has been booted!"
                End If
                AddSocketQue A
            End If
        End If
    End With
End Sub

Sub Hacker(Index As Long, Code As String)
    BanPlayer Index, 0, 3, "Possible Hacking Attempt: Code '" + Code + "' from IP '" + Player(Index).IP + "'", "Server"
    PrintLog Player(Index).Name & "    Possible Hacking Attempt: Code '" + Code + "' from IP '" + Player(Index).IP + "'"
    PrintCheat Player(Index).Name & "    Possible Hacking Attempt: Code '" + Code + "' from IP '" + Player(Index).IP + "'"
End Sub

Function ReadUniqID() As String
    Dim lpReturnedString As String, Valid As Long
    lpReturnedString = Space$(256)
    Valid = GetPrivateProfileString&("UniqID", "ID", "", lpReturnedString, 256, "froogle")
    ReadUniqID = Left$(lpReturnedString, Valid)
End Function

Function WriteUniqID(UniqID As String) As String
    WritePrivateProfileString "UniqID", "ID", UniqID, "froogle"
End Function

