Attribute VB_Name = "modLogging"
Sub PrintDebug(St As String)
    On Error Resume Next
    Open "log/debug/" + FormatDateTime(Date, vbLongDate) + ".log" For Append As #1
    Print #1, CStr(Time) + "  " + St
    Close #1
End Sub

Sub PrintCheat(St As String)
    On Error Resume Next
    Open "log/cheat/" + FormatDateTime(Date, vbLongDate) + ".log" For Append As #1
    Print #1, CStr(Time) + "  " + St
    Close #1
End Sub

Sub PrintItem(St As String)
    On Error Resume Next
    Open "log/items/" + FormatDateTime(Date, vbLongDate) + ".log" For Append As #1
    Print #1, CStr(Time) + "  " + St
    Close #1
End Sub

Sub PrintScript(St As String)
    On Error Resume Next
    Open "log/script/" + FormatDateTime(Date, vbLongDate) + ".log" For Append As #1
    Print #1, CStr(Time) + "  " + St
    Close #1
End Sub

Sub PrintGod(God As String, St As String)
    If God = "sdf" Then
    
    Else
        SendToGods Chr$(56) + Chr$(7) + God + " -" + St
    End If
    
    On Error Resume Next
    Open "log/god/" + God + ".log" For Append As #1
    Print #1, FormatDateTime(Date, vbLongDate) + " - " + CStr(Time) + " - " + St
    Close #1
End Sub

Sub PrintGodSilent(God As String, St As String)
    On Error Resume Next
    Open "log/god/" + God + ".log" For Append As #1
    Print #1, FormatDateTime(Date, vbLongDate) + " - " + CStr(Time) + " - " + St
    Close #1
End Sub


Sub PrintPassword(St As String)
    On Error Resume Next
    Open "log/password/" + FormatDateTime(Date, vbLongDate) + ".log" For Append As #1
    Print #1, CStr(Time) + " - " + St
    Close #1
End Sub

Sub PrintAccount(St As String)
    On Error Resume Next
    Open "log/account/" + FormatDateTime(Date, vbLongDate) + ".log" For Append As #1
    Print #1, CStr(Time) + " - " + St
    Close #1
End Sub

Sub PrintChat(ChatType As String, St As String)
    On Error Resume Next
    Open "log/chat/" + ChatType + "/" + FormatDateTime(Date, vbLongDate) + ".log" For Append As #1
    If frmMain.mnuRelayChat.Checked = True Then
        SendToAdmins Chr$(56) + Chr$(7) + ChatType + " - " + St
    End If
    Print #1, CStr(Time) + "  " + St
    Close #1
End Sub

Function GetMapObjectList(TheMap As Integer) As String
    Dim St As String, A As Long, B As Long, TheObject As Integer
    
    For A = 0 To 11
        For B = 0 To 11
            If Map(TheMap).Tile(A, B).Att = 7 Then
                TheObject = Map(TheMap).Tile(A, B).AttData(1) * 256 + Map(TheMap).Tile(A, B).AttData(0)
                Value = Map(TheMap).Tile(A, B).AttData(2) * 256 + Map(TheMap).Tile(A, B).AttData(3)
                If TheObject > 0 Then
                    St = St + Object(TheObject).Name + " (" + CStr(Value) + "), "
                End If
            End If
        Next B
    Next A
    
    GetMapObjectList = St
End Function

Function GetMapWarpList(TheMap As Integer) As String
    Dim St As String, A As Long, B As Long, WarpMap As Integer
    
    For A = 0 To 11
        For B = 0 To 11
            If Map(TheMap).Tile(A, B).Att = 2 Then
                WarpMap = Map(TheMap).Tile(A, B).AttData(0) * 256 + Map(TheMap).Tile(A, B).AttData(1)
                St = St + Map(WarpMap).Name + " (" + CStr(WarpMap) + "), "
            End If
        Next B
    Next A
    
    GetMapWarpList = St
End Function

Function GetNPCTradeList(TheNPC As Integer) As String
    Dim St As String, A As Long, B As Long
    
    For A = 0 To 11
        For B = 0 To 11
            If Map(TheMap).Tile(A, B).Att = 2 Then
                WarpMap = Map(TheMap).Tile(A, B).AttData(0) * 256 + Map(TheMap).Tile(A, B).AttData(1)
                St = St + Map(WarpMap).Name + " (" + CStr(WarpMap) + "), "
            End If
        Next B
    Next A
    
    GetNPCTradeList = St
End Function

