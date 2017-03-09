Attribute VB_Name = "modLists"
Public ObjectVersionList As String
Public MonsterVersionList As String
Public NPCVersionList As String
Public HallVersionList As String
Public MagicVersionList As String
Public PrefixVersionList As String
Public SuffixVersionList As String

Public Sub GenerateObjectVersionList()
    ObjectVersionList = vbNullString
    For A = 1 To MaxObjects
        ObjectVersionList = ObjectVersionList + Chr$(Object(A).Version)
    Next A
End Sub

Public Sub GenerateMonsterVersionList()
    MonsterVersionList = vbNullString
    For A = 1 To MaxTotalMonsters
        MonsterVersionList = MonsterVersionList + Chr$(Monster(A).Version)
    Next A
End Sub

Public Sub GenerateNPCVersionList()
    NPCVersionList = vbNullString
    For A = 1 To MaxNPCs
        NPCVersionList = NPCVersionList + Chr$(NPC(A).Version)
    Next A
End Sub

Public Sub GenerateHallVersionList()
    HallVersionList = vbNullString
    For A = 1 To MaxHalls
        HallVersionList = HallVersionList + Chr$(Hall(A).Version)
    Next A
End Sub

Public Sub GenerateMagicVersionList()
    MagicVersionList = vbNullString
    For A = 1 To MaxMagic
        MagicVersionList = MagicVersionList + Chr$(Magic(A).Version)
    Next A
End Sub

Public Sub GeneratePrefixVersionList()
    PrefixVersionList = vbNullString
    For A = 1 To MaxModifications
        PrefixVersionList = PrefixVersionList + Chr$(ItemPrefix(A).Version)
    Next A
End Sub

Public Sub GenerateSuffixVersionList()
    SuffixVersionList = vbNullString
    For A = 1 To MaxModifications
        SuffixVersionList = SuffixVersionList + Chr$(ItemSuffix(A).Version)
    Next A
End Sub
