Attribute VB_Name = "modDatabase"
Option Explicit

'Database Objects
Public WS As Workspace
Public DB As Database
Public UserRS As Recordset
Public NPCRS As Recordset
Public MonsterRS As Recordset
Public ObjectRS As Recordset
Public DataRS As Recordset
Public MapRS As Recordset
Public GuildRS As Recordset
Public BanRS As Recordset
Public HallRS As Recordset
Public ScriptRS As Recordset
Public MagicRS As Recordset
Public PrefixRS As Recordset
Public SuffixRS As Recordset

Sub LoadDatabase()
    Dim A As Long, B As Long, St As String, St2 As String, bDataRSError As Boolean

    Set WS = DBEngine.Workspaces(0)
    If Exists("server.dat") Then
        If Startup = True Then
            frmLoading.lblStatus = "Opening Server Database.."
            frmLoading.lblStatus.Refresh
        End If
        If Exists("server.tmp") Then Kill "server.tmp"
        Name "server.dat" As "server.tmp"
        CompactDatabase "server.tmp", "server.dat", , 0, ";pwd=ontario"
        Set DB = WS.OpenDatabase("server.dat", 0, False, ";pwd=ontario")
        Kill "server.tmp"
    Else
        If Startup = True Then
            frmLoading.lblStatus = "Creating Server Database.."
            frmLoading.lblStatus.Refresh
        End If
        CreateDatabase
    End If

    Err.Clear
    Set UserRS = DB.TableDefs("Accounts").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateAccountsTable
        Set UserRS = DB.TableDefs("Accounts").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set NPCRS = DB.TableDefs("NPCs").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateNPCsTable
        Set NPCRS = DB.TableDefs("NPCs").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set MonsterRS = DB.TableDefs("Monsters").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateMonstersTable
        Set MonsterRS = DB.TableDefs("Monsters").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set ObjectRS = DB.TableDefs("Objects").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateObjectsTable
        Set ObjectRS = DB.TableDefs("Objects").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set MapRS = DB.TableDefs("Maps").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateMapsTable
        Set MapRS = DB.TableDefs("Maps").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set BanRS = DB.TableDefs("Bans").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateBansTable
        Set BanRS = DB.TableDefs("Bans").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set GuildRS = DB.TableDefs("Guilds").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateGuildsTable
        Set GuildRS = DB.TableDefs("Guilds").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set HallRS = DB.TableDefs("Halls").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateHallsTable
        Set HallRS = DB.TableDefs("Halls").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set MagicRS = DB.TableDefs("Magic").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateMagicTable
        Set MagicRS = DB.TableDefs("Magic").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set ScriptRS = DB.TableDefs("Scripts").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateScriptsTable
        Set ScriptRS = DB.TableDefs("Scripts").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set PrefixRS = DB.TableDefs("Prefix").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreatePrefixTable
        Set PrefixRS = DB.TableDefs("Prefix").OpenRecordset(dbOpenTable)
    End If

    Err.Clear
    Set SuffixRS = DB.TableDefs("Suffix").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateSuffixTable
        Set SuffixRS = DB.TableDefs("Suffix").OpenRecordset(dbOpenTable)
    End If

    On Error GoTo 0

    UserRS.Index = "User"
    ObjectRS.Index = "Number"
    NPCRS.Index = "Number"
    MonsterRS.Index = "Number"
    MapRS.Index = "Number"
    BanRS.Index = "Number"
    GuildRS.Index = "Number"
    HallRS.Index = "Number"
    MagicRS.Index = "Number"
    ScriptRS.Index = "Name"
    PrefixRS.Index = "Number"
    SuffixRS.Index = "Number"

    If Startup = True Then
        frmLoading.lblStatus = "Loading World Data.."
        frmLoading.lblStatus.Refresh
    End If

ReloadData:
    Set DataRS = DB.TableDefs("Data").OpenRecordset(dbOpenTable)
    'Check if World Data exists
    If DataRS.RecordCount = 0 Then
        'Create default data
        With DataRS
            .AddNew
            !BackupInterval = 5
            !MapResetTime = 120000
            !ObjResetTime = 300000

            !ObjMoney = 6
            !ObjSmallFish = 153
            !ObjMediumFish = 152
            !ObjLargeFish = 148
            !ObjLowOre = 216
            !ObjMedOre = 217
            !ObjHighOre = 218
            !ObjLumber = 215

            !MaxLevel = 100
            !EliteLevel = 60
            !DeathTime = 15
            
            !GuildUpkeepMembers = 1000
            !GuildUpkeepSprite = 1000

            !ServerPort = 5750

            !Cost_Per_Durability = 100
            !Cost_Per_Strength = 100
            !Cost_Per_Modifier = 100

            !GuildJoinLevel = 1
            !GuildJoinPrice = 1
            !GuildNewLevel = 1
            !GuildNewPrice = 1
            !GuildMaxMembers = 20

            !StatStrength = 10
            !StatEndurance = 10
            !StatIntelligence = 10
            !StatConcentration = 10
            !StatConstitution = 10
            !StatStamina = 10
            !StatWisdom = 10

            !MOTD = vbNullString
            !LastUpdate = CLng(Date)
            !ObjectData = vbNullString
            !flags = String$(1024, 0)
            For A = 0 To 4
                .Fields("StartLocationX" + CStr(A)) = 5
                .Fields("StartLocationY" + CStr(A)) = 5
                .Fields("StartLocationMap" + CStr(A)) = 1
                .Fields("StartLocationMessage" + CStr(A)) = vbNullString
            Next A
            For A = 1 To 8
                .Fields("StartingObj" + CStr(A)) = 0
                .Fields("StartingObjVal" + CStr(A)) = 0
            Next A
            .Update
            .MoveFirst
        End With
    End If

    'Load World Data
    LoadObjectData DataRS!ObjectData

    With World
        .BackupInterval = DataRS!BackupInterval
        .MapResetTime = DataRS!MapResetTime
        .ObjResetTime = DataRS!ObjResetTime

        'Objects
        .ObjMoney = DataRS!ObjMoney
        .ObjSmallFish = DataRS!ObjSmallFish
        .ObjMediumFish = DataRS!ObjMediumFish
        .ObjLargeFish = DataRS!ObjLargeFish
        .ObjLumber = DataRS!ObjLumber
        .ObjLowOre = DataRS!ObjLowOre
        .ObjMedOre = DataRS!ObjMedOre
        .ObjHighOre = DataRS!ObjHighOre

        'Max Level
        .MaxLevel = DataRS!MaxLevel
        'Death Time
        .DeathTime = DataRS!DeathTime

        'Reparing
        .Cost_Per_Durability = DataRS!Cost_Per_Durability
        .Cost_Per_Modifier = DataRS!Cost_Per_Modifier
        .Cost_Per_Strength = DataRS!Cost_Per_Strength

        'Guilds
        .GuildJoinLevel = DataRS!GuildJoinLevel
        .GuildJoinPrice = DataRS!GuildJoinPrice
        .GuildNewLevel = DataRS!GuildNewLevel
        .GuildNewPrice = DataRS!GuildNewPrice
        .GuildMaxMembers = DataRS!GuildMaxMembers
        If Not IsNull(DataRS!GuildUpkeepMembers) Then .GuildUpkeepMembers = DataRS!GuildUpkeepMembers Else World.GuildUpkeepMembers = 0
        If Not IsNull(DataRS!GuildUpkeepSprite) Then .GuildUpkeepSprite = DataRS!GuildUpkeepSprite Else World.GuildUpkeepSprite = 0

        'Ports
        .ServerPort = DataRS!ServerPort

        'Stats
        .StatStrength = DataRS!StatStrength
        .StatEndurance = DataRS!StatEndurance
        .StatIntelligence = DataRS!StatIntelligence
        .StatConcentration = DataRS!StatConcentration
        .StatConstitution = DataRS!StatConstitution
        .StatStamina = DataRS!StatStamina
        .StatWisdom = DataRS!StatWisdom

        .LastUpdate = DataRS!LastUpdate
        .MOTD = DataRS!MOTD
        St = DataRS!flags
        For A = 0 To 255
            .Flag(A) = Asc(Mid$(St, A * 4 + 1, 1)) * 16777216 + Asc(Mid$(St, A * 4 + 2, 1)) * 65536 + Asc(Mid$(St, A * 4 + 3, 1)) * 256& + Asc(Mid$(St, A * 4 + 4, 1))
        Next A
        For A = 0 To 4
            .StartLocation(A).X = DataRS("StartLocationX" + CStr(A))
            .StartLocation(A).Y = DataRS("StartLocationY" + CStr(A))
            .StartLocation(A).Map = DataRS("StartLocationMap" + CStr(A))
            .StartLocation(A).Message = DataRS("StartLocationMessage" + CStr(A))
        Next A

        For A = 1 To 8
            .StartObjects(A) = DataRS("StartingObj" + CStr(A))
            .StartObjValues(A) = DataRS("StartingObjVal" + CStr(A))
        Next A
    End With

    On Error GoTo 0
    If bDataRSError Then
        MsgBox "There was an error loading the server options.  The database will be rebuilt, but some data may be lost.", vbOKOnly, TitleString
        DataRS.Close
        DB.TableDefs.Delete "Data"
        CreateDataTable
        bDataRSError = False
        GoTo ReloadData
    End If

    If World.LastUpdate > CLng(Date) Or Abs(World.LastUpdate - CLng(Date)) >= 30 Then
        If MsgBox("Please verify that your system date and time is set correctly -- click ok to go on.", vbOKCancel, TitleString) = vbCancel Then
            ShutdownServer
            End
        End If
    End If
    
    CreateClassData

    If Startup = True Then
        frmLoading.lblStatus = "Loading Guilds.."
        frmLoading.lblStatus.Refresh
    End If

    If GuildRS.BOF = False Then
        GuildRS.MoveFirst
        While GuildRS.EOF = False
            A = GuildRS!number
            If A > 0 Then
                With Guild(A)
                    .Name = GuildRS!Name
                    If .Name = "" Then
                        DeleteGuild A, 3
                    Else
                        .Bank = GuildRS!Bank
                        .DueDate = GuildRS!DueDate
                        .Hall = GuildRS!Hall
                        .Sprite = GuildRS!Sprite
                        .Kills = GuildRS!Kills
                        .Deaths = GuildRS!Deaths
                        .CreationDate = GuildRS!CreationDate
                        .MOTD = GuildRS!MOTD
                        .MOTDCreator = GuildRS!MOTDCreator
                        .MOTDDate = GuildRS!MOTDDate
                        For B = 0 To 19
                            .Member(B).Name = GuildRS("MemberName" + CStr(B))
                            .Member(B).Rank = GuildRS("MemberRank" + CStr(B))
                            .Member(B).JoinDate = GuildRS("MemberJoinDate" + CStr(B))
                            .Member(B).Kills = GuildRS("MemberKills" + CStr(B))
                            .Member(B).Deaths = GuildRS("MemberDeaths" + CStr(B))
                        Next B
                        For B = 0 To DeclarationCount
                            .Declaration(B).Guild = GuildRS("DeclarationGuild" + CStr(B))
                            .Declaration(B).Type = GuildRS("DeclarationType" + CStr(B))
                            .Declaration(B).Date = GuildRS("DeclarationDate" + CStr(B))
                            .Declaration(B).Kills = GuildRS("DeclarationKills" + CStr(B))
                            .Declaration(B).Deaths = GuildRS("DeclarationDeaths" + CStr(B))
                        Next B
                        .MemberCount = CountGuildMembers(A)
                        .Bookmark = GuildRS.Bookmark
                    End If
                End With
            End If
            GuildRS.MoveNext
        Wend
    End If

    If Startup = True Then
        frmLoading.lblStatus = "Checking Accounts.."
        frmLoading.lblStatus.Refresh
    End If

    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If CLng(Date) - UserRS!LastPlayed >= 60 Then
                If UserRS!Level < 20 Then
                    DeleteAccount
                ElseIf UserRS!Name = vbNullString Then
                    DeleteAccount
                Else
                    A = FindGuild(UserRS!Name)
                    If A > 0 Then
                        RemoveFromGuild UserRS!Name, A
                    End If
                End If
            Else
                If UserRS!Name = "" Then
                    DeleteAccount
                End If
            End If
            UserRS.MoveNext
        Wend
    End If

    If Startup = True Then
        frmLoading.lblStatus = "Loading Halls.."
        frmLoading.lblStatus.Refresh
    End If

    If HallRS.BOF = False Then
        HallRS.MoveFirst
        While HallRS.EOF = False
            A = HallRS!number
            If A > 0 Then
                With Hall(A)
                    .Name = HallRS!Name
                    .Price = HallRS!Price
                    .Upkeep = HallRS!Upkeep
                    .StartLocation.Map = HallRS!StartLocationMap
                    .StartLocation.X = HallRS!StartLocationX
                    .StartLocation.Y = HallRS!StartLocationY
                    .Version = HallRS!Version
                End With
            End If
            HallRS.MoveNext
        Wend
    End If
    GenerateHallVersionList

    If Startup = True Then
        frmLoading.lblStatus = "Loading Objects.."
        frmLoading.lblStatus.Refresh
    End If

    If ObjectRS.BOF = False Then
        ObjectRS.MoveFirst
        While ObjectRS.EOF = False
            A = ObjectRS!number
            If A > 0 Then
                If ObjectRS!Name = "" Then
                    ObjectRS.Edit
                    ObjectRS.Delete
                Else
                    With Object(A)
                        .Name = ObjectRS!Name
                        .Picture = ObjectRS!Picture
                        .Type = ObjectRS!Type
                        .flags = ObjectRS!flags
                        .Data(0) = ObjectRS!Data1
                        .Data(1) = ObjectRS!Data2
                        .Data(2) = ObjectRS!Data3
                        .Data(3) = ObjectRS!Data4
                        If Not IsNull(ObjectRS!ClassReq) Then .ClassReq = ObjectRS!ClassReq Else .ClassReq = 0
                        If Not IsNull(ObjectRS!LevelReq) Then .LevelReq = ObjectRS!LevelReq Else .LevelReq = 0
                        If Not IsNull(ObjectRS!Version) Then .Version = ObjectRS!Version Else .Version = 0
                        If Not IsNull(ObjectRS!SellPrice) Then .SellPrice = ObjectRS!SellPrice Else .SellPrice = 0
                    End With
                End If
            End If
            ObjectRS.MoveNext
        Wend
    End If
    GenerateObjectVersionList

    If Startup = True Then
        frmLoading.lblStatus = "Loading NPCs.."
        frmLoading.lblStatus.Refresh
    End If

    If NPCRS.BOF = False Then
        NPCRS.MoveFirst
        While NPCRS.EOF = False
            A = NPCRS!number
            If A > 0 Then
                If NPCRS!Name = "" Then
                    NPCRS.Edit
                    NPCRS.Delete
                Else
                    With NPC(A)
                        .Name = NPCRS!Name
                        .JoinText = NPCRS!JoinText
                        .LeaveText = NPCRS!LeaveText
                        If Not IsNull(NPCRS!Version) Then .Version = NPCRS!Version Else .Version = 0
                        For B = 0 To 4
                            .SayText(B) = NPCRS("SayText" + CStr(B))
                        Next B
                        For B = 0 To 9
                            With .SaleItem(B)
                                .GiveObject = NPCRS("GiveObject" + CStr(B))
                                .GiveValue = NPCRS("GiveValue" + CStr(B))
                                .TakeObject = NPCRS("TakeObject" + CStr(B))
                                .TakeValue = NPCRS("TakeValue" + CStr(B))
                            End With
                        Next B
                        .flags = NPCRS!flags
                    End With
                End If
            End If
            NPCRS.MoveNext
        Wend
    End If
    GenerateNPCVersionList

    If Startup = True Then
        frmLoading.lblStatus = "Loading Monsters.."
        frmLoading.lblStatus.Refresh
    End If

    If MonsterRS.BOF = False Then
        MonsterRS.MoveFirst
        While MonsterRS.EOF = False
            A = MonsterRS!number
            If A > 0 Then
                If MonsterRS!Name = "" Then
                    MonsterRS.Edit
                    MonsterRS.Delete
                Else
                    With Monster(A)
                        .Name = MonsterRS!Name
                        .Sprite = MonsterRS!Sprite
                        .HP = MonsterRS!HP
                        .Strength = MonsterRS!Strength
                        .Armor = MonsterRS!Armor
                        .Speed = MonsterRS!Speed
                        .Sight = MonsterRS!Sight
                        .Agility = MonsterRS!Agility
                        .flags = MonsterRS!flags
                        .Object(0) = MonsterRS!Object0
                        .Value(0) = MonsterRS!Value0
                        .Object(1) = MonsterRS!Object1
                        .Value(1) = MonsterRS!Value1
                        .Object(2) = MonsterRS!Object2
                        .Value(2) = MonsterRS!Value2
                        If Not IsNull(MonsterRS!Experience) Then .Experience = MonsterRS!Experience Else .Experience = 0
                        If Not IsNull(MonsterRS!MagicDefense) Then .MagicDefense = MonsterRS!MagicDefense Else .MagicDefense = 0
                        If Not IsNull(MonsterRS!Version) Then .Version = MonsterRS!Version Else .Version = 0
                    End With
                End If
            End If
            MonsterRS.MoveNext
        Wend
    End If
    GenerateMonsterVersionList

    If Startup = True Then
        frmLoading.lblStatus = "Loading Ban List.."
        frmLoading.lblStatus.Refresh
    End If

    If BanRS.BOF = False Then
        BanRS.MoveFirst
        While BanRS.EOF = False
            A = BanRS!number
            If A > 0 Then
                With Ban(A)
                    If BanRS!UnbanDate <= CLng(Date) Then
                        BanRS.Delete
                    Else
                        .Name = BanRS!Name
                        .Reason = BanRS!Reason
                        .UnbanDate = BanRS!UnbanDate
                        .Banner = BanRS!Banner
                        If Not IsNull(BanRS!IPAddress) Then .IPAddress = BanRS!IPAddress
                        If Not IsNull(BanRS!ComputerID) Then .ComputerID = BanRS!ComputerID
                        .InUse = True
                    End If
                End With
            End If
            BanRS.MoveNext
        Wend
    End If

    If Startup = True Then
        frmLoading.lblStatus = "Loading Magic.."
        frmLoading.lblStatus.Refresh
    End If

    If MagicRS.BOF = False Then
        MagicRS.MoveFirst
        While MagicRS.EOF = False
            A = MagicRS!number
            If A > 0 Then
                If MagicRS!Name = "" Or MagicRS!Level = 255 Then
                    MagicRS.Edit
                    MagicRS.Delete
                Else
                    With Magic(A)
                        .Name = MagicRS!Name
                        .Class = MagicRS!Class
                        .Level = MagicRS!Level
                        .Version = MagicRS!Version
                        '.Icon = MagicRS!Icon
                        '.IconType = MagicRS!IconType
                        '.CastTimer = MagicRS!CastTimer
                        If Not IsNull(MagicRS!Description) Then
                            .Description = MagicRS!Description
                        End If
                    End With
                End If
            End If
            If Magic(A).Icon = 0 Then Magic(A).Icon = 1
            If Magic(A).IconType = 0 Then Magic(A).IconType = 1
            If Magic(A).CastTimer = 0 Then Magic(A).CastTimer = 1
            MagicRS.MoveNext
        Wend
    End If
    GenerateMagicVersionList

    If Startup = True Then
        frmLoading.lblStatus = "Loading Maps.."
        frmLoading.lblStatus.Refresh
    End If

    If MapRS.BOF = False Then
        MapRS.MoveFirst
        While MapRS.EOF = False
            A = MapRS!number
            If A > 0 And A <= MaxMaps Then
                St = MapRS!Data
                St2 = UncompressString(St)
                
                'Map Conversion Line
                'LoadMapOld1997 A, St
                LoadMap A, St2
                
                If Map(A).Name = "" Then
                    MapRS.Edit
                    MapRS.Delete
                End If
            End If
            MapRS.MoveNext
        Wend
    End If

    If Startup = True Then
        frmLoading.lblStatus = "Loading Prefixes.."
        frmLoading.lblStatus.Refresh
    End If

    If PrefixRS.BOF = False Then
        PrefixRS.MoveFirst
        While PrefixRS.EOF = False
            A = PrefixRS!number
            If A > 0 Then
                If PrefixRS!Name = "" Then
                    PrefixRS.Edit
                    PrefixRS.Delete
                Else
                    With ItemPrefix(A)
                        .Name = PrefixRS!Name
                        .ModificationType = PrefixRS!ModificationType
                        .ModificationValue = PrefixRS!ModificationValue
                        .OccursNaturally = PrefixRS!OccursNaturally
                        .Version = PrefixRS!Version
                    End With
                End If
            End If
            PrefixRS.MoveNext
        Wend
    End If
    GeneratePrefixVersionList

    If Startup = True Then
        frmLoading.lblStatus = "Loading Suffixes.."
        frmLoading.lblStatus.Refresh
    End If

    If SuffixRS.BOF = False Then
        SuffixRS.MoveFirst
        While SuffixRS.EOF = False
            A = SuffixRS!number
            If A > 0 Then
                If SuffixRS!Name = "" Then
                    SuffixRS.Edit
                    SuffixRS.Delete
                Else
                    With ItemSuffix(A)
                        .Name = SuffixRS!Name
                        .ModificationType = SuffixRS!ModificationType
                        .ModificationValue = SuffixRS!ModificationValue
                        .OccursNaturally = SuffixRS!OccursNaturally
                        .Version = SuffixRS!Version
                    End With
                End If
            End If
            SuffixRS.MoveNext
        Wend
    End If
    GenerateSuffixVersionList

    Exit Sub

DataRSError:
    bDataRSError = True
    Resume Next
End Sub
Sub CreateDatabase()
'Create Database
    Set DB = WS.CreateDatabase("server.dat", ";pwd=ontario" + dbLangGeneral, dbEncrypt + dbVersion30)

    CreateAccountsTable
    CreateGuildsTable
    CreateNPCsTable
    CreateMonstersTable
    CreateObjectsTable
    CreateDataTable
    CreateMapsTable
    CreateBansTable
    CreateHallsTable
    CreateScriptsTable
    CreateMagicTable
    CreatePrefixTable
    CreateSuffixTable
End Sub

Sub CreateAccountsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Accounts Table
    Set Td = DB.CreateTableDef("Accounts")

    'Create Fields
    'Account Data
    Set NewField = Td.CreateField("User", dbText, 15)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Password", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Email", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Access", dbByte)
    Td.Fields.Append NewField

    'Character Data
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Class", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Gender", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Sprite", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Desc", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField

    'Position Data
    Set NewField = Td.CreateField("Map", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("X", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Y", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("D", dbByte)
    Td.Fields.Append NewField

    'Physical Stat Data
    Set NewField = Td.CreateField("Strength", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Agility", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Endurance", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Intelligence", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Concentration", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Constitution", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Stamina", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Wisdom", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Level", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Experience", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StatPoints", dbInteger)
    Td.Fields.Append NewField

    'Misc Data
    Set NewField = Td.CreateField("Bank", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Status", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("LastPlayed", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Flags", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Skills", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Magic", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField

    'Inventory Data
    For A = 1 To 20
        Set NewField = Td.CreateField("InvObject" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("InvValue" + CStr(A), dbLong)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("InvPrefix" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("InvSuffix" + CStr(A), dbByte)
        Td.Fields.Append NewField
    Next A

    For A = 1 To 6
        Set NewField = Td.CreateField("EquippedObject" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("EquippedVal" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("EquippedPrefix" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("EquippedSuffix" + CStr(A), dbByte)
        Td.Fields.Append NewField
    Next A

    'Item Bank
    For A = 0 To 29
        Set NewField = Td.CreateField("BankObject" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("BankValue" + CStr(A), dbLong)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("BankPrefix" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("BankSuffix" + CStr(A), dbByte)
        Td.Fields.Append NewField
    Next A

    'Create Indexes
    Set NewIndex = Td.CreateIndex("User")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("User")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    Set NewIndex = Td.CreateIndex("Name")
    NewIndex.Primary = False
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Name")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Accounts Table
    DB.TableDefs.Append Td
End Sub

Sub CreateDataTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field

    'Create Accounts Table
    Set Td = DB.CreateTableDef("Data")

    'Create Fields
    Set NewField = Td.CreateField("MOTD", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MapResetTime", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjResetTime", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("BackupInterval", dbLong)
    Td.Fields.Append NewField

    Set NewField = Td.CreateField("MaxLevel", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("EliteLevel", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("DeathTime", dbByte)
    Td.Fields.Append NewField

    Set NewField = Td.CreateField("Cost_Per_Durability", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Cost_Per_Strength", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Cost_Per_Modifier", dbInteger)
    Td.Fields.Append NewField

    Set NewField = Td.CreateField("ObjMoney", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjSmallFish", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjMediumFish", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjLargeFish", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjLumber", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjLowOre", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjMedOre", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjHighOre", dbInteger)
    Td.Fields.Append NewField

    Set NewField = Td.CreateField("ServerPort", dbInteger)
    Td.Fields.Append NewField

    Set NewField = Td.CreateField("GuildJoinLevel", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("GuildJoinPrice", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("GuildNewLevel", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("GuildNewPrice", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("GuildMaxMembers", dbByte)
    Td.Fields.Append NewField
    
    Set NewField = Td.CreateField("GuildUpkeepMembers", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("GuildUpkeepSprite", dbInteger)
    Td.Fields.Append NewField

    Set NewField = Td.CreateField("StatStrength", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StatEndurance", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StatIntelligence", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StatConcentration", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StatConstitution", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StatStamina", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StatWisdom", dbByte)
    Td.Fields.Append NewField

    For A = 0 To 4
        Set NewField = Td.CreateField("StartLocationMap" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("StartLocationX" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("StartLocationY" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("StartLocationMessage" + CStr(A), dbText, 255)
        NewField.AllowZeroLength = True
        Td.Fields.Append NewField
    Next A
    For A = 1 To 8
        Set NewField = Td.CreateField("StartingObj" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("StartingObjVal" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A
    Set NewField = Td.CreateField("LastUpdate", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjectData", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Flags", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField

    'Append Data Table
    DB.TableDefs.Append Td
End Sub
Sub CreateMapsTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Maps Table
    Set Td = DB.CreateTableDef("Maps")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbInteger)
    Td.Fields.Append NewField

    Set NewField = Td.CreateField("Data", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Maps Table
    DB.TableDefs.Append Td
End Sub
Sub CreateGuildsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Guilds Table
    Set Td = DB.CreateTableDef("Guilds")
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 25)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Hall", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Sprite", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Bank", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("DueDate", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Kills", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Deaths", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("CreationDate", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MOTD", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MOTDDate", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MOTDCreator", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField

    For A = 0 To 19
        Set NewField = Td.CreateField("MemberName" + CStr(A), dbText, 15)
        NewField.AllowZeroLength = True
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("MemberRank" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("MemberJoinDate" + CStr(A), dbLong)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("MemberKills" + CStr(A), dbLong)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("MemberDeaths" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A

    For A = 0 To DeclarationCount
        Set NewField = Td.CreateField("DeclarationGuild" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("DeclarationType" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("DeclarationDate" + CStr(A), dbLong)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("DeclarationKills" + CStr(A), dbLong)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("DeclarationDeaths" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Guilds Table
    DB.TableDefs.Append Td
End Sub
Sub CreateBansTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Bans Table
    Set Td = DB.CreateTableDef("Bans")
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Banner", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ComputerID", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("IPAddress", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Reason", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("UnbanDate", dbLong)
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Bans Table
    DB.TableDefs.Append Td
End Sub
Sub CreateHallsTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Halls Table
    Set Td = DB.CreateTableDef("Halls")
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Price", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Upkeep", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StartLocationMap", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StartLocationX", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StartLocationY", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Version", dbByte)
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Bans Table
    DB.TableDefs.Append Td
End Sub

Sub CreateScriptsTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Scripts Table
    Set Td = DB.CreateTableDef("Scripts")
    Set NewField = Td.CreateField("Name", dbText, 25)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Source", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Name")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Name")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Scripts Table
    DB.TableDefs.Append Td
End Sub

Sub CreateNPCsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create NPC Table
    Set Td = DB.CreateTableDef("NPCs")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 35)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("JoinText", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("LeaveText", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    For A = 0 To 4
        Set NewField = Td.CreateField("SayText" + CStr(A), dbText, 255)
        NewField.AllowZeroLength = True
        Td.Fields.Append NewField
    Next A
    Set NewField = Td.CreateField("Flags", dbByte)
    Td.Fields.Append NewField

    For A = 0 To 9
        Set NewField = Td.CreateField("GiveObject" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("GiveValue" + CStr(A), dbLong)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("TakeObject" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("TakeValue" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A

    Set NewField = Td.CreateField("Version", dbByte)
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append NPC Table
    DB.TableDefs.Append Td
End Sub

Sub CreateMonstersTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create NPC Table
    Set Td = DB.CreateTableDef("Monsters")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 35)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Description", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Sprite", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("HP", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Strength", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Armor", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Speed", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Sight", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Agility", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Flags", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Experience", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MagicDefense", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Version", dbByte)
    Td.Fields.Append NewField

    For A = 0 To 2
        Set NewField = Td.CreateField("Object" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("Value" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Monster Table
    DB.TableDefs.Append Td
End Sub
Sub CreateObjectsTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Objects Table
    Set Td = DB.CreateTableDef("Objects")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 35)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Description", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Picture", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Type", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data1", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data2", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data3", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data4", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ClassReq", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("LevelReq", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("SellPrice", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Version", dbByte)
    Td.Fields.Append NewField

    'Flags
    Set NewField = Td.CreateField("Flags", dbByte)
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Object Table
    DB.TableDefs.Append Td
End Sub

Sub CreateMagicTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Objects Table
    Set Td = DB.CreateTableDef("Magic")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 25)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Level", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Class", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Version", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Icon", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("IconType", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("CastTimer", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Description", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Object Table
    DB.TableDefs.Append Td
End Sub

Sub CreatePrefixTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Objects Table
    Set Td = DB.CreateTableDef("Prefix")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 20)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ModificationType", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ModificationValue", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("OccursNaturally", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Version", dbByte)
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Object Table
    DB.TableDefs.Append Td
End Sub

Sub CreateSuffixTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Objects Table
    Set Td = DB.CreateTableDef("Suffix")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 20)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ModificationType", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ModificationValue", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("OccursNaturally", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Version", dbByte)
    Td.Fields.Append NewField

    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex

    'Append Object Table
    DB.TableDefs.Append Td
End Sub

Sub ConvertMap(MapNum As Long)
    Dim X As Long, Y As Long
    Dim St As String, St1 As String * 30, St2 As String

    With Map(MapNum)
        St1 = .Name
        St = St1 + QuadChar(.Version) + DoubleChar$(CLng(.NPC)) + Chr$(.Midi) + DoubleChar$(CLng(.ExitUp)) + DoubleChar$(CLng(.ExitDown)) + DoubleChar$(CLng(.ExitLeft)) + DoubleChar$(CLng(.ExitRight)) + DoubleChar(CLng(.BootLocation.Map)) + Chr$(.BootLocation.X) + Chr$(.BootLocation.Y) + DoubleChar(CLng(.DeathLocation.Map)) + Chr$(.DeathLocation.X) + Chr$(.DeathLocation.Y)
        St = St + Chr$(.flags) + Chr$(.Flags2) + DoubleChar$(CLng(.MonsterSpawn(0).Monster)) + Chr$(.MonsterSpawn(0).Rate) + DoubleChar$(CLng(.MonsterSpawn(1).Monster)) + Chr$(.MonsterSpawn(1).Rate) + DoubleChar$(CLng(.MonsterSpawn(2).Monster)) + Chr$(.MonsterSpawn(2).Rate) + DoubleChar$(CLng(.MonsterSpawn(3).Monster)) + Chr$(.MonsterSpawn(3).Rate) + DoubleChar$(CLng(.MonsterSpawn(4).Monster)) + Chr$(.MonsterSpawn(4).Rate) + DoubleChar$(CLng(.MonsterSpawn(5).Monster)) + Chr$(.MonsterSpawn(5).Rate) + DoubleChar$(CLng(.MonsterSpawn(6).Monster)) + Chr$(.MonsterSpawn(6).Rate) + DoubleChar$(CLng(.MonsterSpawn(7).Monster)) + Chr$(.MonsterSpawn(7).Rate) + DoubleChar$(CLng(.MonsterSpawn(8).Monster)) + Chr$(.MonsterSpawn(8).Rate) + DoubleChar$(CLng(.MonsterSpawn(9).Monster)) + Chr$(.MonsterSpawn(9).Rate)
        For Y = 0 To 11
            For X = 0 To 11
                With .Tile(X, Y)
                    St = St + DoubleChar(CLng(.Ground)) + DoubleChar$(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + DoubleChar(CLng(.BGTile2)) + DoubleChar(CLng(.FGTile)) + DoubleChar(CLng(.FGTile2)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(.Att2)
                End With
            Next X
        Next Y
    End With
    
    St2 = CompressString(St)
    MapRS.Seek "=", MapNum
    If MapRS.NoMatch Then
        MapRS.AddNew
        MapRS!number = MapNum
    Else
        MapRS.Edit
    End If
    MapRS!Data = St2
    MapRS.Update
    LoadMap MapNum, St
End Sub
