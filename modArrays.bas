Attribute VB_Name = "modArrays"
Type ScriptData
    Name As String
    Source As String
    MCode() As Byte
End Type

Type SkillData
    Level As Byte
    Experience As Long
End Type

Type MapStartLocationData
    Map As Integer
    X As Byte
    Y As Byte
End Type

Type BugData
    PlayerName As String
    PlayerUser As String
    PlayerIP As String
    Title As String
    Description As String
    Status As Byte
    ResolverName As String
    ResolverUser As String
End Type

Type HallData
    Name As String
    Price As Long
    Upkeep As Long
    StartLocation As MapStartLocationData
    Version As Byte
End Type

Type GuildDeclarationData
    Guild As Byte
    Type As Byte
    Date As Long
    Kills As Long
    Deaths As Long
End Type

Type GuildMemberData
    Name As String
    Rank As Byte
    JoinDate As Long
    Kills As Long
    Deaths As Long
End Type

Type GuildData
    Name As String
    Member(0 To 19) As GuildMemberData
    MemberCount As Byte
    Declaration(0 To DeclarationCount) As GuildDeclarationData
    Hall As Byte
    Bank As Long
    Sprite As Integer
    DueDate As Long
    CreationDate As Long
    Kills As Long
    Deaths As Long
    Bookmark As Variant
    MOTD As String
    MOTDDate As Long
    MOTDCreator As String
    UpdateFlag As Boolean
End Type

Type InvObject
    Object As Integer
    Value As Long
    ItemPrefix As Byte
    ItemSuffix As Byte
End Type

Type BanData
    Name As String
    Reason As String
    UnbanDate As Long
    Banner As String
    InUse As Boolean
    ComputerID As String
    IPAddress As String
End Type

Type ItemBankData
    Object As Integer
    Value As Long
    ItemPrefix As Byte
    ItemSuffix As Byte
End Type

Type EquippedObjectData
    Object As Integer
    Value As Long
    ItemPrefix As Byte
    ItemSuffix As Byte
End Type

Type ProjectileDamageData
    Live As Boolean
    Damage As Byte
    ShootTime As Currency
End Type

Type PlayerData
    'Socket Data
    Socket As Long
    SocketData As String
    IP As String
    ClientVer As String
    InUse As Boolean
    Mode As Byte
    LastMsg As Currency

    TotalDefense As Byte
    MagicDefense As Byte
    PhysicalAttack As Byte

    'Account Data
    User As String
    Access As Byte
    ComputerID As String
    Email As String

    'Character Data
    Name As String
    Class As Byte
    Gender As Byte
    Sprite As Integer
    desc As String

    'Position Data
    Map As Integer
    X As Byte
    Y As Byte
    D As Byte

    'Vital Stat Data
    MaxHP As Integer
    MaxEnergy As Integer
    MaxMana As Integer
    HP As Integer
    Energy As Integer
    Mana As Integer

    PacketOrder As Integer
    ServerPacketOrder As Integer

    'Physical Stat Data
    Level As Byte
    Experience As Long

    'Misc. Data
    Status As Integer
    Bank As Long
    TimeLeft As Currency

    ScriptTimer(1 To MaxPlayerTimers) As Long
    Script(1 To MaxPlayerTimers) As String

    ItemBank(0 To 29) As ItemBankData
    Skill(1 To 10) As SkillData
    MagicLevel(1 To MaxMagic) As SkillData

    'Guild Data
    Guild As Byte
    GuildRank As Byte
    GuildSlot As Byte
    JoinRequest As Byte

    'Inventory Data
    Inv(1 To 20) As InvObject
    EquippedObject(1 To 6) As EquippedObjectData
    ProjectileDamage(1 To 20) As ProjectileDamageData

    'Flag Data
    Flag(0 To MaxPlayerFlags) As Long

    FloodTimer As Currency
    CastTimer As Currency

    WalkTimer As Currency
    WalkCount As Currency
    ShootTimer As Currency
    AttackTimer As Currency
    SpeedHackTimer As Currency

    IsDead As Boolean
    DeadTick As Currency
    SpeedTick As Currency
    LastSkillUse As Currency

    'Target Data
    CurrentRepairTar As Integer

    SpeedStrikes As Long

    'Database Data
    Bookmark As Variant
End Type

Type ClassData
    MaxHP As Integer
    MaxEnergy As Integer
    MaxMana As Integer
    StartHP As Integer
    StartEnergy As Integer
    StartMana As Integer
End Type

Type ObjectData
    Name As String
    Picture As Integer
    Type As Byte
    Data(0 To 3) As Byte
    flags As Byte
    ClassReq As Byte
    LevelReq As Byte
    Version As Byte
    SellPrice As Long
End Type

Type MonsterData
    Name As String
    Description As String
    Sprite As Integer
    flags As Byte
    HP As Integer
    Strength As Byte
    Armor As Byte
    Speed As Byte
    Sight As Byte
    Agility As Byte
    Object(0 To 2) As Long
    Value(0 To 2) As Long
    Experience As Integer
    MagicDefense As Byte
    Version As Byte
End Type

Type NPCSaleItemData
    GiveObject As Long
    GiveValue As Long
    TakeObject As Long
    TakeValue As Long
End Type

Type MagicData
    Name As String
    Level As Byte
    Class As Byte
    Version As Byte
    Icon As Integer
    IconType As Byte
    CastTimer As Integer
    Description As String
End Type

Type NPCData
    Name As String
    JoinText As String
    LeaveText As String
    SayText(0 To 4) As String
    SaleItem(0 To 9) As NPCSaleItemData
    flags As Byte
    Version As Byte
End Type

Type TileData
    Ground As Integer
    Ground2 As Integer
    BGTile1 As Integer
    BGTile2 As Integer
    FGTile As Integer
    FGTile2 As Integer
    Att As Byte
    AttData(0 To 3) As Byte
    Att2 As Byte
End Type

Type MapDoorData
    Att As Byte
    X As Byte
    Y As Byte
    T As Long
End Type

Type MapObjectData
    Object As Long
    Value As Long
    TimeStamp As Long
    X As Byte
    Y As Byte
    ItemPrefix As Byte
    ItemSuffix As Byte
End Type

Type MapMonsterSpawnData
    Monster As Integer
    Rate As Byte
    Timer As Currency
End Type

Type MapMonsterData
    Monster As Integer
    X As Byte
    Y As Byte
    D As Byte
    Target As Byte
    TargetIsMonster As Boolean
    HP As Integer
    AttackTimer As Currency
    MoveTimer As Currency
    Frame As Byte
End Type

Type MapData
    Name As String
    ExitUp As Integer
    ExitDown As Integer
    ExitLeft As Integer
    ExitRight As Integer
    Tile(0 To 11, 0 To 11) As TileData
    Object(0 To MaxMapObjects) As MapObjectData
    Monster(0 To MaxMonsters) As MapMonsterData
    MonsterSpawn(0 To 9) As MapMonsterSpawnData
    Door(0 To 9) As MapDoorData
    BootLocation As MapStartLocationData
    DeathLocation As MapStartLocationData
    flags As Byte
    Flags2 As Byte
    NumPlayers As Long
    ResetTimer As Currency
    Hall As Byte
    NPC As Integer
    Midi As Byte
    Keep As Boolean
    Version As Long
    CheckSum As Long
    LastUpdate As Currency
End Type

Type StartLocationData
    X As Byte
    Y As Byte
    Map As Integer
    Message As String
End Type

Type WorldData
    LastUpdate As Currency
    MapResetTime As Currency
    ObjResetTime As Currency
    BackupInterval As Long
    StartLocation(0 To 4) As StartLocationData
    MOTD As String
    Flag(0 To 255) As Long
    StartObjects(1 To 8) As Integer
    StartObjValues(1 To 8) As Long
    
    GuildUpkeepMembers As Integer
    GuildUpkeepSprite As Integer
    
    EliteLevel As Byte
    MaxLevel As Byte
    DeathTime As Byte

    'Blacksmithy Costs
    Cost_Per_Durability As Integer
    Cost_Per_Strength As Integer
    Cost_Per_Modifier As Integer

    'Special Objects
    ObjMoney As Integer
    ObjSmallFish As Integer
    ObjMediumFish As Integer
    ObjLargeFish As Integer
    ObjLumber As Integer
    ObjLowOre As Integer
    ObjMedOre As Integer
    ObjHighOre As Integer

    'Ports
    ServerPort As Integer

    'Guilds
    GuildJoinPrice As Long
    GuildNewPrice As Long
    GuildJoinLevel As Byte
    GuildNewLevel As Byte
    GuildMaxMembers As Byte

    'Stats
    StatStrength As Byte
    StatEndurance As Byte
    StatIntelligence As Byte
    StatConcentration As Byte
    StatConstitution As Byte
    StatStamina As Byte
    StatWisdom As Byte
End Type

Type PrefixData
    Name As String
    ModificationType As Byte
    ModificationValue As Byte
    OccursNaturally As Byte
    Version As Byte
End Type

Public CloseSocketQue(1 To MaxUsers) As Long
Public World As WorldData
Public Map(1 To MaxMaps) As MapData
Public Guild(1 To MaxGuilds) As GuildData
Public Magic(1 To MaxMagic) As MagicData
Public Hall(1 To MaxHalls) As HallData
Public Object(1 To MaxObjects) As ObjectData
Public Monster(1 To MaxTotalMonsters) As MonsterData
Public NPC(1 To MaxNPCs) As NPCData
Public Player(1 To MaxUsers + 1) As PlayerData
Public Class(1 To 5) As ClassData
Public Ban(1 To 50) As BanData
Public ItemPrefix(1 To MaxModifications) As PrefixData
Public ItemSuffix(1 To MaxModifications) As PrefixData
Public Bug(1 To 500) As BugData
Public NumUsers As Long
