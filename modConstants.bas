Attribute VB_Name = "modConstants"
Option Explicit

'Misc
Public Const LagHitDistance = 2
Public Const NumClasses = 4

Public Const DeclarationCount = 4

'Maximum Constants
Public Const MaxGuilds = 255
Public Const MaxMaps = 3000
Public Const MaxMapObjects = 79
Public Const MaxPlayerFlags = 1000
Public Const MaxNPCs = 500
Public Const MaxMagic = 500
Public Const MaxHalls = 255
Public Const MaxModifications = 255
Public Const MaxObjects = 1000
Public Const MaxTotalMonsters = 1000
Public Const MaxMonsters = 19
Public Const MaxPlayerTimers = 20
Public Const MaxSkill = 10
Public Const MaxSprite = 643

'Projectile Types
Public Const pttCharacter = 0
Public Const pttPlayer = 1
Public Const pttMonster = 2
Public Const pttTile = 3
Public Const pttProject = 4

'Stats
Public Const statPlayerAgility = 30

'Hooking
Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

'Connection Modes
Public Const modeNotConnected = 0
Public Const modeConnected = 1
Public Const modePlaying = 2
Public Const modeBanned = 3
