VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Odyssey Realms Server [Options]"
   ClientHeight    =   9105
   ClientLeft      =   345
   ClientTop       =   90
   ClientWidth     =   10635
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUpkeepSprite 
      Height          =   285
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   109
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtUpkeepMembers 
      Height          =   285
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   106
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox txtDeathTime 
      Height          =   285
      Left            =   9600
      MaxLength       =   3
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtMaxGuildMembers 
      Height          =   285
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   75
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox txtWisdom 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   103
      Top             =   8640
      Width           =   975
   End
   Begin VB.TextBox txtStamina 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   101
      Top             =   8280
      Width           =   975
   End
   Begin VB.TextBox txtConstitution 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   99
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox txtConcentration 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   97
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtIntelligence 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   93
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtEndurance 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   91
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtStrength 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   87
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtMaxLevel 
      Height          =   285
      Left            =   9600
      MaxLength       =   3
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtCostPerModifier 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   83
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtCostPerStrength 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   79
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtCostPerDurability 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   73
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtJoinGuildCost 
      Height          =   285
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   69
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtNewGuildCost 
      Height          =   285
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   63
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtJoinGuildLevel 
      Height          =   285
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   55
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtNewGuildLevel 
      Height          =   285
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   50
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txtServerPort 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   44
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtLumber 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   38
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox txtHighOre 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   32
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtMediumOre 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   26
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtLowOre 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   24
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtLargeFish 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtMediumFish 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtSmallFish 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtMoneyObject 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9600
      MaxLength       =   3
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtObjectResetTime 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   58
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtBackUpInterval 
      Height          =   285
      Left            =   9600
      MaxLength       =   2
      TabIndex        =   64
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtMapResetTime 
      Height          =   285
      Left            =   9600
      MaxLength       =   6
      TabIndex        =   52
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   8
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   95
      Top             =   8280
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   7
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   89
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   6
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   85
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   5
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   81
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   4
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   77
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   3
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   71
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   2
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   67
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   1
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   62
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   8
      Left            =   120
      MaxLength       =   3
      TabIndex        =   94
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   7
      Left            =   120
      MaxLength       =   3
      TabIndex        =   88
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   6
      Left            =   120
      MaxLength       =   3
      TabIndex        =   84
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   5
      Left            =   120
      MaxLength       =   3
      TabIndex        =   80
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   4
      Left            =   120
      MaxLength       =   3
      TabIndex        =   76
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   3
      Left            =   120
      MaxLength       =   3
      TabIndex        =   70
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   2
      Left            =   120
      MaxLength       =   3
      TabIndex        =   66
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   1
      Left            =   120
      MaxLength       =   3
      TabIndex        =   61
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   4
      Left            =   2640
      MaxLength       =   255
      TabIndex        =   48
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   4
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   47
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   4
      Left            =   960
      MaxLength       =   2
      TabIndex        =   46
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   4
      Left            =   120
      MaxLength       =   4
      TabIndex        =   45
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   3
      Left            =   2640
      MaxLength       =   255
      TabIndex        =   42
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   3
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   41
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   3
      Left            =   960
      MaxLength       =   2
      TabIndex        =   40
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   3
      Left            =   120
      MaxLength       =   4
      TabIndex        =   39
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   2
      Left            =   2640
      MaxLength       =   255
      TabIndex        =   36
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   2
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   35
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   2
      Left            =   960
      MaxLength       =   2
      TabIndex        =   34
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   2
      Left            =   120
      MaxLength       =   4
      TabIndex        =   33
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   1
      Left            =   2640
      MaxLength       =   255
      TabIndex        =   30
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   1
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   29
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   1
      Left            =   960
      MaxLength       =   2
      TabIndex        =   28
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   1
      Left            =   120
      MaxLength       =   4
      TabIndex        =   27
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   0
      Left            =   2640
      MaxLength       =   255
      TabIndex        =   22
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   0
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   21
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   0
      Left            =   960
      MaxLength       =   2
      TabIndex        =   20
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   0
      Left            =   120
      MaxLength       =   4
      TabIndex        =   19
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtMOTD 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Width           =   6135
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   104
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Save"
      Height          =   495
      Left            =   4800
      TabIndex        =   105
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label lblUpkeepSprite 
      Caption         =   "Daily Upkeep for Sprite"
      Height          =   255
      Left            =   2640
      TabIndex        =   108
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label lblUpkeepMembers 
      Caption         =   "Daily Upkeep per Member"
      Height          =   255
      Left            =   2640
      TabIndex        =   107
      Top             =   7320
      Width           =   2535
   End
   Begin VB.Label cptDeathTime 
      Caption         =   "Death Respawn Time (seconds)"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Maximum Guild Members"
      Height          =   255
      Left            =   2640
      TabIndex        =   74
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label lblWisdom 
      Caption         =   "Base Mana (wisdom)"
      Height          =   255
      Left            =   6600
      TabIndex        =   102
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Label lblStamina 
      Caption         =   "Base Energy (stamina)"
      Height          =   255
      Left            =   6600
      TabIndex        =   100
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Label lblConstitution 
      Caption         =   "Base Health (constitution)"
      Height          =   255
      Left            =   6600
      TabIndex        =   98
      Top             =   7920
      Width           =   3015
   End
   Begin VB.Label lblConcentration 
      Caption         =   "Base Magic Damage (concentration)"
      Height          =   255
      Left            =   6600
      TabIndex        =   96
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label lblIntelligence 
      Caption         =   "Base Mana Regeneration (intelligence)"
      Height          =   255
      Left            =   6600
      TabIndex        =   92
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label lblEndurance 
      Caption         =   "Base Health Regeneration (endurance)"
      Height          =   255
      Left            =   6600
      TabIndex        =   90
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label lblStrength 
      Caption         =   "Base Physical Damage (strength)"
      Height          =   255
      Left            =   6600
      TabIndex        =   86
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblMaximumLevel 
      Caption         =   "Maximum Combat Level"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblCostPerModifier 
      Caption         =   "Repair Cost Per Modifier:"
      Height          =   255
      Left            =   6600
      TabIndex        =   82
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblCostPerStrength 
      Caption         =   "Repair Cost Per Strength:"
      Height          =   255
      Left            =   6600
      TabIndex        =   78
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblCostPerDurability 
      Caption         =   "Repair Cost Per Durability:"
      Height          =   255
      Left            =   6600
      TabIndex        =   72
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblJoinGuildCost 
      Caption         =   "Cost to Join Guild"
      Height          =   255
      Left            =   2640
      TabIndex        =   68
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label lblNewGuildCost 
      Caption         =   "Cost to Create New Guild"
      Height          =   375
      Left            =   2640
      TabIndex        =   60
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label lblJoinGuildLevel 
      Caption         =   "Required Level to Join Guild"
      Height          =   255
      Left            =   2640
      TabIndex        =   54
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label lblNewGuildLevel 
      Caption         =   "Required Level to Create Guild"
      Height          =   375
      Left            =   2640
      TabIndex        =   49
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label lblServerPort 
      Caption         =   "Server Port"
      Height          =   255
      Left            =   6600
      TabIndex        =   43
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblLumber 
      Caption         =   "Lumber Object #"
      Height          =   255
      Left            =   6600
      TabIndex        =   37
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label lblHighOre 
      Caption         =   "High Ore Object #"
      Height          =   255
      Left            =   6600
      TabIndex        =   31
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label lblMediumOre 
      Caption         =   "Medium Ore Object #"
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label lblLowOre 
      Caption         =   "Low Ore Object #"
      Height          =   255
      Left            =   6600
      TabIndex        =   23
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label lblLargeFish 
      Caption         =   "High Fish Object #"
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label lblMediumFish 
      Caption         =   "Medium Fish Object #"
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lblSmallFish 
      Caption         =   "Low Fish Object #"
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label cptMoneyObject 
      Caption         =   "Money Object #"
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label cptObjectReset 
      Caption         =   "Object Reset Time (milliseconds)"
      Height          =   255
      Left            =   6600
      TabIndex        =   59
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label cptBackupInterval 
      Caption         =   "Autosave Interval (minutes)"
      Height          =   255
      Left            =   6600
      TabIndex        =   65
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label cptMapResetTime 
      Caption         =   "Map Reset Time (milliseconds)"
      Height          =   255
      Left            =   6600
      TabIndex        =   53
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label cptAmount 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   57
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label cptObject 
      Caption         =   "Object"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label cptStartingObjects 
      Caption         =   "Starting Objects:"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label cptTxt 
      Caption         =   "Player Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label cptY 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label cptX 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label cptMap 
      Caption         =   "Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label cptStartLocation 
      Caption         =   "Start Locations:"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblMOTD 
      Caption         =   "Message of the Day"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim B As Long

    DataRS.Edit

    World.MOTD = txtMOTD
    DataRS!MOTD = txtMOTD
    World.ObjMoney = txtMoneyObject
    DataRS!ObjMoney = txtMoneyObject
    World.MapResetTime = txtMapResetTime
    DataRS!MapResetTime = txtMapResetTime
    World.ObjResetTime = txtObjectResetTime
    DataRS!ObjResetTime = txtObjectResetTime
    World.BackupInterval = txtBackUpInterval
    DataRS!BackupInterval = txtBackUpInterval

    Dim A As Long
    For A = 0 To 4
        With World.StartLocation(A)
            B = Val(txtMap(A))
            If B < 1 Then B = 1
            If B > MaxMaps Then B = MaxMaps
            .Map = B
            B = Val(txtX(A))
            If B < 0 Then B = 0
            If B > 11 Then B = 11
            .X = B
            B = Val(txtY(A))
            If B < 0 Then B = 0
            If B > 11 Then B = 11
            .Y = B
            .Message = txtText(A)

            DataRS.Fields("StartLocationX" + CStr(A)) = .X
            DataRS.Fields("StartLocationY" + CStr(A)) = .Y
            DataRS.Fields("StartLocationMap" + CStr(A)) = .Map
            DataRS.Fields("StartLocationMessage" + CStr(A)) = .Message
        End With
    Next A

    For A = 1 To 8
        If Val(txtObj(A).Text) < 0 Then txtObj(A).Text = 0
        If Val(txtObj(A).Text) > 255 Then txtObj(A).Text = 255
        DataRS.Fields("StartingObj" + CStr(A)) = Val(txtObj(A).Text)
        World.StartObjects(A) = Val(txtObj(A).Text)
    Next A

    For A = 1 To 8
        If Val(txtVal(A).Text) < 0 Then txtVal(A).Text = 0
        If Val(txtVal(A).Text) > 32000 Then txtVal(A).Text = 32000
        DataRS.Fields("StartingObjVal" + CStr(A)) = Val(txtVal(A).Text)
        World.StartObjValues(A) = Val(txtVal(A).Text)
    Next A

    World.MaxLevel = txtMaxLevel
    DataRS!MaxLevel = World.MaxLevel
    World.DeathTime = Val(txtDeathTime)
    DataRS!DeathTime = World.DeathTime

    World.ObjSmallFish = txtSmallFish
    DataRS!ObjSmallFish = World.ObjSmallFish
    World.ObjMediumFish = txtMediumFish
    DataRS!ObjMediumFish = World.ObjMediumFish
    World.ObjLargeFish = txtLargeFish
    DataRS!ObjLargeFish = World.ObjLargeFish
    World.ObjLowOre = txtLowOre
    DataRS!ObjLowOre = World.ObjLowOre
    World.ObjMedOre = txtMediumOre
    DataRS!ObjMedOre = World.ObjMedOre
    World.ObjHighOre = txtHighOre
    DataRS!ObjHighOre = World.ObjHighOre
    World.ObjLumber = txtLumber
    DataRS!ObjLumber = World.ObjLumber

    World.ServerPort = txtServerPort
    DataRS!ServerPort = World.ServerPort

    World.GuildJoinLevel = txtJoinGuildLevel
    DataRS!GuildJoinLevel = World.GuildJoinLevel
    World.GuildJoinPrice = txtJoinGuildCost
    DataRS!GuildJoinPrice = World.GuildJoinPrice
    World.GuildNewLevel = txtNewGuildLevel
    DataRS!GuildNewLevel = World.GuildNewLevel
    World.GuildNewPrice = txtNewGuildCost
    DataRS!GuildNewPrice = World.GuildNewPrice
    World.GuildMaxMembers = txtMaxGuildMembers
    DataRS!GuildMaxMembers = World.GuildMaxMembers

    World.Cost_Per_Durability = txtCostPerDurability
    DataRS!Cost_Per_Durability = World.Cost_Per_Durability
    World.Cost_Per_Strength = txtCostPerStrength
    DataRS!Cost_Per_Strength = World.Cost_Per_Strength
    World.Cost_Per_Modifier = txtCostPerModifier
    DataRS!Cost_Per_Modifier = World.Cost_Per_Modifier

    World.StatStrength = txtStrength
    DataRS!StatStrength = World.StatStrength
    World.StatEndurance = txtEndurance
    DataRS!StatEndurance = World.StatEndurance
    World.StatIntelligence = txtIntelligence
    DataRS!StatIntelligence = World.StatIntelligence
    World.StatConcentration = txtConcentration
    DataRS!StatConcentration = World.StatConcentration
    World.StatConstitution = txtConstitution
    DataRS!StatConstitution = World.StatConstitution
    World.StatStamina = txtStamina
    DataRS!StatStamina = World.StatStamina
    World.StatWisdom = txtWisdom
    DataRS!StatWisdom = World.StatWisdom
    
    World.GuildUpkeepMembers = txtUpkeepMembers
    DataRS!GuildUpkeepMembers = World.GuildUpkeepMembers
    World.GuildUpkeepSprite = txtUpkeepSprite
    DataRS!GuildUpkeepSprite = World.GuildUpkeepSprite

    DataRS.Update

    For A = 1 To MaxUsers
        If Player(A).Mode > 0 Then
            SendServerOptions A
            CalculateStats A
        End If
    Next A

    Unload Me
End Sub

Private Sub Form_Load()
    Dim A As Long

    txtMOTD = World.MOTD
    txtBackUpInterval = World.BackupInterval
    txtMapResetTime = World.MapResetTime
    txtObjectResetTime = World.ObjResetTime
    txtMoneyObject = World.ObjMoney
    txtMaxLevel = World.MaxLevel
    txtDeathTime = World.DeathTime

    txtSmallFish = World.ObjSmallFish
    txtMediumFish = World.ObjMediumFish
    txtLargeFish = World.ObjLargeFish
    txtLowOre = World.ObjLowOre
    txtMediumOre = World.ObjMedOre
    txtHighOre = World.ObjHighOre
    txtLumber = World.ObjLumber

    txtServerPort = World.ServerPort

    txtJoinGuildLevel = World.GuildJoinLevel
    txtJoinGuildCost = World.GuildJoinPrice
    txtNewGuildLevel = World.GuildNewLevel
    txtNewGuildCost = World.GuildNewPrice
    txtMaxGuildMembers = World.GuildMaxMembers

    txtStrength = World.StatStrength
    txtEndurance = World.StatEndurance
    txtIntelligence = World.StatIntelligence
    txtConcentration = World.StatConcentration
    txtConstitution = World.StatConstitution
    txtStamina = World.StatStamina
    txtWisdom = World.StatWisdom

    txtCostPerDurability = CStr(World.Cost_Per_Durability)
    txtCostPerStrength = CStr(World.Cost_Per_Strength)
    txtCostPerModifier = CStr(World.Cost_Per_Modifier)
    
    txtUpkeepMembers = World.GuildUpkeepMembers
    txtUpkeepSprite = World.GuildUpkeepSprite

    For A = 0 To 4
        With World.StartLocation(A)
            txtMap(A) = .Map
            txtX(A) = .X
            txtY(A) = .Y
            txtText(A) = .Message
        End With
    Next A

    For A = 1 To 8
        txtObj(A).Text = World.StartObjects(A)
        txtVal(A).Text = World.StartObjValues(A)
    Next A
End Sub

Private Sub txtMOTD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        KeyAscii = 0
        Beep
    End If
End Sub
