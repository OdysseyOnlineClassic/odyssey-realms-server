VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Odyssey Server Options"
   ClientHeight    =   10215
   ClientLeft      =   345
   ClientTop       =   90
   ClientWidth     =   8940
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUpkeepSprite 
      Height          =   285
      Left            =   3240
      MaxLength       =   6
      TabIndex        =   113
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox txtUpkeepMembers 
      Height          =   285
      Left            =   3240
      MaxLength       =   6
      TabIndex        =   110
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtDeathTime 
      Height          =   285
      Left            =   7680
      MaxLength       =   3
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtMaxGuildMembers 
      Height          =   285
      Left            =   3240
      MaxLength       =   6
      TabIndex        =   78
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox txtWisdom 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   107
      Top             =   9120
      Width           =   975
   End
   Begin VB.TextBox txtStamina 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   105
      Top             =   8880
      Width           =   975
   End
   Begin VB.TextBox txtConstitution 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   103
      Top             =   8640
      Width           =   975
   End
   Begin VB.TextBox txtConcentration 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   101
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox txtIntelligence 
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   97
      Top             =   9240
      Width           =   975
   End
   Begin VB.TextBox txtEndurance 
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   95
      Top             =   9000
      Width           =   975
   End
   Begin VB.TextBox txtStrength 
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   91
      Top             =   8760
      Width           =   975
   End
   Begin VB.TextBox txtMaxLevel 
      Height          =   285
      Left            =   7680
      MaxLength       =   3
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtCostPerModifier 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   87
      Top             =   5715
      Width           =   975
   End
   Begin VB.TextBox txtCostPerStrength 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   82
      Top             =   5235
      Width           =   975
   End
   Begin VB.TextBox txtCostPerDurability 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   76
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtJoinGuildCost 
      Height          =   285
      Left            =   3240
      MaxLength       =   6
      TabIndex        =   72
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtNewGuildCost 
      Height          =   285
      Left            =   3240
      MaxLength       =   6
      TabIndex        =   65
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtJoinGuildLevel 
      Height          =   285
      Left            =   3240
      MaxLength       =   6
      TabIndex        =   56
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtNewGuildLevel 
      Height          =   285
      Left            =   3240
      MaxLength       =   6
      TabIndex        =   50
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtServerPort 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   44
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtLumber 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   38
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtHighOre 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   32
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtMediumOre 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   26
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtLowOre 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   24
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtLargeFish 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   18
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMediumFish 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtSmallFish 
      Height          =   285
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMoneyObject 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7680
      MaxLength       =   3
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtObjectResetTime 
      Height          =   285
      Left            =   6840
      MaxLength       =   6
      TabIndex        =   59
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtBackUpInterval 
      Height          =   285
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   66
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtMapResetTime 
      Height          =   285
      Left            =   6840
      MaxLength       =   6
      TabIndex        =   52
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   8
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   99
      Top             =   8280
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   7
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   93
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   6
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   89
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   5
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   84
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   4
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   80
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   3
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   74
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   2
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   70
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   1
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   64
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   8
      Left            =   120
      MaxLength       =   3
      TabIndex        =   98
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   7
      Left            =   120
      MaxLength       =   3
      TabIndex        =   92
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   6
      Left            =   120
      MaxLength       =   3
      TabIndex        =   88
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   5
      Left            =   120
      MaxLength       =   3
      TabIndex        =   83
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   4
      Left            =   120
      MaxLength       =   3
      TabIndex        =   79
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   3
      Left            =   120
      MaxLength       =   3
      TabIndex        =   73
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   2
      Left            =   120
      MaxLength       =   3
      TabIndex        =   69
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   1
      Left            =   120
      MaxLength       =   3
      TabIndex        =   63
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
      Width           =   5895
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6000
      TabIndex        =   108
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Save"
      Height          =   495
      Left            =   7560
      TabIndex        =   109
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label lblUpkeepSprite 
      Alignment       =   1  'Right Justify
      Caption         =   "Upkeep for Sprite"
      Height          =   495
      Left            =   2280
      TabIndex        =   112
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label lblUpkeepMembers 
      Alignment       =   1  'Right Justify
      Caption         =   "Upkeep for Members"
      Height          =   495
      Left            =   2280
      TabIndex        =   111
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label cptDeathTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Death Time:"
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Guild Members:"
      Height          =   495
      Left            =   2280
      TabIndex        =   77
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label lblWisdom 
      Caption         =   "Wisdom:"
      Height          =   255
      Left            =   6480
      TabIndex        =   106
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label lblStamina 
      Caption         =   "Stamina"
      Height          =   255
      Left            =   6480
      TabIndex        =   104
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Label lblConstitution 
      Caption         =   "Constitution"
      Height          =   255
      Left            =   6480
      TabIndex        =   102
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label lblConcentration 
      Caption         =   "Concentration:"
      Height          =   255
      Left            =   6480
      TabIndex        =   100
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label lblIntelligence 
      Caption         =   "Mana Regen:"
      Height          =   255
      Left            =   120
      TabIndex        =   96
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label lblEndurance 
      Caption         =   "HP Regen:"
      Height          =   255
      Left            =   120
      TabIndex        =   94
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label lblStrength 
      Caption         =   "Damage:"
      Height          =   255
      Left            =   120
      TabIndex        =   90
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label lblPointsPerBonus 
      Caption         =   "Stat Points Per Bonus:"
      Height          =   255
      Left            =   6480
      TabIndex        =   85
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label lblMaximumLevel 
      Alignment       =   1  'Right Justify
      Caption         =   "Maximum Level:"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   645
      Width           =   1215
   End
   Begin VB.Label lblCostPerModifier 
      Alignment       =   1  'Right Justify
      Caption         =   "Repair Cost Per Modifier:"
      Height          =   495
      Left            =   6360
      TabIndex        =   86
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblCostPerStrength 
      Alignment       =   1  'Right Justify
      Caption         =   "Repair Cost Per Strength:"
      Height          =   495
      Left            =   6360
      TabIndex        =   81
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblCostPerDurability 
      Alignment       =   1  'Right Justify
      Caption         =   "Repair Cost Per Durability:"
      Height          =   495
      Left            =   6360
      TabIndex        =   75
      Top             =   4605
      Width           =   1215
   End
   Begin VB.Label lblJoinGuildCost 
      Alignment       =   1  'Right Justify
      Caption         =   "Join Guild Cost:"
      Height          =   495
      Left            =   2280
      TabIndex        =   71
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label lblNewGuildCost 
      Alignment       =   1  'Right Justify
      Caption         =   "New Guild Cost:"
      Height          =   495
      Left            =   2280
      TabIndex        =   62
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lblJoinGuildLevel 
      Alignment       =   1  'Right Justify
      Caption         =   "Join Guild Level:"
      Height          =   495
      Left            =   2280
      TabIndex        =   55
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblNewGuildLevel 
      Alignment       =   1  'Right Justify
      Caption         =   "New Guild Level:"
      Height          =   495
      Left            =   2280
      TabIndex        =   49
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblServerPort 
      Alignment       =   1  'Right Justify
      Caption         =   "Server Port:"
      Height          =   255
      Left            =   6720
      TabIndex        =   43
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblLumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Lumber:"
      Height          =   255
      Left            =   6720
      TabIndex        =   37
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblHighOre 
      Alignment       =   1  'Right Justify
      Caption         =   "High Ore:"
      Height          =   255
      Left            =   6600
      TabIndex        =   31
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblMediumOre 
      Alignment       =   1  'Right Justify
      Caption         =   "Medium Ore:"
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblLowOre 
      Alignment       =   1  'Right Justify
      Caption         =   "Low Ore:"
      Height          =   255
      Left            =   6600
      TabIndex        =   23
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblLargeFish 
      Alignment       =   1  'Right Justify
      Caption         =   "Large Fish:"
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblMediumFish 
      Alignment       =   1  'Right Justify
      Caption         =   "Medium Fish:"
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblSmallFish 
      Alignment       =   1  'Right Justify
      Caption         =   "Small Fish:"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label cptMs 
      Caption         =   "Milliseconds"
      Height          =   255
      Left            =   7920
      TabIndex        =   53
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label cptMilliseconds 
      Caption         =   "Milliseconds"
      Height          =   255
      Left            =   7920
      TabIndex        =   60
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label cptMinutes 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   7920
      TabIndex        =   67
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label cptMoneyObject 
      Alignment       =   1  'Right Justify
      Caption         =   "Money Object:"
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   285
      Width           =   1215
   End
   Begin VB.Label cptObjectReset 
      Alignment       =   1  'Right Justify
      Caption         =   "Obj Reset Time:"
      Height          =   255
      Left            =   5520
      TabIndex        =   61
      Top             =   6525
      Width           =   1215
   End
   Begin VB.Label cptBackupInterval 
      Alignment       =   1  'Right Justify
      Caption         =   "Backup Interval:"
      Height          =   255
      Left            =   5520
      TabIndex        =   68
      Top             =   6885
      Width           =   1215
   End
   Begin VB.Label cptMapResetTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Map Reset Time:"
      Height          =   255
      Left            =   5520
      TabIndex        =   54
      Top             =   6165
      Width           =   1215
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
      TabIndex        =   58
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
      Left            =   240
      TabIndex        =   57
      Top             =   5400
      Width           =   615
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
      Caption         =   "Text"
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
      Width           =   735
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
      Caption         =   "MOTD:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
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
