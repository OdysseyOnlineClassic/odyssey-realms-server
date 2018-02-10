Attribute VB_Name = "modClasses"
Option Explicit

Sub CreateClassData()
    If World.ServerPort > 1 Then 'Class Stats
        With Class(1)    'Knight
            .StartHP = 20
            .StartEnergy = 15
            .StartMana = 10
            .MaxHP = 200
            .MaxEnergy = 140
            .MaxMana = 50
        End With
        With Class(2)    'Mage
            .StartHP = 10
            .StartEnergy = 15
            .StartMana = 20
            .MaxHP = 120
            .MaxEnergy = 100
            .MaxMana = 200
        End With
        With Class(3)    'Thief
            .StartHP = 15
            .StartEnergy = 20
            .StartMana = 10
            .MaxHP = 140
            .MaxEnergy = 175
            .MaxMana = 100
        End With
        With Class(4)    'Cleric
            .StartHP = 10
            .StartEnergy = 15
            .StartMana = 15
            .MaxHP = 140
            .MaxEnergy = 100
            .MaxMana = 150
        End With
        Exit Sub
    End If
End Sub
