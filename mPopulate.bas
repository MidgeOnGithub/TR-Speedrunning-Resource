<<<<<<< HEAD
Attribute VB_Name = "mPopulate"
Option Explicit

Function PopulateColl(Data As String) As Collection
    Set PopulateColl = New Collection 'Return collection is initialized here, items/keys added near end
    
    Dim arrData As Variant
    Select Case Data  'Arrays Hardcoded for TR2.
        Case "Enemy":  'Indexed roughly with respect to appearance in-game.
            arrData = Array("Tiger", "Crow", "Spider", _
                "T-Rex", "Rat", "Dog", "Melee Thug", _
                "Pistol Thug", "Colt", "Shotgun Thug", _
                "Auto Rifle Thug", "Flamethrower", "Scuba Diver", _
                "Shark", "Barracuda", "Uzi Thug", _
                "Burst Pistol Thug", "Snowmobile Thug", "Snow Leopard", _
                "Monastery Monk", "Catfish", "Yeti", "Talion Guardian", _
                "Eagle", "Giant Spider", "Statue", "Spear Statue", _
                "Sword Statue", "Bartoli's Monk", "Dragon")
        Case "Level":  'For TR3, indices of this array change with runner's level order.
            arrData = Array("Great Wall", "Venice", "Bartoli's Hideout", _
                "Opera House", "Offshore Rig", "Diving Area", "40 Fathoms", _
                "Wreck of the Maria Doria", "Living Quarters", _
                "The Deck", "Tibetan Foothills", "Barkhang Monastery", _
                "Catacombs of the Talion", "Ice Palace", "Temple of Xian", _
                "Floating Islands", "Dragon's Lair", "NG+")
        Case "Weapon":  'Indexed with respect to PC hotkeys.
            arrData = Array("Pistols", "Shotgun", "Automatic Pistols", _
                "Uzis", "Harpoon Gun", "M16", "Grenade Launcher") 
        Case Else
            MsgBox "Code error: incorrect argument for PopulateColl. Terminating."
            End
    End Select
    
    'Loop to add items into the collection.----------
    Dim ItemName As Variant
    Dim KeyName As String
    Dim i As Integer: i = 1
    For Each ItemName In arrData
        KeyName = Data & i  'Outputs Data string followed by number, keys used later to call Item name and to make it obvious which index number FindCollIndex will output.
        PopulateColl.Add ItemName, KeyName
        i = i + 1
    Next '-------------------------------------------

    Set arrData = Nothing  'Redundant with VBA End Function; clears memory.
End Function
=======
Attribute VB_Name = "mPopulate"
Option Explicit

Function PopulateColl(Data As String) As Collection
    Set PopulateColl = New Collection 'Return collection is initialized here, items/keys added near end
    
    Dim arrData As Variant
    Select Case Data  'Arrays Hardcoded for TR2.
        Case "Enemy":  'Indexed roughly with respect to appearance in-game.
            arrData = Array("Tiger", "Crow", "Spider", _
                "T-Rex", "Rat", "Dog", "Melee Thug", _
                "Pistol Thug", "Colt", "Shotgun Thug", _
                "Auto Rifle Thug", "Flamethrower", "Scuba Diver", _
                "Shark", "Barracuda", "Uzi Thug", _
                "Burst Pistol Thug", "Snowmobile Thug", "Snow Leopard", _
                "Monastery Monk", "Catfish", "Yeti", "Talion Guardian", _
                "Eagle", "Giant Spider", "Statue", "Spear Statue", _
                "Sword Statue", "Bartoli's Monk", "Dragon")
        Case "Level":  'For TR3, indices of this array change with runner's level order.
            arrData = Array("Great Wall", "Venice", "Bartoli's Hideout", _
                "Opera House", "Offshore Rig", "Diving Area", "40 Fathoms", _
                "Wreck of the Maria Doria", "Living Quarters", _
                "The Deck", "Tibetan Foothills", "Barkhang Monastery", _
                "Catacombs of the Talion", "Ice Palace", "Temple of Xian", _
                "Floating Islands", "Dragon's Lair", "NG+")
        Case "Weapon":  'Indexed with respect to PC hotkeys.
            arrData = Array("Pistols", "Shotgun", "Automatic Pistols", _
                "Uzis", "Harpoon Gun", "M16", "Grenade Launcher")
        Case Else
            MsgBox "Code error: incorrect argument for PopulateColl. Terminating."
            End
    End Select
    
    'Loop to add items into the collection.----------
    Dim ItemName As Variant
    Dim KeyName As String
    Dim i As Integer: i = 1
    For Each ItemName In arrData
        KeyName = Data & i  'Outputs Data string followed by number, keys used later to call Item name and to make it obvious which index number FindCollIndex will output.
        PopulateColl.Add ItemName, KeyName
        i = i + 1
    Next '-------------------------------------------

    Set arrData = Nothing  'Redundant with VBA End Function; clears memory.
End Function
>>>>>>> d7b9fa5... mWeaponsAvailable fixed, extra spaces removed
