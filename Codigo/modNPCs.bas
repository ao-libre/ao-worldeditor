Attribute VB_Name = "modNPCs"
    Dim Obj As String
    Dim datos As String
    Dim Name As String
    Dim NpcType As String
    Dim Desc As String
    Dim Head As String
    Dim Heading As String
    Dim Body As String
    Dim Movement As String
    Dim AguaValida As String
    Dim TierraInvalida As String
    Dim Attackable As String
    Dim Faccion As String
    Dim AtacaDoble As String
    Dim ReSpawn As String
    Dim Hostile As String
    Dim Domable As String
    Dim Alineacion As String
    Dim Comercia As String
    Dim GiveEXP As String
    Dim GiveGLD As String
    Dim MinHP As String
    Dim MaxHP As String
    Dim MaxHIT As String
    Dim MinHIT As String
    Dim DEF As String
    Dim DEFm As String
    Dim AfectaParalisis As String
    Dim Veneno As String
    Dim PoderAtaque As String
    Dim PoderEvasion As String
    Dim Snd1 As String
    Dim Snd2 As String
    Dim Snd3 As String
    Dim NROITEMS As String
    Dim Drop1 As String
    Dim Drop2 As String
    Dim Drop3 As String
    Dim Drop4 As String
    Dim Drop5 As String
        
    Dim Obj1 As String
    Dim Obj2 As String
    Dim Obj3 As String
    Dim Obj4 As String
    Dim Obj5 As String
    Dim Obj6 As String
    Dim Obj7 As String
    Dim Obj8 As String
    Dim Obj9 As String
    Dim Obj10 As String
    Dim Obj11 As String
    Dim Obj12 As String
    Dim Obj13 As String
    Dim Obj14 As String
    Dim Obj15 As String
    Dim Obj16 As String
    Dim Obj17 As String
    Dim Obj18 As String
    Dim Obj19 As String
    Dim Obj20 As String
    Dim Obj21 As String
    Dim Obj22 As String
    Dim Obj23 As String
    Dim Obj24 As String
    Dim Obj25 As String
    
    Dim NROEXP As String
    Dim Exp1 As String
    Dim Exp2 As String
    Dim Exp3 As String
    Dim Exp4 As String
    Dim Exp5 As String
    
    Dim NroCriaturas As String
    Dim CI1 As String
    Dim CI2 As String
    Dim CI3 As String
    Dim CI4 As String
    Dim CI5 As String
    Dim CI6 As String
    Dim CI7 As String
    Dim CI8 As String
    Dim CI9 As String
    Dim CI10 As String
    Dim CI11 As String
    Dim CI12 As String
    Dim CI13 As String
    Dim CI14 As String
    Dim CI15 As String
    Dim CN1 As String
    Dim CN2 As String
    Dim CN3 As String
    Dim CN4 As String
    Dim CN5 As String
    Dim CN6 As String
    Dim CN7 As String
    Dim CN8 As String
    Dim CN9 As String
    Dim CN10 As String
    Dim CN11 As String
    Dim CN12 As String
    Dim CN13 As String
    Dim CN14 As String
    Dim CN15 As String
    
    Dim LanzaSpells As String
    Dim Sp1 As String
    Dim Sp2 As String
    Dim Sp3 As String
    
    Dim Ciudad As String
    
    Dim BackUp As String
    Dim OrigPos As String
    
    Dim TipoItems As String
    Dim InvReSpawn As String
    Dim QuestNumber As String
    
    Dim NumNPCs As String
    
    
      
    Dim ultimaFila As String
    Dim ultimaFilaAuxiliar As String
    Dim Cont As Long
    Dim palabraBusqueda As String
    
Sub CrearMasterNPCs()
'**************************************************************
'* Leemos los datos de NPCs.dat en un exel y los pasamos a una hoja master
'* creado por ReyarB
'* ultima modificacion 30/05/2020
'***************************************************************
    Call modcopiarObj.CrearDirectorio
    Call modcopiarObj.CrearHojasObjyNPCs
    
    Call CrearHojaNpcs("NPCsIndex")
    
    Range("A5:DF5").HorizontalAlignment = xlCenter
    Range("A5:DF5").VerticalAlignment = xlCenter
    Range("A5:DF5").RowHeight = 25
    Range("A5:DF5").Borders.Weight = XlBorderWeight.xlThick
    Range("A2").ColumnWidth = 20
    Range("B2").ColumnWidth = 40
    Range("C8:D2").ColumnWidth = 8
    Range("E2").ColumnWidth = 10
            
    palabraBusqueda = "*" & "NPC" & "*"
    ultimaFila = Sheets("NPCs").Range("B" & Rows.count).End(xlUp).Row
    
    If ultimaFila < 6 Then
        Exit Sub
    End If
    
    For Cont = 6 To ultimaFila
        For Y = 1 To 111
            
            If Sheets("NPCs").Cells(Cont, 1) = "" Then Exit For
            If Sheets("NPCs").Cells(Cont, 1) = "NumNPCs" Then NumNPCs = Sheets("NPCs").Cells(Cont, 2) & ("'") & Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) Like palabraBusqueda Then Obj = Sheets("NPCs").Cells(Cont, 1)
            If Sheets("NPCs").Cells(Cont, 1) = "Name" Then Name = Sheets("NPCs").Cells(Cont, 2) & Sheets("NPCs").Cells(Cont, 3) & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "NpcType" Then NpcType = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Desc" Then Desc = Sheets("NPCs").Cells(Cont, 2) & Sheets("NPCs").Cells(Cont, 3) & Sheets("NPCs").Cells(Cont, 4) & Sheets("NPCs").Cells(Cont, 5) & Sheets("NPCs").Cells(Cont, 6) & Sheets("NPCs").Cells(Cont, 7) & Sheets("NPCs").Cells(Cont, 8) & Sheets("NPCs").Cells(Cont, 9) & Sheets("NPCs").Cells(Cont, 10)
            If Sheets("NPCs").Cells(Cont, 1) = "Head" Then Head = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Heading" Then Heading = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Body" Then Body = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Movement" Then Movement = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "AguaValida" Then AguaValida = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "TierraInvalida" Then TierraInvalida = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Attackable" Then Attackable = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "ReSpawn" Then ReSpawn = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Hostile" Then Hostile = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Domable" Then Domable = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Alineacion" Then Alineacion = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Comercia" Then Comercia = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "GiveEXP" Then GiveEXP = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "GiveGLD" Then GiveGLD = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "MinHP" Then MinHP = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "MaxHP" Then MaxHP = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "MaxHIT" Then MaxHIT = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "MinHIT" Then MinHIT = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "DEF" Then DEF = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "DEFm" Then DEFm = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "AfectaParalisis" Then AfectaParalisis = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Veneno" Then Veneno = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "PoderAtaque" Then PoderAtaque = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "PoderEvasion" Then PoderEvasion = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "SND1" Then Snd1 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "SND2" Then Snd2 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "SND3" Then Snd3 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "NROITEMS" Then NROITEMS = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Drop1" Then Drop1 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Drop2" Then Drop2 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Drop3" Then Drop3 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Drop4" Then Drop4 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Drop5" Then Drop5 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj1" Then Obj1 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj2" Then Obj2 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj3" Then Obj3 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj4" Then Obj4 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj5" Then Obj5 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj6" Then Obj6 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj7" Then Obj7 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj8" Then Obj8 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj9" Then Obj9 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj10" Then Obj10 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj11" Then Obj11 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj12" Then Obj12 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj13" Then Obj13 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj14" Then Obj14 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj15" Then Obj15 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj16" Then Obj16 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj17" Then Obj17 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj18" Then Obj18 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj19" Then Obj19 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj20" Then Obj20 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj21" Then Obj21 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj22" Then Obj22 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj23" Then Obj23 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj24" Then Obj24 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            If Sheets("NPCs").Cells(Cont, 1) = "Obj25" Then Obj25 = Sheets("NPCs").Cells(Cont, 2) & ("-") & Sheets("NPCs").Cells(Cont, 3) & ("'") & Sheets("NPCs").Cells(Cont, 4)
            
            If Sheets("NPCs").Cells(Cont, 1) = "NROEXP" Then NROEXP = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Exp1" Then Exp1 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Exp2" Then Exp2 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Exp3" Then Exp3 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Exp4" Then Exp4 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Exp5" Then Exp5 = Sheets("NPCs").Cells(Cont, 2)
                
            If Sheets("NPCs").Cells(Cont, 1) = "NroCriaturas" Then NroCriaturas = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI1" Then CI1 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI2" Then CI2 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI3" Then CI3 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI4" Then CI4 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI5" Then CI5 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI6" Then CI6 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI7" Then CI7 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI8" Then CI8 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI9" Then CI9 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI10" Then CI10 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI11" Then CI11 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI12" Then CI12 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI13" Then CI13 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI14" Then CI14 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CI15" Then CI15 = Sheets("NPCs").Cells(Cont, 2)
            
            If Sheets("NPCs").Cells(Cont, 1) = "CN1" Then CN1 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN2" Then CN2 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN3" Then CN3 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN4" Then CN4 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN5" Then CN5 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN6" Then CN6 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN7" Then CN7 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN8" Then CN8 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN9" Then CN9 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN10" Then CN10 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN11" Then CN11 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN12" Then CN12 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN13" Then CN13 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN14" Then CN14 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "CN15" Then CN15 = Sheets("NPCs").Cells(Cont, 2)

            If Sheets("NPCs").Cells(Cont, 1) = "LanzaSpells" Then LanzaSpells = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Sp1" Then Sp1 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Sp2" Then Sp2 = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "Sp3" Then Sp3 = Sheets("NPCs").Cells(Cont, 2)
            
            If Sheets("NPCs").Cells(Cont, 1) = "Ciudad" Then BackUp = Sheets("NPCs").Cells(Cont, 2)
  
            If Sheets("NPCs").Cells(Cont, 1) = "BackUp" Then BackUp = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "OrigPos" Then OrigPos = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "TipoItems" Then TipoItems = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "InvReSpawn" Then InvReSpawn = Sheets("NPCs").Cells(Cont, 2)
            If Sheets("NPCs").Cells(Cont, 1) = "QuestNumber" Then QuestNumber = Sheets("NPCs").Cells(Cont, 2)
                       

             Cont = Cont + 1
            
         Next Y

        ultimaFilaAuxiliar = Sheets("NPCsIndex").Range("A" & Rows.count).End(xlUp).Row
    

        
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 1) = Obj
    If NumNPCs <> "" Then
            NumNPCs = Sheets("NPCs").Cells(Cont - 1, 2)
            Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 2) = NumNPCs
        Else
            NumNPCs = Sheets("NPCs").Cells(Cont, 2)
            Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 2) = Name
    End If
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 3) = NpcType
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 4) = Desc
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 5) = Head
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 6) = Heading
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 7) = Body
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 8) = Movement
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 9) = AguaValida
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 10) = TierraInvalida
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 11) = Attackable
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 12) = Faccion
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 13) = AtacaDoble
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 14) = ReSpawn
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 15) = Hostile
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 16) = Domable
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 17) = Alineacion
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 18) = Comercia
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 19) = GiveEXP
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 20) = GiveGLD
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 21) = MinHP
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 22) = MaxHP
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 23) = MaxHIT
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 24) = MinHIT
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 25) = DEF
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 26) = DEFm
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 27) = AfectaParalisis
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 28) = Veneno
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 29) = PoderAtaque
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 30) = PoderEvasion
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 31) = Snd1
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 32) = Snd2
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 33) = Snd3
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 34) = NROITEMS
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 35) = Drop1
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 36) = Drop2
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 37) = Drop3
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 38) = Drop4
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 39) = Drop5
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 40) = Obj1
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 41) = Obj2
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 42) = Obj3
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 43) = Obj4
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 44) = Obj5
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 45) = Obj6
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 46) = Obj7
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 47) = Obj8
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 48) = Obj9
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 49) = Obj10
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 50) = Obj11
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 51) = Obj12
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 52) = Obj13
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 53) = Obj14
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 54) = Obj15
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 55) = Obj16
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 56) = Obj17
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 57) = Obj18
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 58) = Obj19
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 59) = Obj20
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 60) = Obj21
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 61) = Obj22
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 62) = Obj23
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 63) = Obj24
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 64) = Obj25
                
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 65) = NROEXP
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 66) = Exp1
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 67) = Exp2
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 68) = Exp3
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 69) = Exp4
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 70) = Exp5
                
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 71) = NroCriaturas
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 72) = CI1
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 73) = CI2
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 74) = CI3
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 75) = CI4
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 76) = CI5
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 77) = CI6
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 78) = CI7
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 79) = CI8
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 80) = CI9
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 81) = CI10
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 82) = CI11
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 83) = CI12
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 84) = CI13
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 85) = CI14
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 86) = CI15
        
        
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 87) = CN1
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 88) = CN2
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 89) = CN3
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 90) = CN4
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 91) = CN5
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 92) = CN6
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 93) = CN7
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 94) = CN8
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 95) = CN9
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 96) = CN10
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 97) = CN11
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 98) = CN12
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 99) = CN13
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 100) = CN14
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 101) = CN15

        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 102) = LanzaSpells
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 103) = Sp1
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 104) = Sp2
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 105) = Sp3
       
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 106) = BackUp
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 107) = OrigPos
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 108) = TipoItems
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 109) = InvReSpawn
        Sheets("NPCsIndex").Cells(ultimaFilaAuxiliar + 1, 110) = QuestNumber
        
        
                     
        Obj = ""
        Name = ""
        NpcType = ""
        Desc = ""
        Head = ""
        Heading = ""
        Body = ""
        Movement = ""
        AguaValida = ""
        TierraInvalida = ""
        Attackable = ""
        Faccion = ""
        AtacaDoble = ""
        ReSpawn = ""
        Hostile = ""
        Domable = ""
        Alineacion = ""
        Comercia = ""
        GiveEXP = ""
        GiveGLD = ""
        MinHP = ""
        MaxHP = ""
        MaxHIT = ""
        MinHIT = ""
        DEF = ""
        DEFm = ""
        AfectaParalisis = ""
        Veneno = ""
        PoderAtaque = ""
        PoderEvasion = ""
        Snd1 = ""
        Snd2 = ""
        Snd3 = ""
        NROITEMS = ""
        Drop1 = ""
        Drop2 = ""
        Drop3 = ""
        Drop4 = ""
        Drop5 = ""
        Obj1 = ""
        Obj2 = ""
        Obj3 = ""
        Obj4 = ""
        Obj5 = ""
        Obj6 = ""
        Obj7 = ""
        Obj8 = ""
        Obj9 = ""
        Obj10 = ""
        Obj11 = ""
        Obj12 = ""
        Obj13 = ""
        Obj14 = ""
        Obj15 = ""
        Obj16 = ""
        Obj17 = ""
        Obj18 = ""
        Obj19 = ""
        Obj20 = ""
        Obj21 = ""
        Obj22 = ""
        Obj23 = ""
        Obj24 = ""
        Obj25 = ""
        NROEXP = ""
        Exp1 = ""
        Exp2 = ""
        Exp3 = ""
        Exp4 = ""
        Exp5 = ""
        NroCriaturas = ""
        CI1 = ""
        CI2 = ""
        CI3 = ""
        CI4 = ""
        CI5 = ""
        CI6 = ""
        CI7 = ""
        CI8 = ""
        CI9 = ""
        CI10 = ""
        CI11 = ""
        CI12 = ""
        CI13 = ""
        CI14 = ""
        CI15 = ""
        
        CN1 = ""
        CN2 = ""
        CN3 = ""
        CN4 = ""
        CN5 = ""
        CN6 = ""
        CN7 = ""
        CN8 = ""
        CN9 = ""
        CN10 = ""
        CN11 = ""
        CN12 = ""
        CN13 = ""
        CN14 = ""
        CN15 = ""
        CN5 = ""
        LanzaSpells = ""
        Sp1 = ""
        Sp2 = ""
        Sp3 = ""
        Ciudad = ""
        BackUp = ""
        OrigPos = ""
        
        TipoItems = ""
        InvReSpawn = ""
        QuestNumber = ""
        NumNPCs = ""
        
    Next Cont
    
        ultimaFilaAuxiliar = Sheets("NPCsIndex").Range("A" & Rows.count).End(xlUp).Row
        With Sheets("NPCsIndex").Range("A6:DF" & ultimaFilaAuxiliar).Font
        .Name = "arial"
        .Size = 9
        .italic = True
        
        MsgBox "Proceso Terminado. Master de NPCsIndex", vbInformation, "Resultado"
    End With

End Sub

Function CrearHojaNpcs(NameHoja As String) As Boolean
    
    Dim existe As Boolean
    On Error Resume Next
    If NameHoja = (Worksheets(NameHoja).Name) Then
        Sheets(Array(NameHoja)).Delete
    End If
    existe = (Worksheets(NameHoja).Name <> "")
    If Not existe Then
        Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = NameHoja
        ultimaFilaAuxiliar = 4
        
        If NameHoja = "NPCsIndex" Or NameHoja = "Balance" Then
            Sheets(NameHoja).Cells(5, 1) = "Objeto Nº"
            Sheets(NameHoja).Cells(5, 2) = "Name"
            Sheets(NameHoja).Cells(5, 3) = "NpcType"
            Sheets(NameHoja).Cells(5, 4) = "Desc"
            Sheets(NameHoja).Cells(5, 5) = "Head"
            Sheets(NameHoja).Cells(5, 6) = "Heading"
            Sheets(NameHoja).Cells(5, 7) = "Body"
            Sheets(NameHoja).Cells(5, 8) = "Movement"
            Sheets(NameHoja).Cells(5, 9) = "AguaValida"
            Sheets(NameHoja).Cells(5, 10) = "TierraInvalida"
            Sheets(NameHoja).Cells(5, 11) = "Attackable"
            Sheets(NameHoja).Cells(5, 12) = "Faccion"
            Sheets(NameHoja).Cells(5, 13) = "AtacaDoble"
            Sheets(NameHoja).Cells(5, 14) = "ReSpawn"
            Sheets(NameHoja).Cells(5, 15) = "Hostile"
            Sheets(NameHoja).Cells(5, 16) = "Domable"
            Sheets(NameHoja).Cells(5, 17) = "Alineacion"
            Sheets(NameHoja).Cells(5, 18) = "Comercia"
            Sheets(NameHoja).Cells(5, 19) = "GiveEXP"
            Sheets(NameHoja).Cells(5, 20) = "GiveGLD"
            Sheets(NameHoja).Cells(5, 21) = "MinHP"
            Sheets(NameHoja).Cells(5, 22) = "MaxHP"
            Sheets(NameHoja).Cells(5, 23) = "MaxHIT"
            Sheets(NameHoja).Cells(5, 24) = "MinHIT"
            Sheets(NameHoja).Cells(5, 25) = "DEF"
            Sheets(NameHoja).Cells(5, 26) = "DEFm"
            Sheets(NameHoja).Cells(5, 27) = "AfectaParalisis"
            Sheets(NameHoja).Cells(5, 28) = "Veneno"
            Sheets(NameHoja).Cells(5, 29) = "PoderAtaque"
            Sheets(NameHoja).Cells(5, 30) = "PoderEvasion"
            Sheets(NameHoja).Cells(5, 31) = "SND1"
            Sheets(NameHoja).Cells(5, 32) = "SND2"
            Sheets(NameHoja).Cells(5, 33) = "SND3"
            Sheets(NameHoja).Cells(5, 34) = "NROITEMS"
            Sheets(NameHoja).Cells(5, 35) = "Drop1"
            Sheets(NameHoja).Cells(5, 36) = "Drop2"
            Sheets(NameHoja).Cells(5, 37) = "Drop3"
            Sheets(NameHoja).Cells(5, 38) = "Drop4"
            Sheets(NameHoja).Cells(5, 39) = "Drop5"
            Sheets(NameHoja).Cells(5, 40) = "Obj1"
            Sheets(NameHoja).Cells(5, 41) = "Obj2"
            Sheets(NameHoja).Cells(5, 42) = "Obj3"
            Sheets(NameHoja).Cells(5, 43) = "Obj4"
            Sheets(NameHoja).Cells(5, 44) = "Obj5"
            Sheets(NameHoja).Cells(5, 45) = "Obj6"
            Sheets(NameHoja).Cells(5, 46) = "Obj7"
            Sheets(NameHoja).Cells(5, 47) = "Obj8"
            Sheets(NameHoja).Cells(5, 48) = "Obj9"
            Sheets(NameHoja).Cells(5, 49) = "Obj10"
            Sheets(NameHoja).Cells(5, 50) = "Obj11"
            Sheets(NameHoja).Cells(5, 51) = "Obj12"
            Sheets(NameHoja).Cells(5, 52) = "Obj13"
            Sheets(NameHoja).Cells(5, 53) = "Obj14"
            Sheets(NameHoja).Cells(5, 54) = "Obj15"
            Sheets(NameHoja).Cells(5, 55) = "Obj16"
            Sheets(NameHoja).Cells(5, 56) = "Obj17"
            Sheets(NameHoja).Cells(5, 57) = "Obj18"
            Sheets(NameHoja).Cells(5, 58) = "Obj19"
            Sheets(NameHoja).Cells(5, 59) = "Obj20"
            Sheets(NameHoja).Cells(5, 60) = "Obj21"
            Sheets(NameHoja).Cells(5, 61) = "Obj22"
            Sheets(NameHoja).Cells(5, 62) = "Obj23"
            Sheets(NameHoja).Cells(5, 63) = "Obj24"
            Sheets(NameHoja).Cells(5, 64) = "Obj25"
            
            Sheets(NameHoja).Cells(5, 65) = "NROEXP"
            Sheets(NameHoja).Cells(5, 66) = "Exp1"
            Sheets(NameHoja).Cells(5, 67) = "Exp2"
            Sheets(NameHoja).Cells(5, 68) = "Exp3"
            Sheets(NameHoja).Cells(5, 69) = "Exp4"
            Sheets(NameHoja).Cells(5, 70) = "Exp5"
            
            
            Sheets(NameHoja).Cells(5, 71) = "NroCriaturas"
            Sheets(NameHoja).Cells(5, 72) = "CI1"
            Sheets(NameHoja).Cells(5, 73) = "CI2"
            Sheets(NameHoja).Cells(5, 74) = "CI3"
            Sheets(NameHoja).Cells(5, 75) = "CI4"
            Sheets(NameHoja).Cells(5, 76) = "CI5"
            Sheets(NameHoja).Cells(5, 77) = "CI6"
            Sheets(NameHoja).Cells(5, 78) = "CI7"
            Sheets(NameHoja).Cells(5, 79) = "CI8"
            Sheets(NameHoja).Cells(5, 80) = "CI9"
            Sheets(NameHoja).Cells(5, 81) = "CI10"
            Sheets(NameHoja).Cells(5, 82) = "CI11"
            Sheets(NameHoja).Cells(5, 83) = "CI12"
            Sheets(NameHoja).Cells(5, 84) = "CI13"
            Sheets(NameHoja).Cells(5, 85) = "CI14"
            Sheets(NameHoja).Cells(5, 86) = "CI15"
            Sheets(NameHoja).Cells(5, 87) = "CN1"
            Sheets(NameHoja).Cells(5, 88) = "CN2"
            Sheets(NameHoja).Cells(5, 89) = "CN3"
            Sheets(NameHoja).Cells(5, 90) = "CN4"
            Sheets(NameHoja).Cells(5, 91) = "CN5"
            Sheets(NameHoja).Cells(5, 92) = "CN6"
            Sheets(NameHoja).Cells(5, 93) = "CN7"
            Sheets(NameHoja).Cells(5, 94) = "CN8"
            Sheets(NameHoja).Cells(5, 95) = "CN9"
            Sheets(NameHoja).Cells(5, 96) = "CN10"
            Sheets(NameHoja).Cells(5, 97) = "CN11"
            Sheets(NameHoja).Cells(5, 98) = "CN12"
            Sheets(NameHoja).Cells(5, 99) = "CN13"
            Sheets(NameHoja).Cells(5, 100) = "CN14"
            Sheets(NameHoja).Cells(5, 101) = "CN15"
                     
            Sheets(NameHoja).Cells(5, 102) = "LanzaSpells"
            Sheets(NameHoja).Cells(5, 103) = "Sp1"
            Sheets(NameHoja).Cells(5, 104) = "Sp2"
            Sheets(NameHoja).Cells(5, 105) = "Sp3"
            
            Sheets(NameHoja).Cells(5, 106) = "BackUp"
            Sheets(NameHoja).Cells(5, 107) = "OrigPos"
            Sheets(NameHoja).Cells(5, 108) = "TipoItems"
            Sheets(NameHoja).Cells(5, 109) = "InvReSpawn"
            Sheets(NameHoja).Cells(5, 110) = "QuestNumber"
  
            Sheets(NameHoja).Range("A5:DF5").Interior.color = RGB(153, 204, 0)

        End If
    End If
' le damos formato a las celdas del titulo
 If NameHoja = "NPCsIndex" Or NameHoja = "Balance" Then
            Sheets(NameHoja).Range("A5:DF5").Interior.color = RGB(153, 204, 0)
            Range("A2").ColumnWidth = 20
            Range("B2").ColumnWidth = 40
            Range("C8:D2").ColumnWidth = 8
            Range("E2").ColumnWidth = 10
            Range("F2:I2").ColumnWidth = 8
            Range("J2").ColumnWidth = 10
            Range("K2").ColumnWidth = 8
            Range("L2:V2").ColumnWidth = 12
            Range("W2:Y2").ColumnWidth = 8
            Range("Z2:AB2").ColumnWidth = 12
            Range("AC2").ColumnWidth = 15
            Range("AD").ColumnWidth = 5
            Range("AE2:AF2").ColumnWidth = 9
            Range("AG2").ColumnWidth = 40
            Range("AH2").ColumnWidth = 10
            Range("AI2").ColumnWidth = 6
            Range("AJ2").ColumnWidth = 15
            Range("AK2").ColumnWidth = 10
            Range("AL2").ColumnWidth = 16
            Range("AM2:AR2").ColumnWidth = 10
            Range("AS2:AT2").ColumnWidth = 12
            Range("AU2").ColumnWidth = 6
            Range("AV2:BA2").ColumnWidth = 11
            Range("BB2:BD2").ColumnWidth = 14
            Range("BE2").ColumnWidth = 9
            Range("A5:BJ5").HorizontalAlignment = xlCenter
            Range("A5:BJ5").VerticalAlignment = xlCenter
            Range("A5:BJ5").RowHeight = 25
            Range("A5:BJ5").Borders.Weight = XlBorderWeight.xlThick
 End If
    CrearHojaNpcs = existe
     
End Function

'**************************************************************
'* Leemos los datos de objtos.dat en un exel y los pasamos a una hoja master
'* creado por ReyarB
'* ultima modificacion 30/05/2020
'***************************************************************

' Modulo para consultas

Sub CrearConsultaMasterNPCs()

    Call modcopiarObj.CrearDirectorio
    Call modcopiarObj.CrearHojasObjyNPCs
    Call modcopiarObj.AbrirHojasNPCsIndex
    
    Call CrearHojaNpcs("Balance")
    
    Range("A5:BE5").HorizontalAlignment = xlCenter
    Range("A5:BE5").VerticalAlignment = xlCenter
    Range("A5:BE5").RowHeight = 25
    Range("A5:BE5").Borders.Weight = XlBorderWeight.xlMedium
    Range("A5:BE5").Borders.color = RGB(235, 0, 0)
        
    Range("A5:BE5").Font.Name = "arial"
    Range("A5:BE5").Font.Size = 9
    Range("A5:BE5").Font.bold = True
    Range("A5:BE5").Font.italic = True
        
    palabraBusqueda = UserForm1.txtNPCType.Text
    
    
    'palabraBusqueda = "*" & palabraBusqueda & "*"
    ultimaFila = Sheets("NPCsIndex").Range("B" & Rows.count).End(xlUp).Row
    
    If ultimaFila < 6 Then
        Exit Sub
    End If
    
    For Cont = 6 To ultimaFila
          
            
    If Sheets("NPCsIndex").Cells(Cont, 3) Like palabraBusqueda Then
        Obj = Sheets("NPCsIndex").Cells(Cont, 1)
        Name = Sheets("NPCsIndex").Cells(Cont, 2)
        NpcType = Sheets("NPCsIndex").Cells(Cont, 3)
        Desc = Sheets("NPCsIndex").Cells(Cont, 4)
        Head = Sheets("NPCsIndex").Cells(Cont, 5)
        Heading = Sheets("NPCsIndex").Cells(Cont, 6)
        Body = Sheets("NPCsIndex").Cells(Cont, 7)
        Movement = Sheets("NPCsIndex").Cells(Cont, 8)
        AguaValida = Sheets("NPCsIndex").Cells(Cont, 9)
        Attackable = Sheets("NPCsIndex").Cells(Cont, 10)
        ReSpawn = Sheets("NPCsIndex").Cells(Cont, 11)
        Hostile = Sheets("NPCsIndex").Cells(Cont, 12)
        Domable = Sheets("NPCsIndex").Cells(Cont, 13)
        Alineacion = Sheets("NPCsIndex").Cells(Cont, 14)
        Comercia = Sheets("NPCsIndex").Cells(Cont, 15)
        GiveEXP = Sheets("NPCsIndex").Cells(Cont, 16)
        GiveGLD = Sheets("NPCsIndex").Cells(Cont, 17)
        MinHP = Sheets("NPCsIndex").Cells(Cont, 18)
        MaxHP = Sheets("NPCsIndex").Cells(Cont, 19)
        MaxHIT = Sheets("NPCsIndex").Cells(Cont, 20)
        MinHIT = Sheets("NPCsIndex").Cells(Cont, 21)
        DEF = Sheets("NPCsIndex").Cells(Cont, 22)
        AfectaParalisis = Sheets("NPCsIndex").Cells(Cont, 23)
        PoderAtaque = Sheets("NPCsIndex").Cells(Cont, 24)
        PoderEvasion = Sheets("NPCsIndex").Cells(Cont, 25)
        Snd1 = Sheets("NPCsIndex").Cells(Cont, 26)
        Snd2 = Sheets("NPCsIndex").Cells(Cont, 27)
        Snd3 = Sheets("NPCsIndex").Cells(Cont, 28)
        Drop1 = Sheets("NPCsIndex").Cells(Cont, 29)
        Drop2 = Sheets("NPCsIndex").Cells(Cont, 30)
        Drop3 = Sheets("NPCsIndex").Cells(Cont, 31)
        Drop4 = Sheets("NPCsIndex").Cells(Cont, 32)
        Drop5 = Sheets("NPCsIndex").Cells(Cont, 33)
        Drop6 = Sheets("NPCsIndex").Cells(Cont, 34)
        NROITEMS = Sheets("NPCsIndex").Cells(Cont, 35)
        Obj1 = Sheets("NPCsIndex").Cells(Cont, 36)
        Obj2 = Sheets("NPCsIndex").Cells(Cont, 37)
        Obj3 = Sheets("NPCsIndex").Cells(Cont, 38)
        Obj4 = Sheets("NPCsIndex").Cells(Cont, 39)
        Obj5 = Sheets("NPCsIndex").Cells(Cont, 40)
        Obj6 = Sheets("NPCsIndex").Cells(Cont, 41)
        Obj7 = Sheets("NPCsIndex").Cells(Cont, 42)
        Obj8 = Sheets("NPCsIndex").Cells(Cont, 43)
        Obj9 = Sheets("NPCsIndex").Cells(Cont, 44)
        Obj10 = Sheets("NPCsIndex").Cells(Cont, 45)
        Obj11 = Sheets("NPCsIndex").Cells(Cont, 46)
        Obj12 = Sheets("NPCsIndex").Cells(Cont, 47)
        Obj13 = Sheets("NPCsIndex").Cells(Cont, 48)
        Obj14 = Sheets("NPCsIndex").Cells(Cont, 49)
        Obj15 = Sheets("NPCsIndex").Cells(Cont, 50)
        Obj16 = Sheets("NPCsIndex").Cells(Cont, 51)
        Obj17 = Sheets("NPCsIndex").Cells(Cont, 52)
        Obj18 = Sheets("NPCsIndex").Cells(Cont, 53)
        Obj19 = Sheets("NPCsIndex").Cells(Cont, 54)
        Obj20 = Sheets("NPCsIndex").Cells(Cont, 55)
        Obj21 = Sheets("NPCsIndex").Cells(Cont, 56)
        Obj22 = Sheets("NPCsIndex").Cells(Cont, 57)
        Obj23 = Sheets("NPCsIndex").Cells(Cont, 58)
        Obj24 = Sheets("NPCsIndex").Cells(Cont, 59)
        Obj25 = Sheets("NPCsIndex").Cells(Cont, 60)
        
        NroCriaturas = Sheets("NPCsIndex").Cells(Cont, 61)
        CI1 = Sheets("NPCsIndex").Cells(Cont, 62)
        CI2 = Sheets("NPCsIndex").Cells(Cont, 63)
        CI3 = Sheets("NPCsIndex").Cells(Cont, 64)
        CI4 = Sheets("NPCsIndex").Cells(Cont, 65)
        CI5 = Sheets("NPCsIndex").Cells(Cont, 66)
        CN1 = Sheets("NPCsIndex").Cells(Cont, 67)
        CN2 = Sheets("NPCsIndex").Cells(Cont, 68)
        CN3 = Sheets("NPCsIndex").Cells(Cont, 69)
        CN4 = Sheets("NPCsIndex").Cells(Cont, 70)
        CN5 = Sheets("NPCsIndex").Cells(Cont, 71)

        LanzaSpells = Sheets("NPCsIndex").Cells(Cont, 72)
        Sp1 = Sheets("NPCsIndex").Cells(Cont, 73)
        Sp2 = Sheets("NPCsIndex").Cells(Cont, 74)
        Sp3 = Sheets("NPCsIndex").Cells(Cont, 75)
        BackUp = Sheets("NPCsIndex").Cells(Cont, 76)
        OrigPos = Sheets("NPCsIndex").Cells(Cont, 77)

        TipoItems = Sheets("NPCsIndex").Cells(Cont, 78)
        InvReSpawn = Sheets("NPCsIndex").Cells(Cont, 79)
        QuestNumber = Sheets("NPCsIndex").Cells(Cont, 80)
       
        
        
        

        ultimaFilaAuxiliar = Sheets("Balance").Range("A" & Rows.count).End(xlUp).Row
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 1) = Obj
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 2) = Name
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 3) = NpcType
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 4) = Desc
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 5) = Head
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 6) = Heading
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 7) = Body
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 8) = Movement
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 9) = AguaValida
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 10) = Attackable
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 11) = ReSpawn
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 12) = Hostile
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 13) = Domable
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 14) = Alineacion
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 15) = Comercia
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 16) = GiveEXP
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 17) = GiveGLD
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 18) = MinHP
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 19) = MaxHP
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 20) = MaxHIT
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 21) = MinHIT
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 22) = DEF
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 23) = AfectaParalisis
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 24) = PoderAtaque
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 25) = PoderEvasion
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 26) = Snd1
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 27) = Snd2
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 28) = Snd3
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 29) = Drop1
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 30) = Drop2
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 31) = Drop3
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 32) = Drop4
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 33) = Drop5
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 34) = Drop6
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 35) = NROITEMS
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 36) = Obj1
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 37) = Obj2
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 38) = Obj3
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 39) = Obj4
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 40) = Obj5
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 41) = Obj6
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 42) = Obj7
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 43) = Obj8
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 44) = Obj9
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 45) = Obj10
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 46) = Obj11
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 47) = Obj12
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 48) = Obj13
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 49) = Obj14
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 50) = Obj15
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 51) = Obj16
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 52) = Obj17
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 53) = Obj18
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 54) = Obj19
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 55) = Obj20
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 56) = Obj21
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 57) = Obj22
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 58) = Obj23
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 59) = Obj24
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 60) = Obj25
            
            
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 61) = NroCriaturas
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 62) = CI1
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 63) = CI2
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 64) = CI3
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 65) = CI4
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 66) = CI5
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 67) = CN1
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 68) = CN2
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 69) = CN3
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 70) = CN4
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 71) = CN5

            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 72) = LanzaSpells
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 73) = Sp1
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 74) = Sp2
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 75) = Sp3
            
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 76) = BackUp
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 77) = OrigPos

            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 78) = TipoItems
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 79) = InvReSpawn
            Sheets("Balance").Cells(ultimaFilaAuxiliar + 1, 80) = QuestNumber
            
            Sheets("Balance").Range("A5:BE5").Interior.color = RGB(153, 204, 0)
        End If

     Next Cont
    
        ultimaFilaAuxiliar = Sheets("Balance").Range("A" & Rows.count).End(xlUp).Row
        With Sheets("Balance").Range("A6:BE" & ultimaFilaAuxiliar).Font
        .Name = "arial"
        .Size = 9
        .italic = True
        
        MsgBox "Proceso Terminado", vbInformation, "Resultado"
    End With

End Sub


'**************************************************************
'* Leemos los datos de objtos.dat en un exel y los pasamos a una hoja master
'* creado por ReyarB
'* ultima modificacion 30/05/2020
'***************************************************************

' Modulo para consultas

Sub CrearNPCsBalance()

    Call modcopiarObj.CrearDirectorio
    Call modcopiarObj.CrearHojasObjyNPCs
    Call modcopiarObj.AbrirHojasNPCsIndex
    
    Call CrearHojaNpcs("NPCBalance")
    
    Range("A5:W5").HorizontalAlignment = xlCenter
    Range("A5:W5").VerticalAlignment = xlCenter
    Range("A5:W5").RowHeight = 25
    Range("A5:W5").Borders.Weight = XlBorderWeight.xlMedium
    Range("A5:W5").Borders.color = RGB(235, 0, 0)
        
    Range("A5:W5").Font.Name = "arial"
    Range("A5:W5").Font.Size = 9
    Range("A5:W5").Font.bold = True
    Range("A5:W5").Font.italic = True
    Range("A5").Select
    
            Sheets("NPCBalance").Cells(5, 1) = "Objeto Nº"
            Sheets("NPCBalance").Cells(5, 2) = "Name"
            Sheets("NPCBalance").Cells(5, 3) = "NpcType"
            Sheets("NPCBalance").Cells(5, 4) = "Desc"
            Sheets("NPCBalance").Cells(5, 5) = "Attackable"
            Sheets("NPCBalance").Cells(5, 6) = "ReSpawn"
            Sheets("NPCBalance").Cells(5, 7) = "Hostile"
            Sheets("NPCBalance").Cells(5, 8) = "Domable"
            Sheets("NPCBalance").Cells(5, 9) = "Alineacion"
            Sheets("NPCBalance").Cells(5, 10) = "GiveEXP"
            Sheets("NPCBalance").Cells(5, 11) = "GiveGLD"
            Sheets("NPCBalance").Cells(5, 12) = "MinHP"
            Sheets("NPCBalance").Cells(5, 13) = "MaxHP"
            Sheets("NPCBalance").Cells(5, 14) = "MaxHIT"
            Sheets("NPCBalance").Cells(5, 15) = "MinHIT"
            Sheets("NPCBalance").Cells(5, 16) = "DEF"
            Sheets("NPCBalance").Cells(5, 17) = "AfectaParalisis"
            Sheets("NPCBalance").Cells(5, 18) = "PoderAtaque"
            Sheets("NPCBalance").Cells(5, 19) = "PoderEvasion"
            Sheets("NPCBalance").Cells(5, 20) = "SND1"
            Sheets("NPCBalance").Cells(5, 21) = "SND2"
            Sheets("NPCBalance").Cells(5, 22) = "SND3"
            Sheets("NPCBalance").Cells(5, 23) = "BackUp"
            Sheets("NPCBalance").Cells(5, 24) = "OrigPos"
            
        
    palabraBusqueda = "1" 'UserForm1.txtNPCType.Text
    
    
    'palabraBusqueda = "*" & palabraBusqueda & "*"
    ultimaFila = Sheets("NPCsIndex").Range("B" & Rows.count).End(xlUp).Row
    
    If ultimaFila < 6 Then
        Exit Sub
    End If
    
    For Cont = 6 To ultimaFila
          
            
    'If Sheets("NPCsIndex").Cells(Cont, 11) Like palabraBusqueda Then
        Obj = Sheets("NPCsIndex").Cells(Cont, 1)
        Name = Sheets("NPCsIndex").Cells(Cont, 2)
        NpcType = Sheets("NPCsIndex").Cells(Cont, 3)
        Desc = Sheets("NPCsIndex").Cells(Cont, 4)
        Attackable = Sheets("NPCsIndex").Cells(Cont, 9)
        ReSpawn = Sheets("NPCsIndex").Cells(Cont, 10)
        Hostile = Sheets("NPCsIndex").Cells(Cont, 11)
        Domable = Sheets("NPCsIndex").Cells(Cont, 12)
        Alineacion = Sheets("NPCsIndex").Cells(Cont, 13)
        GiveEXP = Sheets("NPCsIndex").Cells(Cont, 15)
        GiveGLD = Sheets("NPCsIndex").Cells(Cont, 16)
        MinHP = Sheets("NPCsIndex").Cells(Cont, 17)
        MaxHP = Sheets("NPCsIndex").Cells(Cont, 18)
        MaxHIT = Sheets("NPCsIndex").Cells(Cont, 19)
        MinHIT = Sheets("NPCsIndex").Cells(Cont, 20)
        DEF = Sheets("NPCsIndex").Cells(Cont, 21)
        AfectaParalisis = Sheets("NPCsIndex").Cells(Cont, 22)
        PoderAtaque = Sheets("NPCsIndex").Cells(Cont, 23)
        PoderEvasion = Sheets("NPCsIndex").Cells(Cont, 24)
        Snd1 = Sheets("NPCsIndex").Cells(Cont, 25)
        Snd2 = Sheets("NPCsIndex").Cells(Cont, 26)
        Snd3 = Sheets("NPCsIndex").Cells(Cont, 27)
        BackUp = Sheets("NPCsIndex").Cells(Cont, 59)
        OrigPos = Sheets("NPCsIndex").Cells(Cont, 60)
        
        
        
        Obj = Sheets("NPCsIndex").Cells(Cont, 1)
        Name = Sheets("NPCsIndex").Cells(Cont, 2)
        NpcType = Sheets("NPCsIndex").Cells(Cont, 3)
        Desc = Sheets("NPCsIndex").Cells(Cont, 4)
        Attackable = Sheets("NPCsIndex").Cells(Cont, 10)
        ReSpawn = Sheets("NPCsIndex").Cells(Cont, 11)
        Hostile = Sheets("NPCsIndex").Cells(Cont, 12)
        Domable = Sheets("NPCsIndex").Cells(Cont, 13)
        Alineacion = Sheets("NPCsIndex").Cells(Cont, 14)
        GiveEXP = Sheets("NPCsIndex").Cells(Cont, 16)
        GiveGLD = Sheets("NPCsIndex").Cells(Cont, 17)
        MinHP = Sheets("NPCsIndex").Cells(Cont, 18)
        MaxHP = Sheets("NPCsIndex").Cells(Cont, 19)
        MaxHIT = Sheets("NPCsIndex").Cells(Cont, 20)
        MinHIT = Sheets("NPCsIndex").Cells(Cont, 21)
        DEF = Sheets("NPCsIndex").Cells(Cont, 22)
        AfectaParalisis = Sheets("NPCsIndex").Cells(Cont, 23)
        PoderAtaque = Sheets("NPCsIndex").Cells(Cont, 24)
        PoderEvasion = Sheets("NPCsIndex").Cells(Cont, 25)
        Snd1 = Sheets("NPCsIndex").Cells(Cont, 26)
        Snd2 = Sheets("NPCsIndex").Cells(Cont, 27)
        Snd3 = Sheets("NPCsIndex").Cells(Cont, 28)
        BackUp = Sheets("NPCsIndex").Cells(Cont, 75)
        OrigPos = Sheets("NPCsIndex").Cells(Cont, 76)

        ultimaFilaAuxiliar = Sheets("NPCBalance").Range("A" & Rows.count).End(xlUp).Row
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 1) = Obj
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 2) = Name
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 3) = NpcType
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 4) = Desc
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 5) = Attackable
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 6) = ReSpawn
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 7) = Hostile
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 8) = Domable
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 9) = Alineacion
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 10) = GiveEXP
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 11) = GiveGLD
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 12) = MinHP
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 13) = MaxHP
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 14) = MaxHIT
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 15) = MinHIT
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 16) = DEF
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 17) = AfectaParalisis
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 18) = PoderAtaque
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 19) = PoderEvasion
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 20) = Snd1
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 21) = Snd2
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 22) = Snd3
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 23) = BackUp
            Sheets("NPCBalance").Cells(ultimaFilaAuxiliar + 1, 24) = OrigPos

            Sheets("NPCBalance").Range("A5:W5").Interior.color = RGB(153, 204, 0)
    'End If

    Next Cont

        ultimaFilaAuxiliar = Sheets("NPCBalance").Range("A" & Rows.count).End(xlUp).Row
        With Sheets("NPCBalance").Range("A6:W" & ultimaFilaAuxiliar).Font
        .Name = "arial"
        .Size = 9
        .italic = True
        
    Columns("A:W").Select
    Columns("A:W").EntireColumn.AutoFit
    Range("A5").Select
    ActiveSheet.Range("$A$5:$W$255").AutoFilter Field:=1
    Columns("A:W").EntireColumn.AutoFit
    Range("C6").Select
    ActiveWindow.FreezePanes = True

    End With

End Sub




