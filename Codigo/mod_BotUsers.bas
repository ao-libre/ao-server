Attribute VB_Name = "mod_BotUsers"
Option Explicit

Private Enum eCabezas
    CASPER_HEAD = 500
    FRAGATA_FANTASMAL = 87
    
    HUMANO_H_PRIMER_CABEZA = 1
    HUMANO_H_ULTIMA_CABEZA = 40 'En verdad es hasta la 51, pero como son muchas estas las dejamos no seleccionables
    HUMANO_H_CUERPO_DESNUDO = 21
    
    ELFO_H_PRIMER_CABEZA = 101
    ELFO_H_ULTIMA_CABEZA = 122
    ELFO_H_CUERPO_DESNUDO = 210
    
    DROW_H_PRIMER_CABEZA = 201
    DROW_H_ULTIMA_CABEZA = 221
    DROW_H_CUERPO_DESNUDO = 32
    
    ENANO_H_PRIMER_CABEZA = 301
    ENANO_H_ULTIMA_CABEZA = 319
    ENANO_H_CUERPO_DESNUDO = 53
    
    GNOMO_H_PRIMER_CABEZA = 401
    GNOMO_H_ULTIMA_CABEZA = 416
    GNOMO_H_CUERPO_DESNUDO = 222
    
    HUMANO_M_PRIMER_CABEZA = 70
    HUMANO_M_ULTIMA_CABEZA = 89
    HUMANO_M_CUERPO_DESNUDO = 39
    
    ELFO_M_PRIMER_CABEZA = 170
    ELFO_M_ULTIMA_CABEZA = 188
    ELFO_M_CUERPO_DESNUDO = 259
    
    DROW_M_PRIMER_CABEZA = 270
    DROW_M_ULTIMA_CABEZA = 288
    DROW_M_CUERPO_DESNUDO = 40
    
    ENANO_M_PRIMER_CABEZA = 370
    ENANO_M_ULTIMA_CABEZA = 384
    ENANO_M_CUERPO_DESNUDO = 60
    
    GNOMO_M_PRIMER_CABEZA = 470
    GNOMO_M_ULTIMA_CABEZA = 484
    GNOMO_M_CUERPO_DESNUDO = 260
End Enum

Public Enum eBotAccion
    quieto = 0
    RevivirYequiparse = 1 ' AQUI INCLUYE, COMPRAR HECHIZOS, ITEMS, COMIDA Y BEBIDA, POCIONES, ETC.
    Viajando = 2
    Agite = 3
    Entrenando = 4
    Trabajando = 5
End Enum

Private Type tBots
    nick As String
    creado As Boolean
    online As Boolean
    accion As eBotAccion
    antAccion As eBotAccion
    npcIndex As Integer
    curMapViaje As Byte
    curRutaViaje As Byte
    targetBot As Byte
    targetUser As Integer
    lastaccion As eBotAccion
    minutosaccion As Byte
    minutosOnline As Byte
End Type

' a los usuarios se muestran como users pero desde el server se manejan como NPCs
'Se guardan los datos en sus charfiles especificos al deslogear
'al logear se carga inventario, skills, hechizos, oro desde el charfile

Public Const MAX_BOTS As Byte = 100

Public MinutosLogeados As Byte
Public NumBots As Byte
Public MinBots As Byte, maxBots As Byte
Public BotsOnline As Byte
Public user_Bot(1 To MAX_BOTS) As tBots

Public Sub LoadBots()
    Dim i As Long
    Dim file As String, Leer As clsIniManager
    file = DatPath & "bots.dat"
    Set Leer = New clsIniManager
    Call Leer.Initialize(file)
    
    NumBots = val(Leer.GetValue("BOTS", "NumBots"))
    
    MinBots = val(Leer.GetValue("BOTS", "MinBotsOnline"))
    maxBots = val(Leer.GetValue("BOTS", "MaxBotsOnline"))
    
    For i = 1 To MAX_BOTS
        user_Bot(i).nick = Leer.GetValue("BOTS", "Bot" & i & "NAME")
        user_Bot(i).creado = FileExist(CharPath & UCase$(user_Bot(i).nick) & ".chr")
    Next i
    
    Set Leer = Nothing
End Sub

Public Sub timer_minuto_bots()
    Static minutoslogear As Byte
    Static minutosDeslogear As Byte
    Dim porcentajelogeado As Byte, frecConn As Byte, numLogs As Byte
    Dim i As Long
    
    If RandomNumber(1, 5) = 5 Then
        minutoslogear = minutoslogear + 1
        minutosDeslogear = minutosDeslogear + 1
    End If
    
    '************CONEXION DE BOTS**************
    
    porcentajelogeado = BotsOnline / MinBots * 100
    
    'Si hay pocos bots, se logean de a mas bots, y con mayor frecuencia
    'Si esta cerca del minimo de bots se logean con menos frecuencia y de a 1
    If porcentajelogeado < 20 Then
        frecConn = 1
        numLogs = 3
    ElseIf porcentajelogeado > 20 And porcentajelogeado < 40 Then
        frecConn = 1
        numLogs = 2
    ElseIf porcentajelogeado > 40 And porcentajelogeado < 70 Then
        frecConn = 2
        numLogs = 1
    ElseIf porcentajelogeado > 70 Then
        frecConn = 5
        numLogs = 1
    ElseIf porcentajelogeado > 90 Then
        frecConn = 8
        numLogs = 1
    End If
    
    If minutoslogear >= frecConn Then
        For i = 1 To numLogs
            Call connectBot
        Next i
        minutoslogear = 0
    End If
    '*************FIN CONEXION DE BOTS********************
    
    '************DESCONEXION DE BOTS*******************
    
    Dim frecDisc As Byte, numDisc As Byte
    Dim porcentMax As Byte
    
    
    porcentMax = BotsOnline / maxBots * 100
    
    If porcentMax > 80 Then
        frecDisc = 3
        numDisc = 1
    ElseIf porcentMax <= 80 And porcentMax > 50 Then
        frecDisc = 5
        numDisc = 1
    ElseIf porcentMax <= 50 And porcentMax > 20 Then
        frecDisc = 8
        numDisc = 1
    ElseIf porcentMax < 20 Then
        frecDisc = 15
        numDisc = 1
    End If
    
    If minutosDeslogear >= frecDisc Then
        For i = 1 To numDisc
            Call disconnectbot
        Next i
    End If
    '****** *********Fin DESCONEXION DE BOTS*************
End Sub


Public Sub crear_Bot(ByVal index As Byte)
    Dim file As String, Leer As clsIniManager
    
    Dim nIndex As Integer
    Set Leer = New clsIniManager

    With user_Bot(index)
        .creado = True
        nIndex = NextOpenNPC
        
        .npcIndex = nIndex
        With Npclist(nIndex)
            .BotData.Clase = RandomNumber(1, NUMCLASES)
            
            .BotData.Lvl = 1
            
            If .BotData.Clase = eClass.Mage Then
                 If RandomNumber(1, 2) = 1 Then
                    .BotData.raza = eRaza.Elfo
                Else
                    .BotData.raza = eRaza.Gnomo
                End If
                
            ElseIf .BotData.Clase = eClass.Bard Or .BotData.Clase = eClass.Cleric Or .BotData.Clase = eClass.Druid _
                    Or .BotData.Clase = eClass.Assasin Or .BotData.Clase = eClass.Paladin Then
                .BotData.raza = RandomNumber(1, 3)
                
            ElseIf .BotData.Clase = eClass.Warrior Or .BotData.Clase = eClass.Bandit Or .BotData.Clase = eClass.Hunter Or .BotData.Clase = eClass.Thief _
                    Or .BotData.Clase = eClass.Pirat Or .BotData.Clase = eClass.Worker Then
                .BotData.raza = eRaza.Enano
                
            End If
            
           
            .BotData.genero = RandomNumber(1, 2)
            
            ' ya le damos clase, raza y genero
            
            'ahora le damos la cabeza y el cuerpo desnudo
            
            Select Case .BotData.raza
                Case eRaza.Humano
                    If .BotData.genero = Hombre Then
                        .Char.Head = RandomNumber(eCabezas.HUMANO_H_PRIMER_CABEZA, eCabezas.HUMANO_H_ULTIMA_CABEZA)
                    Else
                        .Char.Head = RandomNumber(eCabezas.HUMANO_M_PRIMER_CABEZA, eCabezas.HUMANO_M_ULTIMA_CABEZA)
                    End If
                    .Char.body = eCabezas.HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    If .BotData.genero = Hombre Then
                        .Char.Head = RandomNumber(eCabezas.ELFO_H_PRIMER_CABEZA, eCabezas.ELFO_H_ULTIMA_CABEZA)
                    Else
                        .Char.Head = RandomNumber(eCabezas.ELFO_M_PRIMER_CABEZA, eCabezas.ELFO_M_ULTIMA_CABEZA)
                    End If
                    .Char.body = eCabezas.ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    If .BotData.genero = Hombre Then
                        .Char.Head = RandomNumber(eCabezas.GNOMO_H_PRIMER_CABEZA, eCabezas.GNOMO_H_ULTIMA_CABEZA)
                    Else
                        .Char.Head = RandomNumber(eCabezas.GNOMO_M_PRIMER_CABEZA, eCabezas.GNOMO_M_ULTIMA_CABEZA)
                    End If
                    .Char.body = eCabezas.GNOMO_H_CUERPO_DESNUDO
                                            
                Case eRaza.Enano
                    If .BotData.genero = Hombre Then
                        .Char.Head = RandomNumber(eCabezas.ENANO_H_PRIMER_CABEZA, eCabezas.ENANO_H_ULTIMA_CABEZA)
                    Else
                        .Char.Head = RandomNumber(eCabezas.ENANO_M_PRIMER_CABEZA, eCabezas.ENANO_M_ULTIMA_CABEZA)
                    End If
                    .Char.body = eCabezas.ENANO_H_CUERPO_DESNUDO
                                            
                Case eRaza.Drow
                    If .BotData.genero = Hombre Then
                        .Char.Head = RandomNumber(eCabezas.DROW_H_PRIMER_CABEZA, eCabezas.DROW_H_ULTIMA_CABEZA)
                    Else
                        .Char.Head = RandomNumber(eCabezas.DROW_M_PRIMER_CABEZA, eCabezas.DROW_M_ULTIMA_CABEZA)
                    End If
                    .Char.body = eCabezas.DROW_H_CUERPO_DESNUDO
                    
            End Select
            
            'ya tiene raza, clase, genero, cabeza y cuerpo desnudo.
            
            .Char.heading = eHeading.SOUTH
            
            .BotData.stats.UserAtributos(eAtributos.Fuerza) = DiceMaximum
            .BotData.stats.UserAtributos(eAtributos.Agilidad) = DiceMaximum
            .BotData.stats.UserAtributos(eAtributos.Carisma) = DiceMaximum
            .BotData.stats.UserAtributos(eAtributos.Constitucion) = DiceMaximum
            .BotData.stats.UserAtributos(eAtributos.Inteligencia) = DiceMaximum
            
            Call SetAttributesToNewBot(nIndex, .BotData.Clase, .BotData.raza)
            
            Call AddItemsToNewBot(nIndex, .BotData.Clase, .BotData.raza)
            
            .Pos.Map = 1
            .Pos.x = 50
            .Pos.y = 50
            Call guardar_Bot(False, index)
            
          '  Call SpawnBot(index, True)
            
            'Call connectBot()
            
            
        End With
    End With
End Sub

Private Sub SetAttributesToNewBot(ByVal npcIndex As Integer, ByVal UserClase As eClass, ByVal UserRaza As eRaza)

    With Npclist(npcIndex)
        '[Pablo (Toxic Waste) 9/01/08]
        .BotData.stats.UserAtributos(eAtributos.Fuerza) = .BotData.stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
        .BotData.stats.UserAtributos(eAtributos.Agilidad) = .BotData.stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
        .BotData.stats.UserAtributos(eAtributos.Inteligencia) = .BotData.stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
        .BotData.stats.UserAtributos(eAtributos.Carisma) = .BotData.stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
        .BotData.stats.UserAtributos(eAtributos.Constitucion) = .BotData.stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
        '[/Pablo (Toxic Waste)]
    
        Dim i As Long
        For i = 1 To NUMSKILLS
            .BotData.stats.UserSkills(i) = 0
            'Call CheckEluSkill(UserIndex, i, True)
        Next i
    
        .BotData.stats.SkillPts = 10
    
        Dim MiInt As Long

        MiInt = RandomNumber(1, .BotData.stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
        .BotData.stats.MaxHp = 15 + MiInt
        .BotData.stats.MinHp = 15 + MiInt
    
        MiInt = RandomNumber(1, .BotData.stats.UserAtributos(eAtributos.Agilidad) \ 6)

        If MiInt = 1 Then MiInt = 2
    
        .BotData.stats.MaxSta = 20 * MiInt
        .BotData.stats.MinSta = 20 * MiInt
    
        .BotData.stats.MaxAGU = 100
        .BotData.stats.MinAGU = 100
    
        .BotData.stats.MaxHam = 100
        .BotData.stats.MinHam = 100
    
        '<-----------------MANA----------------------->
        If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
            MiInt = .BotData.stats.UserAtributos(eAtributos.Inteligencia) * 3
            .BotData.stats.MaxMAN = MiInt
            .BotData.stats.MinMAN = MiInt
        ElseIf UserClase = eClass.Cleric Or _
               UserClase = eClass.Druid Or _
               UserClase = eClass.Bard Or _
               UserClase = eClass.Assasin Or _
               UserClase = eClass.Bandit Or _
               UserClase = eClass.Paladin Then
            .BotData.stats.MaxMAN = 50
            .BotData.stats.MinMAN = 50
        Else
            .BotData.stats.MaxMAN = 0
            .BotData.stats.MinMAN = 0
        End If
    
        If UserClase = eClass.Cleric Or _
           UserClase = eClass.Druid Or _
           UserClase = eClass.Bard Or _
           UserClase = eClass.Assasin Or _
           UserClase = eClass.Bandit Or _
           UserClase = eClass.Paladin Or _
           UserClase = eClass.Mage Then

            .BotData.stats.UserHechizos(1) = 2
        
            If UserClase = eClass.Druid Then .BotData.stats.UserHechizos(2) = 46

        End If
    
        .BotData.stats.MaxHIT = 2
        .BotData.stats.MinHIT = 1
    
        .BotData.stats.Gld = 0
    
        .BotData.stats.Exp = 0
        .BotData.stats.ELU = 300
        .BotData.stats.ELV = 1
    End With

End Sub

Private Sub AddItemsToNewBot(ByVal npcIndex As Integer, ByVal UserClase As eClass, ByVal UserRaza As eRaza)
'*************************************************
'Author: Lucas Recoaro (Recox)
'Last modified: 19/03/2019
'Agrega items al usuario recien creado
'*************************************************
    Dim Slot As Byte
    Dim IsPaladin As Boolean

    IsPaladin = UserClase = eClass.Paladin
    With Npclist(npcIndex)
        'Pociones Rojas (Newbie)
        Slot = 1
        .invent.Object(Slot).ObjIndex = 857
        .invent.Object(Slot).Amount = 200

        'Pociones azules (Newbie)
        If .BotData.stats.MaxMAN > 0 Or IsPaladin Then
            Slot = Slot + 1
            .invent.Object(Slot).ObjIndex = 856
            .invent.Object(Slot).Amount = 200

        Else
            'Pociones amarillas (Newbie)
            Slot = Slot + 1
            .invent.Object(Slot).ObjIndex = 855
            .invent.Object(Slot).Amount = 100

            'Pociones verdes (Newbie)
            Slot = Slot + 1
            .invent.Object(Slot).ObjIndex = 858
            .invent.Object(Slot).Amount = 50

        End If

        ' Ropa (Newbie)
        Slot = Slot + 1
        Select Case UserRaza
            Case eRaza.Humano
                .invent.Object(Slot).ObjIndex = 463
            Case eRaza.Elfo
                .invent.Object(Slot).ObjIndex = 464
            Case eRaza.Drow
                .invent.Object(Slot).ObjIndex = 465
            Case eRaza.Enano, eRaza.Gnomo
                .invent.Object(Slot).ObjIndex = 466
        End Select

        ' Equipo ropa
        .invent.Object(Slot).Amount = 1
        .invent.Object(Slot).Equipped = 1

        .invent.ArmourEqpSlot = Slot
        .invent.ArmourEqpObjIndex = .invent.Object(Slot).ObjIndex

        'Arma (Newbie)
        Slot = Slot + 1
        Select Case UserClase
            Case eClass.Hunter
                ' Arco (Newbie)
                .invent.Object(Slot).ObjIndex = 859
            Case eClass.Worker
                ' Herramienta (Newbie)
                .invent.Object(Slot).ObjIndex = RandomNumber(561, 565)
            Case Else
                ' Daga (Newbie)
                .invent.Object(Slot).ObjIndex = 460
        End Select

        ' Equipo arma
        .invent.Object(Slot).Amount = 1
        .invent.Object(Slot).Equipped = 1

        .invent.WeaponEqpObjIndex = .invent.Object(Slot).ObjIndex
        .invent.WeaponEqpSlot = Slot

        .Char.WeaponAnim = GetWeaponAnimBot(UserRaza, .invent.WeaponEqpObjIndex)

        ' Municiones (Newbie)
        If UserClase = eClass.Hunter Then
            Slot = Slot + 1
            .invent.Object(Slot).ObjIndex = 860
            .invent.Object(Slot).Amount = 150

            ' Equipo flechas
            .invent.Object(Slot).Equipped = 1
            .invent.MunicionEqpSlot = Slot
            .invent.MunicionEqpObjIndex = 860
        End If

        ' Manzanas (Newbie)
        Slot = Slot + 1
        .invent.Object(Slot).ObjIndex = 467
        .invent.Object(Slot).Amount = 100

        ' Jugos (Nwbie)
        Slot = Slot + 1
        .invent.Object(Slot).ObjIndex = 468
        .invent.Object(Slot).Amount = 100

        ' Sin casco y escudo
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco

        ' Total Items
        .invent.NroItems = Slot

        Dim i As Long
        
        'For i = 1 To MAXAMIGOS
       '     .Amigos(i).Nombre = vbNullString
      '      .Amigos(i).Ignorado = 0
        '    .Amigos(i).index = 0
       ' Next i

     End With
     
End Sub

Private Function GetWeaponAnimBot(ByVal raza As eRaza, _
                              ByVal ObjIndex As Integer) As Integer

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 03/29/10
    '
    '***************************************************
    Dim Tmp As Integer

        Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
        If Tmp > 0 Then
            If raza = eRaza.Enano Or raza = eRaza.Gnomo Then
                GetWeaponAnimBot = Tmp
                Exit Function

            End If

        End If
        
        GetWeaponAnimBot = ObjData(ObjIndex).WeaponAnim

End Function


Private Sub guardar_Bot(ByVal ALL As Boolean, Optional ByVal index As Byte)
    Dim LoopC As Long
    Dim manager As clsIniManager
    Set manager = New clsIniManager
    
    
    With Npclist(user_Bot(index).npcIndex)
       ' Call manager.Initialize(CharPath & UCase(user_Bot(index).nick) & ".chr")
        
        Call manager.ChangeValue("FACCIONES", "CIUDMATADOS", .BotData.faccion.CiudadanosMatados)
        Call manager.ChangeValue("FACCIONES", "CRIMMATADOS", .BotData.faccion.CriminalesMatados)
        Call manager.ChangeValue("FACCIONES", "EJERCITOCAOS", .BotData.faccion.FuerzasCaos)
        Call manager.ChangeValue("FACCIONES", "EJERCITOREAL", .BotData.faccion.ArmadaReal)
        Call manager.ChangeValue("FACCIONES", "NEXTRECOMPENSA", .BotData.faccion.NextRecompensa)
        Call manager.ChangeValue("FACCIONES", "Reenlistadas", .BotData.faccion.Reenlistadas)
        Call manager.ChangeValue("FACCIONES", "NivelIngreso", .BotData.faccion.NivelIngreso)
        Call manager.ChangeValue("FACCIONES", "FechaIngreso", .BotData.faccion.FechaIngreso)
        Call manager.ChangeValue("FACCIONES", "MatadosIngreso", .BotData.faccion.MatadosIngreso)
        Call manager.ChangeValue("FACCIONES", "NEXTRECOMPENSA", .BotData.faccion.NextRecompensa)
        
        Call manager.ChangeValue("FLAGS", "DESNUDO", .BotData.flags.Desnudo)
        Call manager.ChangeValue("FLAGS", "ENVENENADO", .BotData.flags.Envenenado)
        Call manager.ChangeValue("FLAGS", "MUERTO", .BotData.flags.Muerto)
        Call manager.ChangeValue("FLAGS", "NAVEGANDO", .BotData.flags.Navegando)
        
        Call manager.ChangeValue("INIT", "ARMA", .Char.WeaponAnim)
        Call manager.ChangeValue("INIT", "BODY", .Char.body)
        Call manager.ChangeValue("INIT", "CASCO", .Char.CascoAnim)
        Call manager.ChangeValue("INIT", "ESCUDO", .Char.ShieldAnim)
        Call manager.ChangeValue("INIT", "CLASE", .BotData.Clase)
        Call manager.ChangeValue("INIT", "RAZA", .BotData.raza)
        Call manager.ChangeValue("INIT", "HEAD", .Char.Head)
        Call manager.ChangeValue("INIT", "HEADING", .Char.heading)
        
        Call manager.ChangeValue("INIT", "GENERO", .BotData.genero)
        Call manager.ChangeValue("INIT", "POSITION", .Pos.Map & "-" & .Pos.x & "-" & .Pos.y)
        
        Call manager.ChangeValue("INVENTORY", "ANILLOSLOT", .invent.AnilloEqpSlot)
        Call manager.ChangeValue("INVENTORY", "ARMOUREQPSLOT", .invent.ArmourEqpSlot)
        Call manager.ChangeValue("INVENTORY", "BARCOSLOT", .invent.BarcoSlot)
        Call manager.ChangeValue("INVENTORY", "CANTIDADITEMS", .invent.NroItems)
        Call manager.ChangeValue("INVENTORY", "CASCOEQPSLOT", .invent.CascoEqpSlot)
        Call manager.ChangeValue("INVENTORY", "ESCUDOEQPSLOT", .invent.EscudoEqpSlot)
        Call manager.ChangeValue("INVENTORY", "MONTURAEQPSLOT", .invent.MonturaEqpSlot)
        Call manager.ChangeValue("INVENTORY", "MUNICIONSLOT", .invent.MunicionEqpSlot)
        Call manager.ChangeValue("INVENTORY", "WEAPONEQPSLOT", .invent.WeaponEqpSlot)
        
        Call manager.ChangeValue("REP", "Asesino", CLng(.Reputacion.AsesinoRep))
        Call manager.ChangeValue("REP", "Bandido", CLng(.Reputacion.BandidoRep))
        Call manager.ChangeValue("REP", "Burguesia", CLng(.Reputacion.BurguesRep))
        Call manager.ChangeValue("REP", "Ladrones", CLng(.Reputacion.LadronesRep))
        Call manager.ChangeValue("REP", "Nobles", CLng(.Reputacion.NobleRep))
        Call manager.ChangeValue("REP", "Plebe", CLng(.Reputacion.PlebeRep))
        Call manager.ChangeValue("REP", "Promedio", CLng(.Reputacion.Promedio))
        
    
        Call manager.ChangeValue("STATS", "GLD", CLng(.BotData.stats.Gld))
        Call manager.ChangeValue("STATS", "BANCO", CLng(.BotData.stats.Banco))
    
        Call manager.ChangeValue("STATS", "MaxHP", CInt(.BotData.stats.MaxHp))
        Call manager.ChangeValue("STATS", "MinHP", CInt(.BotData.stats.MinHp))
    
        Call manager.ChangeValue("STATS", "MaxSTA", CInt(.BotData.stats.MaxSta))
        Call manager.ChangeValue("STATS", "MinSTA", CInt(.BotData.stats.MinSta))
    
        Call manager.ChangeValue("STATS", "MaxMAN", CInt(.BotData.stats.MaxMAN))
        Call manager.ChangeValue("STATS", "MinMAN", CInt(.BotData.stats.MinMAN))
    
        Call manager.ChangeValue("STATS", "MaxHIT", CInt(.BotData.stats.MaxHIT))
        Call manager.ChangeValue("STATS", "MinHIT", CInt(.BotData.stats.MinHIT))
    
        Call manager.ChangeValue("STATS", "MaxAGU", CByte(.BotData.stats.MaxAGU))
        Call manager.ChangeValue("STATS", "MinAGU", CByte(.BotData.stats.MinAGU))
    
        Call manager.ChangeValue("STATS", "MaxHAM", CByte(.BotData.stats.MaxHam))
        Call manager.ChangeValue("STATS", "MinHAM", CByte(.BotData.stats.MinHam))
    
        Call manager.ChangeValue("STATS", "SkillPtsLibres", CInt(.BotData.stats.SkillPts))
    
        Call manager.ChangeValue("STATS", "EXP", CDbl(.BotData.stats.Exp))
        Call manager.ChangeValue("STATS", "ELV", CByte(.BotData.stats.ELV))
      
        Call manager.ChangeValue("STATS", "ELU", CLng(.BotData.stats.ELU))
    
        Call manager.ChangeValue("MUERTES", "UserMuertes", CLng(.BotData.stats.UsuariosMatados))
        Call manager.ChangeValue("MUERTES", "NpcsMuertes", CInt(.BotData.stats.NPCsMuertos))
      
      
       If Not .BotData.flags.TomoPocion Then

            For LoopC = 1 To UBound(.BotData.stats.UserAtributos)
                Call manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.BotData.stats.UserAtributos(LoopC)))
            Next LoopC

        Else

            For LoopC = 1 To UBound(.BotData.stats.UserAtributos)
                '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
                Call manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.BotData.stats.UserAtributosBackUP(LoopC)))
            Next LoopC

        End If
    
        For LoopC = 1 To UBound(.BotData.stats.UserSkills)
            Call manager.ChangeValue("SKILLS", "SK" & LoopC, CStr(.BotData.stats.UserSkills(LoopC)))
            Call manager.ChangeValue("SKILLS", "ELUSK" & LoopC, CStr(.BotData.stats.EluSkills(LoopC)))
            Call manager.ChangeValue("SKILLS", "EXPSK" & LoopC, CStr(.BotData.stats.ExpSkills(LoopC)))
        Next LoopC
        
        Dim cad As String
    
        For LoopC = 1 To MAXUSERHECHIZOS
            cad = .BotData.stats.UserHechizos(LoopC)
            Call manager.ChangeValue("HECHIZOS", "H" & LoopC, cad)
        Next
        
        Dim i As Long
        
        For i = 1 To MAX_INVENTORY_SLOTS
            Call manager.ChangeValue("INVENTORY", "OBJ" & i, .invent.Object(i).ObjIndex & "-" & .invent.Object(i).Amount & "-" & .invent.Object(i).Equipped)
        Next i
        
        Call manager.DumpFile(CharPath & UCase$(user_Bot(index).nick) & ".chr")
        
        Set manager = Nothing
    End With
End Sub

Public Function OpenNPCBot(ByVal botNumber As Integer) As Integer

    Dim npcIndex As Integer

    Dim Leer     As clsIniManager

    Dim LoopC    As Long

    Dim ln       As String
    
    Set Leer = New clsIniManager 'LeerNPCs
    
    'If requested index is invalid, abort
    'If Not leer.KeyExists("NPC" & NpcNumber) Then
     '   OpenNPCBot = MAXNPCS + 1
     '   Exit Function

  '  End If
    If Not FileExist(CharPath & UCase$(user_Bot(botNumber).nick) & ".chr", vbNormal) Then
        OpenNPCBot = -1
        Exit Function
    End If
    
    Leer.Initialize CharPath & UCase$(user_Bot(botNumber).nick) & ".chr"
    
    npcIndex = NextOpenNPC
    
    If npcIndex > MAXNPCS Then 'Limite de npcs
        OpenNPCBot = npcIndex
        Exit Function

    End If
    
    With Npclist(npcIndex)
        .esBot = True
        .BotData.botindex = botNumber
        
        .Name = user_Bot(botNumber).nick 'leer.GetValue("NPC" & NpcNumber, "Name")
        
        '.Desc = leer.GetValue("NPC" & NpcNumber, "Desc")
        
    
            .BotData.stats.Gld = val(Leer.GetValue("STATS", "GLD"))
            .BotData.stats.Banco = val(Leer.GetValue("STATS", "BANCO"))
            .BotData.stats.MaxHp = val(Leer.GetValue("STATS", "MaxHP"))
            .BotData.stats.MinHp = val(Leer.GetValue("STATS", "MinHP"))
            .BotData.stats.MaxSta = val(Leer.GetValue("STATS", "MaxSTA"))
            .BotData.stats.MinSta = val(Leer.GetValue("STATS", "MinSTA"))
            .BotData.stats.MinMAN = val(Leer.GetValue("STATS", "MaxMAN"))
            .BotData.stats.MaxMAN = val(Leer.GetValue("STATS", "MinMAN"))
            .BotData.stats.MaxHIT = val(Leer.GetValue("STATS", "MaxHIT"))
            .BotData.stats.MinHIT = val(Leer.GetValue("STATS", "MinHIT"))
            .BotData.stats.MaxAGU = val(Leer.GetValue("STATS", "MaxAGU"))
            .BotData.stats.MinAGU = val(Leer.GetValue("STATS", "MinAGU"))
            .BotData.stats.MaxHam = val(Leer.GetValue("STATS", "MaxHAM"))
            .BotData.stats.MinHam = val(Leer.GetValue("STATS", "MinHam"))
            .BotData.stats.SkillPts = val(Leer.GetValue("STATS", "SkillPtsLibres"))
            .BotData.stats.Exp = val(Leer.GetValue("STATS", "EXP"))
            .BotData.stats.ELV = val(Leer.GetValue("STATS", "ELV"))
            .BotData.stats.ELU = val(Leer.GetValue("STATS", "ELU"))
            .BotData.stats.UsuariosMatados = val(Leer.GetValue("MUERTES", "UserMuertes"))
            .BotData.stats.NPCsMuertos = val(Leer.GetValue("MUERTES", "NpcsMuertes"))
        
        With .BotData.faccion
            .CiudadanosMatados = val(Leer.GetValue("FACCIONES", "CIUDMATADOS"))
            .CriminalesMatados = val(Leer.GetValue("FACCIONES", "CRIMMATADOS")) ', .BotData.faccion.CriminalesMatados)
            .FuerzasCaos = val(Leer.GetValue("FACCIONES", "EJERCITOCAOS"))  ', .BotData.faccion.FuerzasCaos)
            .ArmadaReal = val(Leer.GetValue("FACCIONES", "ARMADAREAL"))
            .NextRecompensa = val(Leer.GetValue("FACCIONES", "NEXTRECOMPENSA"))
            .Reenlistadas = val(Leer.GetValue("FACCIONES", "Reenlistadas"))
            .NivelIngreso = val(Leer.GetValue("FACCIONES", "NivelIngreso"))
            .FechaIngreso = val(Leer.GetValue("FACCIONES", "FechaIngreso"))
            .MatadosIngreso = val(Leer.GetValue("FACCIONES", "MatadosIngreso"))
            .NextRecompensa = val(Leer.GetValue("FACCIONES", "NE"))
        End With
        
        With .BotData.flags
                   
            .Desnudo = val(Leer.GetValue("FLAGS", "DESNUDO")) ' .BotData.flags.Desnudo)
            .Envenenado = val(Leer.GetValue("FLAGS", "ENVENENADO")) ' leer.ChangeValue("FLAGS", "ENVENENADO", .BotData.flags.Envenenado)
            .Muerto = val(Leer.GetValue("FLAGS", "MUERTO")) ', .BotData.flags.Muerto)
            .Navegando = val(Leer.GetValue("FLAGS", "NAVEGANDO"))  '.BotData.flags.Navegando)
            
        End With
        
        .BotData.genero = val(Leer.GetValue("INIT", "GENERO"))
        .BotData.Clase = val(Leer.GetValue("INIT", "CLASE"))
        .BotData.raza = val(Leer.GetValue("INIT", "RAZA"))
        .Char.heading = val(Leer.GetValue("INIT", "HEADING"))
        
        .OrigChar.heading = eHeading.SOUTH
        .OrigChar.WeaponAnim = val(Leer.GetValue("INIT", "ARMA"))
        .OrigChar.body = val(Leer.GetValue("INIT", "BODY"))
        .OrigChar.CascoAnim = val(Leer.GetValue("INIT", "CASCO"))
        .OrigChar.ShieldAnim = val(Leer.GetValue("INIT", "ESCUDO"))
        .OrigChar.Head = val(Leer.GetValue("INIT", "HEAD"))
        
        .invent.NroItems = CInt(Leer.GetValue("Inventory", "CantidadItems"))
        
        .Pos.Map = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
        .Pos.x = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
        .Pos.y = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))
        
        With .invent
            .AnilloEqpSlot = val(Leer.GetValue("INVENTORY", "ANILLOSLOT"))
            .ArmourEqpSlot = val(Leer.GetValue("INVENTORY", "ARMOUREQPSLOT"))
            .BarcoSlot = val(Leer.GetValue("INVENTORY", "BARCOSLOT"))
            .NroItems = val(Leer.GetValue("INVENTORY", "CANTIDADITEMS"))
            .CascoEqpSlot = val(Leer.GetValue("INVENTORY", "CASCOEQPSLOT"))
            .EscudoEqpSlot = val(Leer.GetValue("INVENTORY", "ESCUDOEQPSLOT"))
            .MonturaEqpSlot = val(Leer.GetValue("INVENTORY", "MONTURAEQPSLOT"))
            .MunicionEqpSlot = val(Leer.GetValue("INVENTORY", "MUNICIONSLOT"))
            .WeaponEqpSlot = val(Leer.GetValue("INVENTORY", "WEAPONEQPSLOT"))
        End With
    
        
          'Lista de objetos
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            ln = Leer.GetValue("Inventory", "Obj" & LoopC)
            If (val(ReadField(1, ln, 45))) > NumObjDatas Then
                .invent.Object(LoopC).ObjIndex = 0
                .invent.Object(LoopC).Amount = 0
                .invent.Object(LoopC).Equipped = 0
            Else
                .invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
                .invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
            End If

        Next LoopC
        
        .invent.WeaponEqpSlot = CByte(Leer.GetValue("Inventory", "WeaponEqpSlot"))
        .invent.ArmourEqpSlot = CByte(Leer.GetValue("Inventory", "ArmourEqpSlot"))
        .invent.EscudoEqpSlot = CByte(Leer.GetValue("Inventory", "EscudoEqpSlot"))
        .invent.CascoEqpSlot = CByte(Leer.GetValue("Inventory", "CascoEqpSlot"))
        .invent.BarcoSlot = CByte(Leer.GetValue("Inventory", "BarcoSlot"))
        
        'Si no existe MonturaEqpSlot, se agrega al charfile.
        If Not Leer.KeyExists("MonturaEqpSlot") Then
            .invent.MonturaEqpSlot = 0
        Else
            .invent.MonturaEqpSlot = CByte(Leer.GetValue("Inventory", "MonturaEqpSlot"))
        End If
        
        .invent.MunicionEqpSlot = CByte(Leer.GetValue("Inventory", "MunicionSlot"))
        .invent.AnilloEqpSlot = CByte(Leer.GetValue("Inventory", "AnilloSlot"))
        .invent.MochilaEqpSlot = val(Leer.GetValue("Inventory", "MochilaSlot"))
        
        For LoopC = 1 To NUMATRIBUTOS
            .BotData.stats.UserAtributos(LoopC) = CByte(Leer.GetValue("ATRIBUTOS", "AT" & LoopC))
            .BotData.stats.UserAtributosBackUP(LoopC) = CByte(.BotData.stats.UserAtributos(LoopC))
        Next LoopC
    
        For LoopC = 1 To NUMSKILLS
            .BotData.stats.UserSkills(LoopC) = CByte(Leer.GetValue("SKILLS", "SK" & LoopC))
            .BotData.stats.EluSkills(LoopC) = CLng(Leer.GetValue("SKILLS", "ELUSK" & LoopC))
            .BotData.stats.ExpSkills(LoopC) = CLng(Leer.GetValue("SKILLS", "EXPSK" & LoopC))
        Next LoopC
        
        For LoopC = 1 To MAXUSERHECHIZOS
            .BotData.stats.UserHechizos(LoopC) = CInt(Leer.GetValue("Hechizos", "H" & LoopC))
        Next LoopC
        
        
        With .Reputacion
            .AsesinoRep = CLng(Leer.GetValue("REP", "Asesino"))
            .BandidoRep = CLng(Leer.GetValue("REP", "Bandido"))
            .BurguesRep = CLng(Leer.GetValue("REP", "Burguesia"))
            .LadronesRep = CLng(Leer.GetValue("REP", "Ladrones"))
            .NobleRep = CLng(Leer.GetValue("REP", "Nobles"))
            .PlebeRep = CLng(Leer.GetValue("REP", "Plebe"))
            .Promedio = CLng(Leer.GetValue("REP", "Promedio"))
    
        End With
        
        
                 'Obtiene el indice-objeto del arma
                If .invent.WeaponEqpSlot > 0 Then
                    .invent.WeaponEqpObjIndex = .invent.Object(.invent.WeaponEqpSlot).ObjIndex
        
                End If
        
                'Obtiene el indice-objeto del armadura
                If .invent.ArmourEqpSlot > 0 Then
                    .invent.ArmourEqpObjIndex = .invent.Object(.invent.ArmourEqpSlot).ObjIndex
                    .BotData.flags.Desnudo = 0
                Else
                    .BotData.flags.Desnudo = 1
                    
                End If
        
                'Obtiene el indice-objeto del escudo
                If .invent.EscudoEqpSlot > 0 Then
                    .invent.EscudoEqpObjIndex = .invent.Object(.invent.EscudoEqpSlot).ObjIndex
        
                End If
                
                'Obtiene el indice-objeto del casco
                If .invent.CascoEqpSlot > 0 Then
                    .invent.CascoEqpObjIndex = .invent.Object(.invent.CascoEqpSlot).ObjIndex
        
                End If
        
                'Obtiene el indice-objeto barco
                If .invent.BarcoSlot > 0 Then
                    .invent.BarcoObjIndex = .invent.Object(.invent.BarcoSlot).ObjIndex
        
                End If
        
                'Obtiene el indice-objeto municion
                If .invent.MunicionEqpSlot > 0 Then
                    .invent.MunicionEqpObjIndex = .invent.Object(.invent.MunicionEqpSlot).ObjIndex
        
                End If
        
                '[Alejo]
                'Obtiene el indice-objeto anilo
                If .invent.AnilloEqpSlot > 0 Then
                    .invent.AnilloEqpObjIndex = .invent.Object(.invent.AnilloEqpSlot).ObjIndex
        
                End If
        
                If .invent.MonturaObjIndex > 0 Then
                    .invent.MonturaObjIndex = .invent.Object(.invent.MonturaObjIndex).ObjIndex
                End If
        
                If .BotData.flags.Muerto = 0 Then
                    .Char = .OrigChar
                Else
                    .Char.body = iCuerpoMuerto
                    .Char.Head = iCabezaMuerto
                    .Char.WeaponAnim = NingunArma
                    .Char.ShieldAnim = NingunEscudo
                    .Char.CascoAnim = NingunCasco
                    .Char.heading = eHeading.SOUTH
                End If
                
        

        With .flags
            .NPCActive = True

            .AfectaParalisis = 1 'val(leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
        End With

    End With
    
    'Update contadores de NPCs
    If npcIndex > LastNPC Then LastNPC = npcIndex
    NumNPCs = NumNPCs + 1
    
    'Devuelve el nuevo Indice
    OpenNPCBot = npcIndex

End Function

Private Function nextBot() As Integer
    Dim i As Integer
    i = RandomNumber(1, MAX_BOTS)
    
    Do While user_Bot(i).online = True
        i = RandomNumber(1, MAX_BOTS)
        
    Loop

    If user_Bot(i).creado = True Then
        nextBot = i
    Else 'no hay un personaje creado en ese slot, lo mandamos a crear
        crear_Bot i
        nextBot = i
    End If
    
End Function

Private Sub connectBot()

    Dim npcIndex As Integer, botindex As Integer
    botindex = nextBot
    
    If botindex > 0 Then
        user_Bot(botindex).npcIndex = SpawnBot(botindex, True)
        user_Bot(botindex).online = True
        BotsOnline = BotsOnline + 1
    End If
    
End Sub
'npclist().
'esBot As Boolean
'BotData As tBotData
Private Sub disconnectbot()
    Dim i As Long
    For i = 1 To MAX_BOTS
        If user_Bot(i).online Then
            'If user_Bot(i).minutosOnline >= MinutosLogeados Then
                user_Bot(i).online = False
                Call guardar_Bot(False, i)
                Call QuitarNPC(user_Bot(i).npcIndex)
            'End If
        End If
    Next i
End Sub

Function SpawnBot(ByVal botindex As Integer, _
                  ByVal FX As Boolean) As Integer

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/15/2008
    '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
    '06/15/2008 -> Optimize el codigo. (NicoNZ)
    '***************************************************
    Dim newPos         As WorldPos

    Dim altpos         As WorldPos

    Dim nIndex         As Integer

    Dim PosicionValida As Boolean

    Dim PuedeAgua      As Boolean

    Dim PuedeTierra    As Boolean

    Dim Map            As Integer

    Dim x              As Integer

    Dim y              As Integer

    nIndex = OpenNPCBot(botindex)    'Conseguimos un indice
    
    If nIndex > MAXNPCS Then
        SpawnBot = 0
        Exit Function
    End If

    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
    
    Call ClosestLegalPos(Npclist(nIndex).Pos, newPos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
    Call ClosestLegalPos(Npclist(nIndex).Pos, altpos, PuedeAgua)
    
    'Si X e Y son iguales a 0 significa que no se encontro posicion valida

    If newPos.x <> 0 And newPos.y <> 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        Npclist(nIndex).Pos.Map = newPos.Map
        Npclist(nIndex).Pos.x = newPos.x
        Npclist(nIndex).Pos.y = newPos.y
        PosicionValida = True
    Else

        If altpos.x <> 0 And altpos.y <> 0 Then
            Npclist(nIndex).Pos.Map = altpos.Map
            Npclist(nIndex).Pos.x = altpos.x
            Npclist(nIndex).Pos.y = altpos.y
            PosicionValida = True
        Else
            PosicionValida = False
        End If

    End If

    If Not PosicionValida Then
        Call QuitarNPC(nIndex)
        SpawnBot = 0
        Exit Function

    End If

    'asignamos las nuevas coordenas
    Map = Npclist(nIndex).Pos.Map
    x = Npclist(nIndex).Pos.x
    y = Npclist(nIndex).Pos.y

    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, x, y)

    If FX Then
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, x, y))
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))
    End If

    SpawnBot = nIndex

End Function


Public Sub ai_bots()
    Dim i As Long
    For i = 1 To MAX_BOTS
        With user_Bot(i)
            If .online = True Then
                
            End If
        End With
    Next i
End Sub

Public Function RandomizaCharla(ByVal numOpciones As Byte, ByVal opcion1 As String, Optional ByVal opcion2 As String, _
                                                          Optional ByVal opcion3 As String, Optional ByVal opcion4 As String, Optional ByVal opcion5 As String, _
                                                          Optional ByVal opcion6 As String, Optional ByVal opcion7 As String, Optional ByVal opcion8 As String, _
                                                          Optional ByVal opcion9 As String, Optional ByVal opcion10 As String) As String
    
    Dim random As Byte
    
    If numOpciones > 10 Then numOpciones = 10
    random = RandomNumber(1, numOpciones)
    
    
    Select Case random
        Case 1
            RandomizaCharla = opcion1
        
        Case 2
            RandomizaCharla = opcion2
        
        Case 3
            RandomizaCharla = opcion3
        
        Case 4
            RandomizaCharla = opcion4
        
        Case 5
            RandomizaCharla = opcion5
        
        Case 6
            RandomizaCharla = opcion6
        
        Case 7
            RandomizaCharla = opcion7
        
        Case 8
            RandomizaCharla = opcion8
        
        Case 9
            RandomizaCharla = opcion9
        
        Case 10
            RandomizaCharla = opcion10
    End Select
End Function





