Attribute VB_Name = "modDatabase"
'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018

Option Explicit

Public Database_Enabled As Boolean
Public Database_Host As String
Public Database_Name As String
Public Database_Username As String
Public Database_Password As String
Public Database_Connection As ADODB.Connection
Public Database_RecordSet As ADODB.Recordset
 
Public Sub Database_Connect()
'***************************************************
'Author: Juan Andres Dalmasso
'Last Modification: 18/09/2018
'***************************************************
On Error GoTo ErrorHandler
 
Set Database_Connection = New ADODB.Connection
 
Database_Connection.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Database_Host & ";DATABASE=" & Database_Name & ";UID=" & Database_Username & ";PWD=" & Database_Password & "; OPTION=3"
Database_Connection.CursorLocation = adUseClient
Database_Connection.Open

Exit Sub
ErrorHandler:
    Call LogDatabaseError("Unable to connect to Mysql Database: " & Err.Number & " - " & Err.description)
End Sub

Public Sub Database_Close()
'***************************************************
'Author: Juan Andres Dalmasso
'Last Modification: 18/09/2018
'***************************************************
On Error GoTo ErrorHandler
     
Database_Connection.Close
Set Database_Connection = Nothing
     
Exit Sub
     
ErrorHandler:
    Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.description)
End Sub

Sub SaveUserToDatabase(ByVal UserIndex As Integer, Optional ByVal SaveTimeOnline As Boolean = True)
'*************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last modified: 14/10/2018
'Saves the User to the database
'*************************************************

On Error GoTo ErrorHandler

    With UserList(UserIndex)
        If .ID > 0 Then
            Call InsertUserToDatabase(UserIndex, SaveTimeOnline)
        Else
            Call UpdateUserToDatabase(UserIndex, SaveTimeOnline)
        End If
    End With

ErrorHandler:
        Call LogDatabaseError("Unable to save User to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)
End Sub

Sub InsertUserToDatabase(ByVal UserIndex As Integer, Optional ByVal SaveTimeOnline As Boolean = True)
'*************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last modified: 04/10/2018
'Inserts a new user to the database, then gets its ID and assigns it
'*************************************************

On Error GoTo ErrorHandler
    Dim query As String
    Dim UserId As Integer
    Dim LoopC As Byte

    Call Database_Connect

    'Basic user data
    With UserList(UserIndex)
        query = "INSERT INTO user SET "
        query = query & "name = '" & .Name & "', "
        query = query & "level = " & .Stats.ELV & ", "
        query = query & "exp = " & .Stats.Exp & ", "
        query = query & "elu = " & .Stats.ELU & ", "
        query = query & "genre_id = " & .Genero & ", "
        query = query & "race_id = " & .raza & ", "
        query = query & "class_id = " & .clase & ", "
        query = query & "home_id = " & .Hogar & ", "
        query = query & "description = '" & .desc & "', "
        query = query & "gold = " & .Stats.GLD & ", "
        query = query & "free_skillpoints = " & .Stats.SkillPts & ", "
        query = query & "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
        query = query & "pos_map = " & .Pos.Map & ", "
        query = query & "pos_x = " & .Pos.X & ", "
        query = query & "pos_x = " & .Pos.Y & ", "
        query = query & "body_id = " & .Char.body & ", "
        query = query & "head_id = " & .Char.body & ", "
        query = query & "weapon_id = " & .Char.WeaponAnim & ", "
        query = query & "helmet_id = " & .Char.CascoAnim & ", "
        query = query & "shield_id = " & .Char.ShieldAnim & ", "
        query = query & "items_amount = " & .Invent.NroItems & ", "
        query = query & "min_hp = " & .Stats.MinHp & ", "
        query = query & "max_hp = " & .Stats.MaxHp & ", "
        query = query & "min_man = " & .Stats.MinMAN & ", "
        query = query & "max_man = " & .Stats.MaxMAN & ", "
        query = query & "min_sta = " & .Stats.MinSta & ", "
        query = query & "max_sta = " & .Stats.MaxSta & ", "
        query = query & "min_ham = " & .Stats.MinHam & ", "
        query = query & "max_ham = " & .Stats.MaxHam & ", "
        query = query & "min_sed = " & .Stats.MinAGU & ", "
        query = query & "max_sed = " & .Stats.MaxAGU & ", "
        query = query & "min_hit = " & .Stats.MinHIT & ", "
        query = query & "max_hit = " & .Stats.MaxHIT & ", "
        query = query & "rep_noble = " & .Reputacion.NobleRep & ", "
        query = query & "rep_plebe = " & .Reputacion.PlebeRep & ", "
        query = query & "rep_average = " & .Reputacion.Promedio & ";"

        'Insert the user
        Call Database_Connection.Execute(query)

        'Get the user ID
        Set Database_RecordSet = Database_Connection.Execute("SELECT LAST_INSERT_ID();")
        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            UserId = 1
        End If
        UserId = val(Database_RecordSet.Fields(0).value)
        Set Database_RecordSet = Nothing

        .ID = UserId

        'User attributes
        query = "INSERT INTO attribute (user_id, number, value) VALUES "
        For LoopC = 1 To NUMATRIBUTOS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserAtributos(LoopC) & ")"
            If LoopC < NUMATRIBUTOS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

        'User spells
        query = "INSERT INTO spell (user_id, number, spell_id) VALUES "
        For LoopC = 1 To MAXUSERHECHIZOS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserHechizos(LoopC) & ")"
            If LoopC < MAXUSERHECHIZOS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

        'User inventory
        query = "INSERT INTO inventory_item (user_id, number, item_id, amount, is_equipped) VALUES "
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Invent.Object(LoopC).ObjIndex & ", "
            query = query & .Invent.Object(LoopC).Amount & ", "
            query = query & .Invent.Object(LoopC).Equipped & ")"
            If LoopC < MAX_INVENTORY_SLOTS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

        'User skills
        query = "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "
        For LoopC = 1 To NUMSKILLS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserSkills(LoopC) & ", "
            query = query & .Stats.ExpSkills(LoopC) & ", "
            query = query & .Stats.EluSkills(LoopC) & ")"
            If LoopC < NUMSKILLS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

    End With

    Call Database_Close
    Exit Sub

ErrorHandler:
        Call LogDatabaseError("Unable to INSERT User to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)
End Sub

Sub UpdateUserToDatabase(ByVal UserIndex As Integer, Optional ByVal SaveTimeOnline As Boolean = True)
'*************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last modified: 04/10/2018
'Updates an existing user in the database
'*************************************************

On Error GoTo ErrorHandler
    Dim query As String
    Dim UserId As Integer
    Dim LoopC As Byte

    Call Database_Connect

    'Basic user data
    With UserList(UserIndex)
        query = "UPDATE user SET "
        query = query & "name = '" & .Name & "', "
        query = query & "level = " & .Stats.ELV & ", "
        query = query & "exp = " & .Stats.Exp & ", "
        query = query & "elu = " & .Stats.ELU & ", "
        query = query & "genre_id = " & .Genero & ", "
        query = query & "race_id = " & .raza & ", "
        query = query & "class_id = " & .clase & ", "
        query = query & "home_id = " & .Hogar & ", "
        query = query & "description = '" & .desc & "', "
        query = query & "gold = " & .Stats.GLD & ", "
        query = query & "bank_gold = " & .Stats.Banco & ", "
        query = query & "free_skillpoints = " & .Stats.SkillPts & ", "
        query = query & "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
        query = query & "pet_amount = " & .NroMascotas & ", "
        query = query & "pos_map = " & .Pos.Map & ", "
        query = query & "pos_x = " & .Pos.X & ", "
        query = query & "pos_x = " & .Pos.Y & ", "
        query = query & "last_map = " & .flags.lastMap & ", "
        query = query & "body_id = " & .Char.body & ", "
        query = query & "head_id = " & .OrigChar.Head & ", "
        query = query & "weapon_id = " & .Char.WeaponAnim & ", "
        query = query & "helmet_id = " & .Char.CascoAnim & ", "
        query = query & "shield_id = " & .Char.ShieldAnim & ", "
        query = query & "heading = " & .Char.heading & ", "
        query = query & "items_amount = " & .Invent.NroItems & ", "
        query = query & "slot_armour = " & .Invent.ArmourEqpSlot & ", "
        query = query & "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
        query = query & "slot_helmet = " & .Invent.CascoEqpSlot & ", "
        query = query & "slot_shield = " & .Invent.EscudoEqpSlot & ", "
        query = query & "slot_ammo = " & .Invent.MunicionEqpSlot & ", "
        query = query & "slot_ship = " & .Invent.BarcoSlot & ", "
        query = query & "slot_ring = " & .Invent.AnilloEqpSlot & ", "
        query = query & "slot_bag = " & .Invent.MochilaEqpSlot & ", "
        query = query & "min_hp = " & .Stats.MinHp & ", "
        query = query & "max_hp = " & .Stats.MaxHp & ", "
        query = query & "min_man = " & .Stats.MinMAN & ", "
        query = query & "max_man = " & .Stats.MaxMAN & ", "
        query = query & "min_sta = " & .Stats.MinSta & ", "
        query = query & "max_sta = " & .Stats.MaxSta & ", "
        query = query & "min_ham = " & .Stats.MinHam & ", "
        query = query & "max_ham = " & .Stats.MaxHam & ", "
        query = query & "min_sed = " & .Stats.MinAGU & ", "
        query = query & "max_sed = " & .Stats.MaxAGU & ", "
        query = query & "min_hit = " & .Stats.MinHIT & ", "
        query = query & "max_hit = " & .Stats.MaxHIT & ", "
        query = query & "killed_npcs = " & .Stats.NPCsMuertos & ", "
        query = query & "killed_users = " & .Stats.UsuariosMatados & ", "
        query = query & "rep_asesino = " & .Reputacion.AsesinoRep & ", "
        query = query & "rep_bandido = " & .Reputacion.BandidoRep & ", "
        query = query & "rep_bugues = " & .Reputacion.BurguesRep & ", "
        query = query & "rep_ladron = " & .Reputacion.LadronesRep & ", "
        query = query & "rep_noble = " & .Reputacion.NobleRep & ", "
        query = query & "rep_plebe = " & .Reputacion.PlebeRep & ", "
        query = query & "rep_average = " & .Reputacion.Promedio & ", "
        query = query & "is_naked = " & .flags.Desnudo & ", "
        query = query & "is_poisoned = " & .flags.Envenenado & ", "
        query = query & "is_hidden = " & .flags.Escondido & ", "
        query = query & "is_hungry = " & .flags.Hambre & ", "
        query = query & "is_thirsty = " & .flags.Sed & ", "
        query = query & "is_ban = " & .flags.Ban & ", "
        query = query & "is_dead = " & .flags.Muerto & ", "
        query = query & "is_sailing = " & .flags.Navegando & ", "
        query = query & "is_paralyzed = " & .flags.Paralizado & ", "
        query = query & "counter_pena = " & .Counters.Pena & ", "
        query = query & "pertenece_consejo_real = " & (.flags.Privilegios And PlayerType.RoyalCouncil) & ", "
        query = query & "pertenece_consejo_caos = " & (.flags.Privilegios And PlayerType.ChaosCouncil) & ", "
        query = query & "pertenece_real = " & .Faccion.ArmadaReal & ", "
        query = query & "pertenece_caos = " & .Faccion.FuerzasCaos & ", "
        query = query & "ciudadanos_matados = " & .Faccion.CiudadanosMatados & ", "
        query = query & "criminales_matados = " & .Faccion.CriminalesMatados & ", "
        query = query & "recibio_armadura_real = " & .Faccion.RecibioArmaduraReal & ", "
        query = query & "recibio_armadura_caos = " & .Faccion.RecibioArmaduraCaos & ", "
        query = query & "recibio_exp_real = " & .Faccion.RecibioExpInicialReal & ", "
        query = query & "recibio_exp_caos = " & .Faccion.RecibioExpInicialCaos & ", "
        query = query & "recompensas_real = " & .Faccion.RecompensasReal & ", "
        query = query & "recompensas_caos = " & .Faccion.RecompensasCaos & ", "
        query = query & "reenlistadas = " & .Faccion.Reenlistadas & ", "
        query = query & "fecha_ingreso = " & .Faccion.FechaIngreso & ", "
        query = query & "nivel_ingreso = " & .Faccion.NivelIngreso & ", "
        query = query & "matados_ingreso = " & .Faccion.MatadosIngreso & ", "
        query = query & "siguiente_recompensa = " & .Faccion.NextRecompensa & ", "
        query = query & "guild_index = " & .GuildIndex
        query = query & "WHERE user_id = " & .ID & ";"
        Call Database_Connection.Execute(query)

        'User attributes
        query = "DELETE FROM attribute WHERE user_id = " & .ID & ";"
        Call Database_Connection.Execute(query)

        query = "INSERT INTO attribute (user_id, number, value) VALUES "
        For LoopC = 1 To NUMATRIBUTOS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserAtributos(LoopC) & ")"
            If LoopC < NUMATRIBUTOS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

        'User spells
        query = "DELETE FROM spell WHERE user_id = " & .ID & ";"
        Call Database_Connection.Execute(query)

        query = "INSERT INTO spell (user_id, number, spell_id) VALUES "
        For LoopC = 1 To MAXUSERHECHIZOS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserHechizos(LoopC) & ")"
            If LoopC < MAXUSERHECHIZOS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

        'User inventory
        query = "DELETE FROM inventory_item WHERE user_id = " & .ID & ";"
        Call Database_Connection.Execute(query)

        query = "INSERT INTO inventory_item (user_id, number, item_id, amount, is_equipped) VALUES "
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Invent.Object(LoopC).ObjIndex & ", "
            query = query & .Invent.Object(LoopC).Amount & ", "
            query = query & .Invent.Object(LoopC).Equipped & ")"
            If LoopC < MAX_INVENTORY_SLOTS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

        'User bank inventory
        query = "DELETE FROM bank_item WHERE user_id = " & .ID & ";"
        Call Database_Connection.Execute(query)

        query = "INSERT INTO bank_item (user_id, number, item_id, amount) VALUES "
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .BancoInvent.Object(LoopC).ObjIndex & ", "
            query = query & .BancoInvent.Object(LoopC).Amount & ")"
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

        'User skills
        query = "DELETE FROM skillpoint WHERE user_id = " & .ID & ";"
        Call Database_Connection.Execute(query)

        query = "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "
        For LoopC = 1 To NUMSKILLS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserSkills(LoopC) & ", "
            query = query & .Stats.ExpSkills(LoopC) & ", "
            query = query & .Stats.EluSkills(LoopC) & ")"
            If LoopC < NUMSKILLS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

        'User pets
        Dim petType As Integer
        query = "DELETE FROM pet WHERE user_id = " & .ID & ";"
        Call Database_Connection.Execute(query)

        query = "INSERT INTO pet (user_id, number, pet_id) VALUES "
        For LoopC = 1 To MAXMASCOTAS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "

            'CHOTS | I got this logic from SaveUserToCharfile
            If .MascotasIndex(LoopC) > 0 Then
                If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(LoopC)
                Else
                    petType = 0
                End If
            Else
                petType = .MascotasType(LoopC)
            End If

            query = query & petType & ")"
            If LoopC < MAXMASCOTAS Then
                query = query & ", "
            Else
                query = query & ";"
            End If
        Next LoopC
        Call Database_Connection.Execute(query)

    End With

    Call Database_Close

    Exit Sub

ErrorHandler:
        Call LogDatabaseError("Unable to UPDATE User to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)
End Sub

Sub LoadUserFromDatabase(ByVal UserIndex As Integer)
'*************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last modified: 09/10/2018
'Loads the user from the database
'*************************************************

On Error GoTo ErrorHandler

    Dim query As String
    Dim LoopC As Byte

    Call Database_Connect

    'Basic user data
    With UserList(UserIndex)
        query = "SELECT * FROM user WHERE name ='" & UCase$(.Name) & "';"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then Exit Sub

        'Start setting data
        .ID = Database_RecordSet!ID
        .Name = Database_RecordSet!Name
        .Stats.ELV = Database_RecordSet!level
        .Stats.Exp = Database_RecordSet!Exp
        .Stats.ELU = Database_RecordSet!ELU
        .Genero = Database_RecordSet!genre_id
        .raza = Database_RecordSet!race_id
        .clase = Database_RecordSet!class_id
        .Hogar = Database_RecordSet!home_id
        .desc = Database_RecordSet!description
        .Stats.GLD = Database_RecordSet!gold
        .Stats.Banco = Database_RecordSet!bank_gold
        .Stats.SkillPts = Database_RecordSet!free_skillpoints
        .Counters.AsignedSkills = Database_RecordSet!assigned_skillpoints
        .NroMascotas = Database_RecordSet!pet_amount
        .Pos.Map = Database_RecordSet!pos_map
        .Pos.X = Database_RecordSet!pos_x
        .Pos.Y = Database_RecordSet!pos_y
        .flags.lastMap = Database_RecordSet!last_map
        .OrigChar.body = Database_RecordSet!body_id
        .OrigChar.Head = Database_RecordSet!head_id
        .OrigChar.WeaponAnim = Database_RecordSet!weapon_id
        .OrigChar.CascoAnim = Database_RecordSet!helmet_id
        .OrigChar.ShieldAnim = Database_RecordSet!shield_id
        .OrigChar.heading = Database_RecordSet!heading
        .Invent.NroItems = Database_RecordSet!items_amount
        .Invent.ArmourEqpSlot = Database_RecordSet!slot_armour
        .Invent.WeaponEqpSlot = Database_RecordSet!slot_weapon
        .Invent.CascoEqpSlot = Database_RecordSet!slot_helmet
        .Invent.EscudoEqpSlot = Database_RecordSet!slot_shield
        .Invent.MunicionEqpSlot = Database_RecordSet!slot_ammo
        .Invent.BarcoSlot = Database_RecordSet!slot_ship
        .Invent.AnilloEqpSlot = Database_RecordSet!slot_ring
        .Invent.MochilaEqpSlot = Database_RecordSet!slot_bag
        .Stats.MinHp = Database_RecordSet!min_hp
        .Stats.MaxHp = Database_RecordSet!max_hp
        .Stats.MinMAN = Database_RecordSet!min_man
        .Stats.MaxMAN = Database_RecordSet!max_man
        .Stats.MinSta = Database_RecordSet!min_sta
        .Stats.MaxSta = Database_RecordSet!max_sta
        .Stats.MinHam = Database_RecordSet!min_ham
        .Stats.MaxHam = Database_RecordSet!max_ham
        .Stats.MinAGU = Database_RecordSet!min_sed
        .Stats.MaxAGU = Database_RecordSet!max_sed
        .Stats.MinHIT = Database_RecordSet!min_hit
        .Stats.MaxHIT = Database_RecordSet!max_hit
        .Stats.NPCsMuertos = Database_RecordSet!killed_npcs
        .Stats.UsuariosMatados = Database_RecordSet!killed_users
        .Reputacion.AsesinoRep = Database_RecordSet!rep_asesino
        .Reputacion.BandidoRep = Database_RecordSet!rep_bandido
        .Reputacion.BurguesRep = Database_RecordSet!rep_bugues
        .Reputacion.LadronesRep = Database_RecordSet!rep_ladron
        .Reputacion.NobleRep = Database_RecordSet!rep_noble
        .Reputacion.PlebeRep = Database_RecordSet!rep_plebe
        .Reputacion.Promedio = Database_RecordSet!rep_average
        .flags.Desnudo = Database_RecordSet!is_naked
        .flags.Envenenado = Database_RecordSet!is_poisoned
        .flags.Escondido = Database_RecordSet!is_hidden
        .flags.Hambre = Database_RecordSet!is_hungry
        .flags.Sed = Database_RecordSet!is_thirsty
        .flags.Ban = Database_RecordSet!is_ban
        .flags.Muerto = Database_RecordSet!is_dead
        .flags.Navegando = Database_RecordSet!is_sailing
        .flags.Paralizado = Database_RecordSet!is_paralyzed
        .Counters.Pena = Database_RecordSet!counter_pena

        If Database_RecordSet!pertenece_consejo_real Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil
        End If

        If Database_RecordSet!pertenece_consejo_caos Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil
        End If

        .Faccion.ArmadaReal = Database_RecordSet!pertenece_real
        .Faccion.FuerzasCaos = Database_RecordSet!pertenece_caos
        .Faccion.CiudadanosMatados = Database_RecordSet!ciudadanos_matados
        .Faccion.CriminalesMatados = Database_RecordSet!criminales_matados
        .Faccion.RecibioArmaduraReal = Database_RecordSet!recibio_armadura_real
        .Faccion.RecibioArmaduraCaos = Database_RecordSet!recibio_armadura_caos
        .Faccion.RecibioExpInicialReal = Database_RecordSet!recibio_exp_real
        .Faccion.RecibioExpInicialCaos = Database_RecordSet!recibio_exp_caos
        .Faccion.RecompensasReal = Database_RecordSet!recompensas_real
        .Faccion.RecompensasCaos = Database_RecordSet!recompensas_caos
        .Faccion.Reenlistadas = Database_RecordSet!Reenlistadas
        .Faccion.FechaIngreso = Database_RecordSet!fecha_ingreso
        .Faccion.NivelIngreso = Database_RecordSet!nivel_ingreso
        .Faccion.MatadosIngreso = Database_RecordSet!matados_ingreso
        .Faccion.NextRecompensa = Database_RecordSet!siguiente_recompensa

        .GuildIndex = Database_RecordSet!guild_index

        Set Database_RecordSet = Nothing

        'User attributes
        query = "SELECT * FROM attribute WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)
    
        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            While Not Database_RecordSet.EOF
                .Stats.UserAtributos(Database_RecordSet!Number) = Database_RecordSet!value
                .Stats.UserAtributosBackUP(Database_RecordSet!Number) = .Stats.UserAtributos(Database_RecordSet!Number)

                Database_RecordSet.MoveNext
            Wend
        End If

        Set Database_RecordSet = Nothing

        'User spells
        query = "SELECT * FROM spell WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            While Not Database_RecordSet.EOF
                .Stats.UserHechizos(Database_RecordSet!Number) = Database_RecordSet!spell_id

                Database_RecordSet.MoveNext
            Wend
        End If

        Set Database_RecordSet = Nothing

        'User pets
        query = "SELECT * FROM pet WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            While Not Database_RecordSet.EOF
                .MascotasType(Database_RecordSet!Number) = Database_RecordSet!pet_id

                Database_RecordSet.MoveNext
            Wend
        End If

        Set Database_RecordSet = Nothing

        'User inventory
        query = "SELECT * FROM inventory_item WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            While Not Database_RecordSet.EOF
                .Invent.Object(Database_RecordSet!Number).ObjIndex = Database_RecordSet!item_id
                .Invent.Object(Database_RecordSet!Number).Amount = Database_RecordSet!Amount
                .Invent.Object(Database_RecordSet!Number).Equipped = Database_RecordSet!is_equipped

                Database_RecordSet.MoveNext
            Wend
        End If

        Set Database_RecordSet = Nothing

        'User bank inventory
        query = "SELECT * FROM bank_item WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            While Not Database_RecordSet.EOF
                .BancoInvent.Object(Database_RecordSet!Number).ObjIndex = Database_RecordSet!item_id
                .BancoInvent.Object(Database_RecordSet!Number).Amount = Database_RecordSet!Amount

                Database_RecordSet.MoveNext
            Wend
        End If

        Set Database_RecordSet = Nothing

        'User skills
        query = "SELECT * FROM skillpoint WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            While Not Database_RecordSet.EOF
                .Stats.UserSkills(Database_RecordSet!Number) = Database_RecordSet!value
                .Stats.ExpSkills(Database_RecordSet!Number) = Database_RecordSet!Exp
                .Stats.EluSkills(Database_RecordSet!Number) = Database_RecordSet!ELU

                Database_RecordSet.MoveNext
            Wend
        End If

        Set Database_RecordSet = Nothing

    End With

    Exit Sub

    Call Database_Close

ErrorHandler:
        Call LogDatabaseError("Unable to LOAD User from Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)
End Sub

Public Function PersonajeExisteDatabase(ByVal UserName As String) As Boolean
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT id FROM user WHERE name = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        PersonajeExisteDatabase = False
        Exit Function
    End If

    PersonajeExisteDatabase = (Database_RecordSet.RecordCount > 0)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in PersonajeExisteDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function BANCheckDatabase(ByVal UserName As String) As Boolean
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 09/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT is_ban FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        BANCheckDatabase = False
        Exit Function
    End If

    BANCheckDatabase = CBool(Database_RecordSet!is_ban)

    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in BANCheckDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Sub BorrarUsuarioDatabase(ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET name = '" & UCase$(UserName) & "_deleted', deleted = TRUE WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub

ErrorHandler:
        Call LogDatabaseError("Error in BorrarUsuarioDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub UnBanDatabase(ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET is_ban = FALSE WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub

ErrorHandler:
        Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Function GetUserGuildIndexDatabase(ByVal UserName As String) As Integer
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 09/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT guild_index FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildIndexDatabase = 0
        Exit Function
    End If

    GetUserGuildIndexDatabase = CInt(Database_RecordSet!guild_index)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Sub CopyUserDatabase(ByVal UserName As String, ByVal newName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET name = '" & UCase$(newName) & "' WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub

ErrorHandler:
        Call LogDatabaseError("Error in CopyUserDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub MarcarPjComoQueYaVotoDatabase(ByVal UserIndex As Integer, ByVal NumeroEncuesta As Integer)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET votes_amount = " & NumeroEncuesta & " WHERE id = " & UserList(UserIndex).ID & ";"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub

ErrorHandler:
        Call LogDatabaseError("Error in MarcarPjComoQueYaVotoDatabase: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)
End Sub

Public Function PersonajeCantidadVotosDatabase(ByVal UserName As String) As Integer
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT votes_amount FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        PersonajeCantidadVotosDatabase = 0
        Exit Function
    End If

    PersonajeCantidadVotosDatabase = CInt(Database_RecordSet!votes_amount)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in PersonajeCantidadVotosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Sub SaveBanDatabase(ByVal UserName As String, ByVal Reason As String, ByVal BannedBy As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String
    Dim cantPenas As Byte
    cantPenas = GetUserAmountOfPunishmentsDatabase(UserName)

    Call Database_Connect

    query = "UPDATE user SET is_ban = TRUE WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "'), "
    query = query & "number = " & (cantPenas + 1) & ", "
    query = query & "reason = '" & BannedBy & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub

ErrorHandler:
        Call LogDatabaseError("Error in SaveBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Function GetUserAmountOfPunishmentsDatabase(ByVal UserName As String) As Integer
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT COUNT(1) as punishments FROM punishment WHERE user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "')"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserAmountOfPunishmentsDatabase = 0
        Exit Function
    End If

    GetUserAmountOfPunishmentsDatabase = CInt(Database_RecordSet!punishments)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserAmountOfPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Sub SendUserPunishmentsDatabase(ByVal UserIndex As Integer, ByVal UserName As String, ByVal Count As Integer)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT * FROM punishment WHERE user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Not Database_RecordSet.RecordCount = 0 Then
        Database_RecordSet.MoveFirst
        While Not Database_RecordSet.EOF

            Call WriteConsoleMsg(UserIndex, Database_RecordSet!Number & " - " & Database_RecordSet!Reason, FontTypeNames.FONTTYPE_INFO)

            Database_RecordSet.MoveNext
        Wend
    End If

    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SendUserPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Function GetUserPosDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT pos_map, pos_x, pos_y FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserPosDatabase = vbNullString
        Exit Function
    End If

    GetUserPosDatabase = Database_RecordSet!pos_map & "-" & Database_RecordSet!pos_x & "-" & Database_RecordSet!pos_y
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserPosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserSaltDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT salt FROM account WHERE id = (SELECT account_id from user WHERE name = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserSaltDatabase = vbNullString
        Exit Function
    End If

    GetUserSaltDatabase = Database_RecordSet!Salt
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserSaltDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserPasswordDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT password FROM account WHERE id = (SELECT account_id from user WHERE name = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserPasswordDatabase = vbNullString
        Exit Function
    End If

    GetUserPasswordDatabase = Database_RecordSet!Password
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserPasswordDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserEmailDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT username FROM account WHERE id = (SELECT account_id from user WHERE name = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserEmailDatabase = vbNullString
        Exit Function
    End If

    GetUserEmailDatabase = Database_RecordSet!UserName
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserEmailDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Sub StorePasswordSaltDatabase(ByVal UserName As String, ByVal Password As String, ByVal Salt As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE account SET "
    query = query & "password = '" & Password & "', "
    query = query & "salt = '" & Salt & "' "
    query = query & "WHERE user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "');"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in StorePasswordSaltDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SaveUserEmailDatabase(ByVal UserName As String, ByVal Email As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE account SET "
    query = query & "username = '" & Email & "', """
    query = query & "WHERE user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "');"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserEmailDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SaveUserPunishmentDatabase(ByVal UserName As String, ByVal Number As Integer, ByVal Reason As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "'), "
    query = query & "number = " & Number & ", "
    query = query & "reason = '" & Reason & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserPunishmentDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub AlterUserPunishmentDatabase(ByVal UserName As String, ByVal Number As Integer, ByVal Reason As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE punishment SET "
    query = query & "reason = '" & Reason & "' "
    query = query & "WHERE number = " & Number & " AND user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "');"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in AlterUserPunishmentDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub ResetUserFaccionesDatabase(ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "pertenece_real = FALSE, "
    query = query & "pertenece_caos = FALSE, "
    query = query & "ciudadanos_matados = 0, "
    query = query & "criminales_matados = FALSE, "
    query = query & "recibio_armadura_real = FALSE, "
    query = query & "recibio_armadura_caos = FALSE, "
    query = query & "recibio_exp_real = FALSE, "
    query = query & "recibio_exp_caos = FALSE, "
    query = query & "recompensas_real = 0, "
    query = query & "recompensas_caos = 0, "
    query = query & "reenlistadas = 0, "
    query = query & "fecha_ingreso = NULL, "
    query = query & "nivel_ingreso = NULL, "
    query = query & "matados_ingreso = NULL, "
    query = query & "siguiente_recompensa = NULL "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in ResetUserFaccionesDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub KickUserCouncilsDatabase(ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "pertenece_consejo_real = FALSE, "
    query = query & "pertenece_consejo_caos = FALSE "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in KickUserCouncilsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub KickUserFaccionesDatabase(ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "pertenece_real = FALSE, "
    query = query & "pertenece_caos = FALSE "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in KickUserFaccionesDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub KickUserChaosLegionDatabase(ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "pertenece_caos = FALSE, "
    query = query & "reenlistadas = 200 "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in KickUserChaosLegionDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub KickUserRoyalArmyDatabase(ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "pertenece_real = FALSE, "
    query = query & "reenlistadas = 200 "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in KickUserRoyalArmyDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub UpdateUserLoggedDatabase(ByVal UserName As String, ByVal Logged As Byte)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "is_logged = " & IIf(Logged = 1, "TRUE", "FALSE") & " "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in UpdateUserLoggedDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Function GetUserLastIpsDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT last_ip FROM account WHERE id = (SELECT account_id from user WHERE name = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserLastIpsDatabase = vbNullString
        Exit Function
    End If

    GetUserLastIpsDatabase = Database_RecordSet!last_ip
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserLastIpsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserSkillsDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String
    GetUserSkillsDatabase = vbNullString

    Call Database_Connect

    query = "SELECT number, value FROM skillpoint WHERE user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Not Database_RecordSet.RecordCount = 0 Then
        Database_RecordSet.MoveFirst
        While Not Database_RecordSet.EOF
            GetUserSkillsDatabase = GetUserSkillsDatabase & "CHAR>" & SkillsNames(Database_RecordSet!Number) & " = " & Database_RecordSet!value & vbCrLf

            Database_RecordSet.MoveNext
        Wend
    End If

    Set Database_RecordSet = Nothing

    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserSkillsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserFreeSkillsDatabase(ByVal UserName As String) As Integer
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT free_skillpoints FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserFreeSkillsDatabase = vbNullString
        Exit Function
    End If

    GetUserFreeSkillsDatabase = CInt(Database_RecordSet!free_skillpoints)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserFreeSkillsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Sub SaveUserTrainingTimeDatabase(ByVal UserName As String, ByVal trainingTime As Long)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "counter_training = " & trainingTime & " "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserTrainingTimeDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Function GetUserTrainingTimeDatabase(ByVal UserName As String) As Long
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT counter_training FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserTrainingTimeDatabase = 0
        Exit Function
    End If

    GetUserTrainingTimeDatabase = CLng(Database_RecordSet!counter_training)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
        Call LogDatabaseError("Error in GetUserTrainingTimeDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function UserBelongsToRoyalArmyDatabase(ByVal UserName As String) As Boolean
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT pertenece_real FROM user WHERE name = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        UserBelongsToRoyalArmyDatabase = False
        Exit Function
    End If

    UserBelongsToRoyalArmyDatabase = CBool(Database_RecordSet!pertenece_real)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in UserBelongsToRoyalArmyDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function UserBelongsToChaosLegionDatabase(ByVal UserName As String) As Boolean
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT pertenece_caos FROM user WHERE name = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        UserBelongsToChaosLegionDatabase = False
        Exit Function
    End If

    UserBelongsToChaosLegionDatabase = CBool(Database_RecordSet!pertenece_caos)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in UserBelongsToChaosLegionDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserLevelDatabase(ByVal UserName As String) As Byte
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 09/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT level FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserLevelDatabase = 0
        Exit Function
    End If

    GetUserLevelDatabase = CByte(Database_RecordSet!level)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in GetUserLevelDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserPromedioDatabase(ByVal UserName As String) As Long
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 09/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT rep_average FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserPromedioDatabase = 0
        Exit Function
    End If

    GetUserPromedioDatabase = CLng(Database_RecordSet!rep_average)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in GetUserPromedioDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserReenlistsDatabase(ByVal UserName As String) As Byte
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 09/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT reenlistadas FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserReenlistsDatabase = 0
        Exit Function
    End If

    GetUserReenlistsDatabase = CByte(Database_RecordSet!Reenlistadas)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in GetUserReenlistsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Sub SaveUserReenlistsDatabase(ByVal UserName As String, ByVal Reenlists As Byte)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 10/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "reenlistadas = " & Reenlists & " "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserReenlistsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SendUserStatsTxtDatabase(ByVal sendIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    If PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & UserName, FontTypeNames.FONTTYPE_INFO)

        Call Database_Connect

        query = "SELECT level, exp. elu, min_sta, max_sta, min_hp, max_hp, min_man, max_man, min_hit, max_hit, gold FROM user WHERE name = '" & UCase$(UserName) & "';"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Call WriteConsoleMsg(sendIndex, "Nivel: " & Database_RecordSet!level & "  EXP: " & Database_RecordSet!Exp & "/" & Database_RecordSet!ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energía: " & Database_RecordSet!min_sta & "/" & Database_RecordSet!max_sta, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & Database_RecordSet!min_hp & "/" & Database_RecordSet!max_hp, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Mana: " & Database_RecordSet!min_man & "/" & Database_RecordSet!max_man, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Golpe: " & Database_RecordSet!min_hit & "/" & Database_RecordSet!max_hit, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro: " & Database_RecordSet!gold, FontTypeNames.FONTTYPE_INFO)

        Set Database_RecordSet = Nothing
        Call Database_Close
    End If
    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SendUserStatsTxtDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SendUserMiniStatsTxtFromDatabase(ByVal sendIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    If PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & UserName, FontTypeNames.FONTTYPE_INFO)

        Call Database_Connect

        query = "SELECT killed_npcs, killed_users, ciudadanos_matados, criminales_matados, class_id, genre_id, race_id FROM user WHERE name = '" & UCase$(UserName) & "';"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & Database_RecordSet!ciudadanos_matados & ", CriminalesMatados: " & Database_RecordSet!criminales_matados & ", UsuariosMatados: " & Database_RecordSet!killed_users, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & Database_RecordSet!killed_npcs, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(Database_RecordSet!class_id), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Genero: " & IIf(CByte(Database_RecordSet!ciudadanos_matados) = eGenero.Hombre, "Hombre", "Mujer"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Raza: " & ListaRazas(Database_RecordSet!race_id), FontTypeNames.FONTTYPE_INFO)

        Set Database_RecordSet = Nothing
        Call Database_Close
    End If
    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SendUserMiniStatsTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SendUserOROTxtFromDatabase(ByVal sendIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    If PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call Database_Connect

        query = "SELECT bank_gold FROM user WHERE name = '" & UCase$(UserName) & "';"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro en banco: " & Database_RecordSet!bank_gold, FontTypeNames.FONTTYPE_INFO)

        Set Database_RecordSet = Nothing
        Call Database_Close
    End If
    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SendUserOROTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SendUserInvTxtFromDatabase(ByVal sendIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    If PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call Database_Connect

        query = "SELECT number, item_id, amount FROM inventory_item WHERE user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "')"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            While Not Database_RecordSet.EOF
                Call WriteConsoleMsg(sendIndex, "Objeto " & Database_RecordSet!Number & " " & ObjData(Database_RecordSet!item_id).Name & " Cantidad:" & Database_RecordSet!Amount, FontTypeNames.FONTTYPE_INFO)

                Database_RecordSet.MoveNext
            Wend
        Else
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        End If

        Set Database_RecordSet = Nothing
        Call Database_Close
    End If
    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SendUserInvTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SendUserBovedaTxtFromDatabase(ByVal sendIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    If PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call Database_Connect

        query = "SELECT number, item_id, amount FROM bank_item WHERE user_id = (SELECT id from user WHERE name = '" & UCase$(UserName) & "')"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            While Not Database_RecordSet.EOF
                Call WriteConsoleMsg(sendIndex, "Objeto " & Database_RecordSet!Number & " " & ObjData(Database_RecordSet!item_id).Name & " Cantidad:" & Database_RecordSet!Amount, FontTypeNames.FONTTYPE_INFO)

                Database_RecordSet.MoveNext
            Wend
        Else
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        End If

        Set Database_RecordSet = Nothing
        Call Database_Close
    End If
    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SendUserBovedaTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SendCharacterInfoDatabase(ByVal UserIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim gName       As String
    Dim Miembro     As String
    Dim GuildActual As Integer
    Dim query As String

    Call Database_Connect

    query = "SELECT race_id, class_id, genre_id, level, gold, bank_gold, rep_average, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    ' Get the character's current guild
    GuildActual = CInt(Database_RecordSet!guild_index)
    If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
        gName = "<" & GuildName(GuildActual) & ">"
    Else
        gName = "Ninguno"
    End If

    'Get previous guilds
    Miembro = Database_RecordSet!guild_member_history
    If Len(Miembro) > 400 Then
        Miembro = ".." & Right$(Miembro, 400)
    End If

    Call Protocol.WriteCharacterInfo(UserIndex, UserName, Database_RecordSet!race_id, Database_RecordSet!class_id, _
        Database_RecordSet!genre_id, Database_RecordSet!level, Database_RecordSet!gold, _
        Database_RecordSet!bank_gold, Database_RecordSet!rep_average, Database_RecordSet!guild_requests_history, _
        gName, Miembro, Database_RecordSet!pertenece_real, Database_RecordSet!pertenece_caos, _
        Database_RecordSet!ciudadanos_matados, Database_RecordSet!criminales_matados)

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Function GetUserGuildMemberDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT guild_member_history FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildMemberDatabase = vbNullString
        Exit Function
    End If

    GetUserGuildMemberDatabase = Database_RecordSet!guild_member_history
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserGuildAspirantDatabase(ByVal UserName As String) As Integer
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT guild_aspirant_index FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildAspirantDatabase = 0
        Exit Function
    End If

    GetUserGuildAspirantDatabase = CInt(Database_RecordSet!guild_aspirant_index)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserGuildRejectionReasonDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT guild_rejected_because FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildRejectionReasonDatabase = vbNullString
        Exit Function
    End If

    GetUserGuildRejectionReasonDatabase = Database_RecordSet!guild_rejected_because
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Function GetUserGuildPedidosDatabase(ByVal UserName As String) As String
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "SELECT guild_requests_history FROM user WHERE name = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildPedidosDatabase = vbNullString
        Exit Function
    End If

    GetUserGuildPedidosDatabase = Database_RecordSet!guild_requests_history
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function

Public Sub SaveUserGuildRejectionReasonDatabase(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "guild_rejected_because = '" & Reason & "' "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, ByVal GuildIndex As Integer)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "guild_index = " & GuildIndex & " "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, ByVal AspirantIndex As Integer)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "guild_aspirant_index = " & AspirantIndex & " "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "guild_member_history = '" & guilds & "' "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 11/10/2018
'***************************************************
On Error GoTo ErrorHandler
    Dim query As String

    Call Database_Connect

    query = "UPDATE user SET "
    query = query & "guild_requests_history = '" & Pedidos & "' "
    query = query & "WHERE name = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub
ErrorHandler:
        Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Sub
