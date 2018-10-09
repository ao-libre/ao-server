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

    BANCheckDatabase = Database_RecordSet!!is_ban

    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
        Call LogDatabaseError("Error in BANCheckDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
End Function
