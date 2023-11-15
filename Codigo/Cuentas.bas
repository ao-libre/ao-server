Attribute VB_Name = "Cuentas"
Option Explicit

Sub LoadUserFromCharfile(ByVal Userindex As Integer)

    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 14/09/2018
    '14/09/2018: CHOTS - Load the User data from the Charfile using a single function
    '*************************************************
    Dim Leer As clsIniManager

    Set Leer = New clsIniManager

    Call Leer.Initialize(CharPath & UCase$(UserList(Userindex).Name) & ".chr")

    'Cargamos los datos del personaje
    Call LoadUserInit(Userindex, Leer)
    
    'Cargamos las estadisticas del usuario
    Call LoadUserStats(Userindex, Leer)
    
    'Cargamos las estadisticas de las quests
    Call LoadQuestStats(Userindex, Leer)

    Call LoadUserReputacion(Userindex, Leer)

    Set Leer = Nothing

End Sub

Public Function BANCheckCharfile(ByVal UserName As String) As Boolean
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/09/2018
    '***************************************************

    BANCheckCharfile = (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Sub BorrarUsuarioCharfile(ByVal UserName As String)

'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 07/01/2020
'Ahora se pueden borrar los charfiles de la cuenta correctamente (Recox)
'***************************************************
    On Error GoTo ErrorHandler


    If PersonajeExiste(UserName) Then
        UserName = UCase$(UserName)

        Dim AccountHash        As String
        Dim LoopC              As Long
        Dim NumberOfCharacters As Byte
        Dim LastCharacterName  As String
        Dim AccountCharfile    As String
        Dim CurrentCharacter   As String

        AccountHash = GetVar(CharPath & UserName & ".chr", "INIT", "AccountHash")
        AccountCharfile = AccountPath & AccountHash & ".ach"
        NumberOfCharacters = val(GetVar(AccountCharfile, "INIT", "CantidadPersonajes"))

        'Informacion del ultimo pj
        LastCharacterName = GetVar(AccountCharfile, "PERSONAJES", "Personaje" & NumberOfCharacters)


        For LoopC = 1 To NumberOfCharacters
            CurrentCharacter = GetVar(AccountCharfile, "PERSONAJES", "Personaje" & LoopC)

            If UCase$(CurrentCharacter) = UserName Then
                'Movemos el ultimo personaje al slot del borrado
                Call WriteVar(AccountCharfile, "PERSONAJES", "Personaje" & LoopC, LastCharacterName)
                
                'Borramos el nombre del pj de la ultima posicion
                Call WriteVar(AccountCharfile, "PERSONAJES", "Personaje" & NumberOfCharacters, vbNullString)

                'Restamos uno la cantidad de personajes del archivo ach
                Call WriteVar(AccountCharfile, "INIT", "CANTIDADPERSONAJES", NumberOfCharacters - 1)

                'Por ultimo borramos el archivo.
                Kill (CharPath & UCase$(UserName) & ".chr")
                
                Exit Sub
            End If

        Next LoopC
    End If

ErrorHandler:
    Call LogError("Error in BorrarUsuarioCharfile: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function PersonajeExisteCharfile(ByVal UserName As String) As Boolean
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/09/2018
    '***************************************************

    PersonajeExisteCharfile = FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal)

End Function

Public Sub UnBanCharfile(ByVal UserName As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "0")

End Sub

Public Sub SaveBanCharfile(ByVal UserName As String, _
                           ByVal Reason As String, _
                           ByVal BannedBy As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/09/2018
    '***************************************************
    Dim cantPenas As Byte

    cantPenas = GetUserAmountOfPunishmentsCharfile(UserName)

    UserName = UCase$(UserName)
    Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, BannedBy & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)

End Sub

Public Sub CopyUserCharfile(ByVal UserName As String, ByVal newName As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/10/2018
    '***************************************************
    UserName = UCase$(UserName)
    newName = UCase$(newName)

    Dim AccountHash        As String

    Dim LoopC              As Long

    Dim NumberOfCharacters As Byte

    Dim AccountCharfile    As String

    Dim CurrentCharacter   As String

    AccountHash = GetVar(CharPath & UserName & ".chr", "INIT", "AccountHash")
    AccountCharfile = AccountPath & AccountHash & ".ach"
    NumberOfCharacters = val(GetVar(AccountCharfile, "INIT", "CantidadPersonajes"))

    If NumberOfCharacters > 0 Then

        For LoopC = 1 To NumberOfCharacters
            CurrentCharacter = GetVar(AccountCharfile, "PERSONAJES", "Personaje" & LoopC)

            If UCase$(CurrentCharacter) = UserName Then
                Call WriteVar(AccountCharfile, "PERSONAJES", "Personaje" & LoopC, newName)

            End If

        Next LoopC

    End If

    Call FileCopy(CharPath & UserName & ".chr", CharPath & newName & ".chr")

End Sub

Public Function PersonajeCantidadVotosCharfile(ByVal UserName As String) As Integer
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    '***************************************************

    PersonajeCantidadVotosCharfile = val(GetVar(CharPath & UserName & ".chr", "CONSULTAS", "Voto"))

End Function

Public Sub MarcarPjComoQueYaVotoCharfile(ByVal Userindex As Integer, _
                                         ByVal NumeroEncuesta As Integer)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    '***************************************************
    Call WriteVar(CharPath & UserList(Userindex).Name & ".chr", "CONSULTAS", "Voto", str(NumeroEncuesta))

End Sub

Public Function GetUserAmountOfPunishmentsCharfile(ByVal UserName As String) As Integer
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    '***************************************************

    GetUserAmountOfPunishmentsCharfile = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))

End Function

Public Sub SendUserPunishmentsCharfile(ByVal Userindex As Integer, _
                                       ByVal UserName As String, _
                                       ByVal Count As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    '***************************************************
    While Count > 0

        Call WriteConsoleMsg(Userindex, Count & " - " & GetVar(CharPath & UserName & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
        Count = Count - 1
    Wend

End Sub

Public Function GetUserPosCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    '***************************************************

    GetUserPosCharfile = GetVar(CharPath & UserName & ".chr", "INIT", "POSITION")

End Function

Public Function GetUserSaltCharfile(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    '***************************************************
    Dim AccountHash As String

    Dim AccountName As String

    AccountHash = GetVar(CharPath & UserName & ".chr", "INIT", "AccountHash")
    AccountName = GetVar(AccountPath & AccountHash & ".ach", "INIT", "UserName")

    GetUserSaltCharfile = GetVar(AccountPath & AccountName & ".acc", "INIT", "Salt")

End Function

Public Function GetUserPasswordCharfile(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    '***************************************************
    Dim AccountHash As String

    Dim AccountName As String

    AccountHash = GetVar(CharPath & UserName & ".chr", "INIT", "AccountHash")
    AccountName = GetVar(AccountPath & AccountHash & ".ach", "INIT", "UserName")

    GetUserPasswordCharfile = GetVar(AccountPath & AccountName & ".acc", "INIT", "Password")

End Function

Public Function GetAccountSaltCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/10/2018
    '***************************************************

    GetAccountSaltCharfile = GetVar(AccountPath & UserName & ".acc", "INIT", "Salt")

End Function

Public Function GetAccountPasswordCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    '***************************************************

    GetAccountPasswordCharfile = GetVar(AccountPath & UserName & ".acc", "INIT", "Password")

End Function

Public Function GetUserEmailCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    '***************************************************

    GetUserEmailCharfile = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")

End Function

Sub StorePasswordSaltCharfile(ByVal UserName As String, _
                              ByVal Password As String, _
                              ByVal Salt As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    '***************************************************
    Dim AccountHash As String

    Dim AccountName As String

    AccountHash = GetVar(CharPath & UserName & ".chr", "INIT", "AccountHash")
    AccountName = GetVar(AccountPath & AccountHash & ".ach", "INIT", "UserName")

    Call WriteVar(AccountPath & AccountName & ".acc", "INIT", "Password", Password)
    Call WriteVar(AccountPath & AccountName & ".acc", "INIT", "Salt", Salt)

End Sub

Sub SaveUserEmailCharfile(ByVal UserName As String, ByVal Email As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", Email)

End Sub

Sub SaveUserPunishmentCharfile(ByVal UserName As String, _
                               ByVal Number As Integer, _
                               ByVal Reason As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Number)
    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Number, Reason)

End Sub

Sub AlterUserPunishmentCharfile(ByVal UserName As String, _
                                ByVal Number As Integer, _
                                ByVal Reason As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Number, Reason)

End Sub

Sub ResetUserFaccionesCharfile(ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************
    Dim Char As String

    Char = CharPath & UserName & ".chr"

    Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
    Call WriteVar(Char, "FACCIONES", "CiudMatados", 0)
    Call WriteVar(Char, "FACCIONES", "CrimMatados", 0)
    Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
    Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingreso a ninguna Faccion")
    Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
    Call WriteVar(Char, "FACCIONES", "rArReal", 0)
    Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
    Call WriteVar(Char, "FACCIONES", "rExReal", 0)
    Call WriteVar(Char, "FACCIONES", "recCaos", 0)
    Call WriteVar(Char, "FACCIONES", "recReal", 0)
    Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
    Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
    Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
    Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)

End Sub

Sub KickUserCouncilsCharfile(ByVal UserName As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)

End Sub

Sub KickUserFaccionesCharfile(ByVal UserName As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)

End Sub

Sub KickUserChaosLegionCharfile(ByVal UserName As String, ByVal KickerName As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)
    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & KickerName)

End Sub

Sub KickUserRoyalArmyCharfile(ByVal UserName As String, ByVal KickerName As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & KickerName)

End Sub

Sub UpdateUserLoggedCharfile(ByVal UserName As String, ByVal Logged As Byte)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Logged", Logged)

End Sub

Public Function GetUserLastIpsCharfile(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************
    Dim i    As Byte

    Dim list As String

    For i = 1 To 5
        list = list & i & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & i) & vbCrLf
    Next i

    GetUserLastIpsCharfile = list

End Function

Public Function GetUserSkillsCharfile(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************
    Dim i       As Byte

    Dim Message As String

    For i = 1 To NUMSKILLS
        Message = Message & "CHAR>" & SkillsNames(i) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & i) & vbCrLf
    Next i

    GetUserSkillsCharfile = Message

End Function

Public Function GetUserFreeSkillsCharfile(ByVal UserName As String) As Integer
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************

    GetUserFreeSkillsCharfile = val(GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"))

End Function

Public Function GetUserTrainingTimeCharfile(ByVal UserName As String) As Long
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************

    GetUserTrainingTimeCharfile = val(GetVar(CharPath & UserName & ".chr", "RESEARCH", "TrainingTime"))

End Function

Sub SaveUserTrainingTimeCharfile(ByVal UserName As String, ByVal trainingTime As Long)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "RESEARCH", "TrainingTime", trainingTime)

End Sub

Public Function GetUserGuildIndexCharfile(ByRef UserName As String) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/09/2018
    '26/09/2018 CHOTS: Moved to FileIO
    '***************************************************
    Dim Temps As String
    
    Temps = GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX")

    If IsNumeric(Temps) Then
        GetUserGuildIndexCharfile = CInt(Temps)
    Else
        GetUserGuildIndexCharfile = 0

    End If

End Function

Public Function GetUserGuildMemberCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserGuildMemberCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "Miembro")

End Function

Public Function GetUserGuildAspirantCharfile(ByVal UserName As String) As Integer
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserGuildAspirantCharfile = val(GetVar(CharPath & UserName & ".chr", "GUILD", "ASPIRANTEA"))

End Function

Public Function GetUserGuildRejectionReasonCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserGuildRejectionReasonCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "MotivoRechazo")

End Function

Sub SaveUserGuildRejectionReasonCharfile(ByVal UserName As String, ByVal Reason As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "MotivoRechazo", Reason)

End Sub

Public Function UserBelongsToRoyalArmyCharfile(ByVal UserName As String) As Boolean
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    UserBelongsToRoyalArmyCharfile = CByte(GetVar(CharPath & UserName & ".chr", "Facciones", "EjercitoReal")) = 1

End Function

Public Function UserBelongsToChaosLegionCharfile(ByVal UserName As String) As Boolean
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    UserBelongsToChaosLegionCharfile = CByte(GetVar(CharPath & UserName & ".chr", "Facciones", "EjercitoCaos")) = 1

End Function

Public Function GetUserLevelCharfile(ByVal UserName As String) As Byte
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserLevelCharfile = val(GetVar(CharPath & UserName & ".chr", "Stats", "ELV"))

End Function

Public Function GetUserPromedioCharfile(ByVal UserName As String) As Long
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserPromedioCharfile = val(GetVar(CharPath & UserName & ".chr", "REP", "Promedio"))

End Function

Public Function GetUserReenlistsCharfile(ByVal UserName As String) As Byte
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserReenlistsCharfile = val(GetVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas"))

End Function

Sub SaveUserReenlistsCharfile(ByVal UserName As String, ByVal Reenlists As Byte)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", Reenlists)

End Sub

Public Function GetUserGuildPedidosCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserGuildPedidosCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "Pedidos")

End Function

Sub SaveUserGuildPedidosCharfile(ByVal UserName As String, ByVal Pedidos As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "Pedidos", Pedidos)

End Sub

Sub SaveUserGuildMemberCharfile(ByVal UserName As String, ByVal guilds As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "Miembro", guilds)

End Sub

Sub SaveUserGuildIndexCharfile(ByVal UserName As String, ByVal GuildIndex As Integer)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX", GuildIndex)

End Sub

Sub SaveUserGuildAspirantCharfile(ByVal UserName As String, _
                                  ByVal AspirantIndex As Integer)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "ASPIRANTEA", AspirantIndex)

End Sub

Sub SendCharacterInfoCharfile(ByVal Userindex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************
    Dim gName       As String

    Dim UserFile    As clsIniManager

    Dim Miembro     As String

    Dim GuildActual As Integer

    ' Get the character's current guild
    GuildActual = GetUserGuildIndex(UserName)

    If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
        gName = "<" & GuildName(GuildActual) & ">"
    Else
        gName = "Ninguno"

    End If
    
    'Get previous guilds
    Miembro = GetUserGuildMember(UserName)

    If Len(Miembro) > 400 Then
        Miembro = ".." & Right$(Miembro, 400)

    End If

    Set UserFile = New clsIniManager

    With UserFile
        .Initialize (CharPath & UserName & ".chr")
    
        Call Protocol.WriteCharacterInfo(Userindex, UserName, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), .GetValue("STATS", "Banco"), .GetValue("REP", "Promedio"), .GetValue("GUILD", "Pedidos"), gName, Miembro, .GetValue("FACCIONES", "EjercitoReal"), .GetValue("FACCIONES", "EjercitoCaos"), .GetValue("FACCIONES", "CiudMatados"), .GetValue("FACCIONES", "CrimMatados"))

    End With
    
    Set UserFile = Nothing

End Sub

Public Sub SaveNewAccountCharfile(ByVal UserName As String, _
                                  ByVal Password As String, _
                                  ByVal Salt As String, _
                                  ByVal Hash As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Manager     As clsIniManager

    Dim AccountFile As String

    'CHOTS | First the account itself
    Set Manager = New clsIniManager
    AccountFile = AccountPath & UCase$(UserName) & ".acc"

    With Manager

        Call .ChangeValue("INIT", "Password", Password)
        Call .ChangeValue("INIT", "Salt", Salt)
        Call .ChangeValue("INIT", "Hash", Hash)
        Call .ChangeValue("INIT", "FechaCreado", Date & " " & time)

        Call .DumpFile(AccountFile)

    End With

    Set Manager = Nothing

    'CHOTS | Now the account char files
    Set Manager = New clsIniManager
    AccountFile = AccountPath & Hash & ".ach"

    With Manager
        Call .ChangeValue("INIT", "UserName", UCase$(UserName))
        Call .ChangeValue("INIT", "CantidadPersonajes", 0)

        .DumpFile (AccountFile)

    End With

    Set Manager = Nothing

    Exit Sub
ErrorHandler:
    Call LogError("Error in SaveNewAccountCharfile: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function CuentaExisteCharfile(ByVal UserName As String) As Boolean
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/10/2018
    '***************************************************

    CuentaExisteCharfile = FileExist(AccountPath & UCase$(UserName) & ".acc", vbNormal)

End Function

Public Function PersonajePerteneceCuentaCharfile(ByVal UserName As String, _
                                                 ByVal AccountHash As String) As Boolean

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/10/2018
    '***************************************************
    Dim CharfileHash As String

    CharfileHash = GetVar(CharPath & UserName & ".chr", "INIT", "AccountHash")

    PersonajePerteneceCuentaCharfile = (AccountHash = CharfileHash)

End Function

Public Sub SaveUserToAccountCharfile(ByVal UserName As String, _
                                     ByVal AccountHash As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/10/2018
    '***************************************************
    Dim CantidadPersonajes As Byte

    Dim AccountCharfile    As String

    AccountCharfile = AccountPath & AccountHash & ".ach"

    If FileExist(AccountCharfile) Then
        CantidadPersonajes = val(GetVar(AccountCharfile, "INIT", "CantidadPersonajes"))
        CantidadPersonajes = CantidadPersonajes + 1

        If CantidadPersonajes <= 10 Then
            Call WriteVar(AccountCharfile, "INIT", "CantidadPersonajes", CantidadPersonajes)
            Call WriteVar(AccountCharfile, "PERSONAJES", "Personaje" & CantidadPersonajes, UserName)
        Else
            Call LogError("Error in SaveUserToAccountCharfile. Se intento crear mas de 10 personajes. Username: " & UserName & ". Hash: " & AccountHash)

        End If

    Else
        Call LogError("Error in SaveUserToAccountCharfile. Cuenta inexistente de " & UserName & ". Hash: " & AccountHash)

    End If

End Sub

Public Sub LoginAccountCharfile(ByVal Userindex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Account            As clsIniManager
    Dim CharFile           As clsIniManager
    Dim i                  As Long
    Dim AccountHash        As String
    Dim NumberOfCharacters As Byte
    Dim Characters()       As AccountUser
    Dim CurrentCharacter   As String

    Set Account = New clsIniManager
    Set CharFile = New clsIniManager

    AccountHash = GetVar(AccountPath & UCase$(UserName) & ".acc", "INIT", "Hash")
    Call Account.Initialize(AccountPath & AccountHash & ".ach")
    NumberOfCharacters = val(Account.GetValue("INIT", "CantidadPersonajes"))

    If NumberOfCharacters > 0 Then
        ReDim Characters(1 To NumberOfCharacters) As AccountUser

        For i = 1 To NumberOfCharacters
            CurrentCharacter = Account.GetValue("PERSONAJES", "Personaje" & i)

            Call CharFile.Initialize(CharPath & CurrentCharacter & ".chr")

            Characters(i).Name = CurrentCharacter
            Characters(i).body = val(CharFile.GetValue("INIT", "Body"))
            Characters(i).Head = val(CharFile.GetValue("INIT", "Head"))
            Characters(i).weapon = val(CharFile.GetValue("INIT", "Arma"))
            Characters(i).shield = val(CharFile.GetValue("INIT", "Escudo"))
            Characters(i).helmet = val(CharFile.GetValue("INIT", "Casco"))
            Characters(i).Class = val(CharFile.GetValue("INIT", "Clase"))
            Characters(i).race = val(CharFile.GetValue("INIT", "Raza"))
            Characters(i).Map = val(ReadField(1, CharFile.GetValue("INIT", "Position"), 45))
            Characters(i).level = val(CharFile.GetValue("STATS", "ELV"))
            Characters(i).Gold = val(CharFile.GetValue("STATS", "GLD"))
            Characters(i).criminal = (val(CharFile.GetValue("REP", "Promedio")) < 0)
            Characters(i).dead = CBool(val(CharFile.GetValue("FLAGS", "Muerto")))
            Characters(i).gameMaster = EsGmChar(CurrentCharacter)
        Next i

    End If

    Set Account = Nothing
    Set CharFile = Nothing

    Call WriteUserAccountLogged(Userindex, UserName, AccountHash, NumberOfCharacters, Characters)

    Call SaveLastIpsAccountCharfile(UserName, UserList(Userindex).ip)

    Exit Sub
ErrorHandler:
    Call LogError("Error in LoginAccountCharfile: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveBan(ByVal UserName As String, _
                   ByVal Reason As String, _
                   ByVal BannedBy As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/09/2018
    'Saves the ban flag and reason
    '***************************************************
    If Not Database_Enabled Then
        Call SaveBanCharfile(UserName, Reason, BannedBy)
    Else
        Call SaveBanDatabase(UserName, Reason, BannedBy)

    End If

End Sub

Public Function GetUserAmountOfPunishments(ByVal UserName As String) As Integer

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    'Get the user number of punishments
    '***************************************************
    If Not Database_Enabled Then
        GetUserAmountOfPunishments = GetUserAmountOfPunishmentsCharfile(UserName)
    Else
        GetUserAmountOfPunishments = GetUserAmountOfPunishmentsDatabase(UserName)

    End If

End Function

Public Sub SendUserPunishments(ByVal Userindex As Integer, _
                               ByVal UserName As String, _
                               ByVal Count As Integer)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/09/2018
    'Writes a console msg for each punishment
    '***************************************************
    If Not Database_Enabled Then
        Call SendUserPunishmentsCharfile(Userindex, UserName, Count)
    'Else
       ' Call SendUserPunishmentsDatabase(Userindex, UserName, Count)

    End If

End Sub

Public Function GetUserPos(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    'Get the user position
    '***************************************************
    If Not Database_Enabled Then
        GetUserPos = GetUserPosCharfile(UserName)
    Else
        GetUserPos = GetUserPosDatabase(UserName)

    End If

End Function

Public Function GetAccountSalt(ByVal AccountName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Password Salt
    '***************************************************
    If Not Database_Enabled Then
        GetAccountSalt = GetAccountSaltCharfile(AccountName)
    Else
        GetAccountSalt = GetAccountSaltDatabase(AccountName)

    End If

End Function

Public Function GetUserSalt(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Password Salt
    '***************************************************
    If Not Database_Enabled Then
        GetUserSalt = GetUserSaltCharfile(UserName)
    Else
        GetUserSalt = GetUserSaltDatabase(UserName)

    End If

End Function

Public Function GetAccountPassword(ByVal AccountName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Password
    '***************************************************
    If Not Database_Enabled Then
        GetAccountPassword = GetAccountPasswordCharfile(AccountName)
    Else
        GetAccountPassword = GetAccountPasswordDatabase(AccountName)

    End If

End Function

Public Function GetUserPassword(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Password
    '***************************************************
    If Not Database_Enabled Then
        GetUserPassword = GetUserPasswordCharfile(UserName)
    Else
        GetUserPassword = GetUserPasswordDatabase(UserName)

    End If

End Function

Public Function GetUserEmail(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Email
    '***************************************************
    If Not Database_Enabled Then
        GetUserEmail = GetUserEmailCharfile(UserName)
    Else
        GetUserEmail = GetUserEmailDatabase(UserName)

    End If

End Function

Public Sub StorePasswordSalt(ByVal UserName As String, _
                             ByVal Password As String, _
                             ByVal Salt As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves the password and salt
    '***************************************************
    If Not Database_Enabled Then
        Call StorePasswordSaltCharfile(UserName, Password, Salt)
    Else
        Call StorePasswordSaltDatabase(UserName, Password, Salt)

    End If

End Sub

Public Sub SaveUserEmail(ByVal UserName As String, ByVal Email As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves the email
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserEmailCharfile(UserName, Email)
    Else
        Call SaveUserEmailDatabase(UserName, Email)

    End If

End Sub

Public Sub SaveUserPunishment(ByVal UserName As String, _
                              ByVal Number As Integer, _
                              ByVal Reason As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves a new punishment
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserPunishmentCharfile(UserName, Number, Reason)
    Else
        Call SaveUserPunishmentDatabase(UserName, Number, Reason)

    End If

End Sub

Public Sub AlterUserPunishment(ByVal UserName As String, _
                               ByVal Number As Integer, _
                               ByVal Reason As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves a new punishment
    '***************************************************
    If Not Database_Enabled Then
        Call AlterUserPunishmentCharfile(UserName, Number, Reason)
    Else
        Call AlterUserPunishmentDatabase(UserName, Number, Reason)

    End If

End Sub

Public Sub ResetUserFacciones(ByVal UserName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Reset the imperial an legionary armies
    '***************************************************
    If Not Database_Enabled Then
        Call ResetUserFaccionesCharfile(UserName)
    Else
        Call ResetUserFaccionesDatabase(UserName)

    End If

End Sub

Public Sub KickUserCouncils(ByVal UserName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Kicks the user from both councils
    '***************************************************
    If Not Database_Enabled Then
        Call KickUserCouncilsCharfile(UserName)
    Else
        Call KickUserCouncilsDatabase(UserName)

    End If

End Sub

Public Sub KickUserFacciones(ByVal UserName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Kicks the user from both factions
    '***************************************************
    If Not Database_Enabled Then
        Call KickUserFaccionesCharfile(UserName)
    Else
        Call KickUserFaccionesDatabase(UserName)

    End If

End Sub

Public Sub KickUserChaosLegion(ByVal UserName As String, ByVal KickerName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Kicks the user from ChaosLegion
    '***************************************************
    If Not Database_Enabled Then
        Call KickUserChaosLegionCharfile(UserName, KickerName)
    Else
        Call KickUserChaosLegionDatabase(UserName)

    End If

End Sub

Public Sub KickUserRoyalArmy(ByVal UserName As String, ByVal KickerName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Kicks the user from RoyalArmy
    '***************************************************
    If Not Database_Enabled Then
        Call KickUserRoyalArmyCharfile(UserName, KickerName)
    Else
        Call KickUserRoyalArmyDatabase(UserName)

    End If

End Sub

Public Sub UpdateUserLogged(ByVal UserName As String, ByVal Logged As Byte)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Updates the logged value for the user
    '***************************************************
    If Not Database_Enabled Then
        Call UpdateUserLoggedCharfile(UserName, Logged)
    Else
        Call UpdateUserLoggedDatabase(UserName, Logged)

    End If

End Sub

Public Function GetUserLastIps(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Last IPs list
    '***************************************************
    If Not Database_Enabled Then
        GetUserLastIps = GetUserLastIpsCharfile(UserName)
    Else
        GetUserLastIps = GetUserLastIpsDatabase(UserName)

    End If

End Function

Public Function GetUserSkills(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Skills list
    '***************************************************
    If Not Database_Enabled Then
        GetUserSkills = GetUserSkillsCharfile(UserName)
    Else
        GetUserSkills = GetUserSkillsDatabase(UserName)

    End If

End Function

Public Function GetUserFreeSkills(ByVal UserName As String) As Integer

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Get the number of free skillspoints
    '***************************************************
    If Not Database_Enabled Then
        GetUserFreeSkills = GetUserFreeSkillsCharfile(UserName)
    Else
        GetUserFreeSkills = GetUserFreeSkillsDatabase(UserName)

    End If

End Function

Public Sub SaveUserTrainingTime(ByVal UserName As String, ByVal trainingTime As Long)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Updates the trainingTime value for the user
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserTrainingTimeCharfile(UserName, trainingTime)
    Else
        Call SaveUserTrainingTimeDatabase(UserName, trainingTime)

    End If

End Sub

Public Function GetUserTrainingTime(ByVal UserName As String) As Long

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Get the training time in minutes
    '***************************************************
    If Not Database_Enabled Then
        GetUserTrainingTime = GetUserTrainingTimeCharfile(UserName)
    Else
        GetUserTrainingTime = GetUserTrainingTimeDatabase(UserName)

    End If

End Function

Public Function UserBelongsToRoyalArmy(ByVal UserName As String) As Boolean

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Check if the user belongs to Royal Army
    '***************************************************
    If Not Database_Enabled Then
        UserBelongsToRoyalArmy = UserBelongsToRoyalArmyCharfile(UserName)
    Else
        UserBelongsToRoyalArmy = UserBelongsToRoyalArmyDatabase(UserName)

    End If

End Function

Public Function UserBelongsToChaosLegion(ByVal UserName As String) As Boolean

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Check if the user belongs to Chaos Legion
    '***************************************************
    If Not Database_Enabled Then
        UserBelongsToChaosLegion = UserBelongsToChaosLegionCharfile(UserName)
    Else
        UserBelongsToChaosLegion = UserBelongsToChaosLegionDatabase(UserName)

    End If

End Function

Public Function GetUserLevel(ByVal UserName As String) As Byte

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Get the User Level
    '***************************************************
    If Not Database_Enabled Then
        GetUserLevel = GetUserLevelCharfile(UserName)
    Else
        GetUserLevel = GetUserLevelDatabase(UserName)

    End If

End Function

Public Function GetUserPromedio(ByVal UserName As String) As Long

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Get the User Reputation Average
    '***************************************************
    If Not Database_Enabled Then
        GetUserPromedio = GetUserPromedioCharfile(UserName)
    Else
        GetUserPromedio = GetUserPromedioDatabase(UserName)

    End If

End Function

Public Function GetUserReenlists(ByVal UserName As String) As Byte

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Get the User Legion reenlists
    '***************************************************
    If Not Database_Enabled Then
        GetUserReenlists = GetUserReenlistsCharfile(UserName)
    Else
        GetUserReenlists = GetUserReenlistsDatabase(UserName)

    End If

End Function

Public Sub SaveUserReenlists(ByVal UserName As String, ByVal Reenlists As Byte)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the number of reenlists
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserReenlistsCharfile(UserName, Reenlists)
    Else
        Call SaveUserReenlistsDatabase(UserName, Reenlists)

    End If

End Sub

Public Sub SaveNewAccount(ByVal UserName As String, _
                          ByVal Password As String, _
                          ByVal Salt As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    'Saves a new account
    '***************************************************
    Dim Hash As String

    Hash = RandomString(32)

    If Not Database_Enabled Then
        Call SaveNewAccountCharfile(UserName, Password, Salt, Hash)
    Else
        Call SaveNewAccountDatabase(UserName, Password, Salt, Hash)

    End If

End Sub

Public Function GetLastIpsAccount(ByVal UserName As String) As String

    '***************************************************
    'Autor: Lucas Recoaro (RecoX)
    'Last Modification: 15/07/2020
    'Get the account Last IPs list
    '***************************************************
    If Not Database_Enabled Then
        GetLastIpsAccount = GetLastIpsAccountCharfile(UserName)
    Else
        'TODO: Obtain this from MYSQL
        GetLastIpsAccount = "TODO: Obtain this from MYSQL"
    End If

End Function

Private Function GetLastIpsAccountCharfile(ByVal UserName As String) As String

    '***************************************************
    'Author: Lucas Recoaro (Recox)
    'Last Modification: 15/07/2020
    '***************************************************
    Dim i    As Byte

    Dim list As String

    For i = 1 To 5
        list = list & i & " - " & GetVar(AccountPath & UserName & ".acc", "INIT", "LastIP" & i) & vbCrLf
    Next i

    GetLastIpsAccountCharfile = list

End Function

Public Sub SaveLastIpsAccountCharfile(ByVal UserName As String, ByVal CurrentIp As String)

    '***************************************************
    'Author: Lucas Recoaro (Recox)
    'Last Modification: 15/07/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Manager     As clsIniManager
    Dim AccountFile As String

    Set Manager = New clsIniManager
    AccountFile = AccountPath & UCase$(UserName) & ".acc"

    Call Manager.Initialize(AccountFile)

    'Copio y pego logica de SaveUserToCharfile
    'First time around?
    If Manager.GetValue("INIT", "LastIP1") = vbNullString Then
        Call Manager.ChangeValue("INIT", "LastIP1", CurrentIp & " - " & Date & ":" & time)
        'Is it a different ip from last time?
    ElseIf CurrentIp <> Left$(Manager.GetValue("INIT", "LastIP1"), InStr(1, Manager.GetValue("INIT", "LastIP1"), " ") - 1) Then

        Dim i As Integer

        For i = 5 To 2 Step -1
            Call Manager.ChangeValue("INIT", "LastIP" & i, Manager.GetValue("INIT", "LastIP" & CStr(i - 1)))
        Next i

        Call Manager.ChangeValue("INIT", "LastIP1", CurrentIp & " - " & Date & ":" & time)
        'Same ip, just update the date
    Else
        Call Manager.ChangeValue("INIT", "LastIP1", CurrentIp & " - " & Date & ":" & time)

    End If

    Call Manager.DumpFile(AccountFile)

    Set Manager = Nothing

    Exit Sub
ErrorHandler:
    Call LogError("Error in SaveLastIpsAccountCharfile: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub
