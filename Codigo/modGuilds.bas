Attribute VB_Name = "modGuilds"
Option Explicit

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'º¬

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE   As String
'archivo .\guilds\guildinfo.ini o similar

Private Const ORDENARLISTADECLANES = True
'True si se envia la lista ordenada por alineacion

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

Public Guilds()         As clsClan
'array global de guilds, se indexa por userlist().guildindex

Public Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Public Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Public Const MAXANTIFACCION As Byte = 5
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

Public GMsEscuchando As New Collection

Public Enum ALINEACION_GUILD
    ALINEACION_LEGION = 1
    ALINEACION_CRIMINAL = 2
    ALINEACION_NEUTRO = 3
    ALINEACION_CIUDA = 4
    ALINEACION_ARMADA = 5
    ALINEACION_MASTER = 6
End Enum
'alineaciones permitidas

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum
'numero de .wav del cliente

Public Enum RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum
'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadGuildsDB()

Dim CantClanes  As String
Dim i           As Integer
Dim TempStr     As String
Dim Alin        As ALINEACION_GUILD
    
    GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"

    CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    
    If IsNumeric(CantClanes) Then
        CANTIDADDECLANES = CInt(CantClanes)
    Else
        CANTIDADDECLANES = 0
    End If
    
    For i = 1 To CANTIDADDECLANES
        Set Guilds(i) = New clsClan
        TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
        Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
        Call Guilds(i).Inicializar(TempStr, i, Alin)
    Next i
    
End Sub

Public Function m_ConectarMiembroAClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
Dim NuevoL  As Boolean
Dim NuevaA  As Boolean
Dim News    As String

    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        Call Guilds(GuildIndex).ConectarMiembro(UserIndex)
        UserList(UserIndex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(UserIndex, True, NuevaA, NuevoL)
        If NuevoL Then News = "El clan tiene nuevo líder."
        If NuevaA Then News = News & "El clan tiene nueva alineación."
        If NuevoL Or NuevaA Then Call Guilds(GuildIndex).SetGuildNews(News)
    End If

End Function


Public Function m_ValidarPermanencia(ByVal UserIndex As Integer, ByVal SumaAntifaccion As Boolean, ByRef CambioAlineacion As Boolean, ByRef CambioLider As Boolean) As Boolean
Dim GuildIndex  As Integer
Dim ML          As String
Dim M           As String
Dim UI          As Integer
Dim Sale        As Boolean
Dim i           As Integer

    m_ValidarPermanencia = True
    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    
    If Not m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
    
        Call LogClanes(UserList(UserIndex).name & " de " & Guilds(GuildIndex).GuildName & " es expulsado en validar permanencia")
    
        m_ValidarPermanencia = False
        If SumaAntifaccion Then Guilds(GuildIndex).PuntosAntifaccion = Guilds(GuildIndex).PuntosAntifaccion + 1
        
        CambioAlineacion = (m_EsGuildFounder(UserList(UserIndex).name, GuildIndex) Or Guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION)
        
        Call LogClanes(UserList(UserIndex).name & " de " & Guilds(GuildIndex).GuildName & IIf(CambioAlineacion, " SI ", " NO ") & "provoca cambio de alinaecion. MAXANT:" & (Guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION) & ", GUILDFOU:" & m_EsGuildFounder(UserList(UserIndex).name, GuildIndex))
        
        If CambioAlineacion Then
            'aca tenemos un problema, el fundador acaba de cambiar el rumbo del clan o nos zarpamos de antifacciones
            'Tenemos que resetear el lider, revisar si el lider permanece y si no asignarle liderazgo al fundador

            Call Guilds(GuildIndex).CambiarAlineacion(ALINEACION_NEUTRO)
            Guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
            'para la nueva alineacion, hay que revisar a todos los Pjs!

            'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo alterar
            'internamente al iterador en el proceso
            CambioLider = False
            i = 1
            ML = Guilds(GuildIndex).GetMemberList(",")
            M = ReadField(i, ML, Asc(","))
            While M <> vbNullString

                'vamos a violar un poco de capas..
                UI = NameIndex(M)
                If UI > 0 Then
                    Sale = Not m_EstadoPermiteEntrar(UI, GuildIndex)
                Else
                    Sale = Not m_EstadoPermiteEntrarChar(M, GuildIndex)
                End If

                If Sale Then
                    If m_EsGuildFounder(M, GuildIndex) Then 'hay que sacarlo de las armadas
                        If UI > 0 Then
                            UserList(UI).Faccion.FuerzasCaos = 0
                            UserList(UI).Faccion.ArmadaReal = 0
                            UserList(UI).Faccion.Reenlistadas = 200
                        Else
                            If FileExist(CharPath & M & ".chr") Then
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "EjercitoCaos", 0)
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "ArmadaReal", 0)
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "Reenlistadas", 200)
                            End If
                        End If
                        m_ValidarPermanencia = True
                    Else    'sale si no es guildfounder
                        If m_EsGuildLeader(M, GuildIndex) Then
                            'pierde el liderazgo
                            CambioLider = True
                            Call Guilds(GuildIndex).SetLeader(Guilds(GuildIndex).Fundador)
                        End If

                        Call m_EcharMiembroDeClan(-1, M)
                    End If
                End If
                i = i + 1
                M = ReadField(i, ML, Asc(","))
            Wend
        Else
            'no se va el fundador, el peor caso es que se vaya el lider
            
            'If m_EsGuildLeader(UserList(UserIndex).Name, GuildIndex) Then
            '    Call LogClanes("Se transfiere el liderazgo de: " & Guilds(GuildIndex).GuildName & " a " & Guilds(GuildIndex).Fundador)
            '    Call Guilds(GuildIndex).SetLeader(Guilds(GuildIndex).Fundador)  'transferimos el lideraztgo
            'End If
            Call m_EcharMiembroDeClan(-1, UserList(UserIndex).name)   'y lo echamos
        End If
    End If
    

End Function

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
    If UserList(UserIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call Guilds(GuildIndex).DesConectarMiembro(UserIndex)
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(Guilds(GuildIndex).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(Guilds(GuildIndex).Fundador)))
End Function


'Public Function GetLeader(ByVal GuildIndex As Integer) As String
'    GetLeader = vbNullString
'
'    If GuildIndex <= 0 Then Exit Function
'    GetLeader = Guilds(GuildIndex).GetLeader()
'End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
'UI echa a Expulsado del clan de Expulsado
Dim UserIndex   As Integer
Dim GI          As Integer
    
    m_EcharMiembroDeClan = 0

    UserIndex = NameIndex(Expulsado)
    If UserIndex > 0 Then
        'pj online
        GI = UserList(UserIndex).GuildIndex
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then Guilds(GI).SetLeader (Guilds(GI).Fundador)
                Call Guilds(GI).DesConectarMiembro(UserIndex)
                Call Guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & Guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(UserIndex).GuildIndex = 0
                Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    Else
        'pj offline
        GI = GetGuildIndexFromChar(Expulsado)
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then Guilds(GI).SetLeader (Guilds(GI).Fundador)
                Call Guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & Guilds(GI).GuildName & " Expulsador = " & Expulsador)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    End If

End Function

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
Dim GI As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then Exit Sub
    
    Call Guilds(GI).SetURL(Web)
    
End Sub


Public Sub ActualizarCodexYDesc(ByRef Datos As String, ByVal GuildIndex As Integer)
Dim CantCodex       As Integer
Dim i               As Integer

    If GuildIndex = 0 Then Exit Sub
    Call Guilds(GuildIndex).SetDesc(ReadField(1, Datos, Asc("¬")))
    CantCodex = CInt(ReadField(2, Datos, Asc("¬")))
    For i = 1 To CantCodex
        Call Guilds(GuildIndex).SetCodex(i, ReadField(2 + i, Datos, Asc("¬")))
    Next i
    For i = CantCodex + 1 To CANTIDADMAXIMACODEX
        Call Guilds(GuildIndex).SetCodex(i, vbNullString)
    Next i

End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
Dim GI              As Integer

    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then Exit Sub
    
    Call Guilds(GI).SetGuildNews(Datos)
        
End Sub

Public Function CrearNuevoClan(ByRef GuildInfo As String, ByVal FundadorIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
Dim GuildName       As String
Dim Descripcion     As String
Dim URL             As String
Dim codex()         As String
Dim CantCodex       As Integer
Dim i               As Integer
Dim DummyString     As String

    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If

    GuildName = Trim$(ReadField(2, GuildInfo, Asc("¬")))

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan inválido."
        Exit Function
    End If
    
    If YaExiste(GuildName) Then
        refError = "Ya existe un clan con ese nombre."
        Exit Function
    End If

    Descripcion = ReadField(1, GuildInfo, Asc("¬"))
    URL = ReadField(3, GuildInfo, Asc("¬"))
    CantCodex = CInt(ReadField(4, GuildInfo, Asc("¬")))

    If CantCodex > 0 Then
        ReDim codex(1 To CantCodex) As String
        For i = 1 To CantCodex
            codex(i) = ReadField(4 + i, GuildInfo, Asc("¬"))
        Next i
    End If

    'tenemos todo para fundar ya
    If CANTIDADDECLANES < UBound(Guilds) Then
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan

        'constructor custom de la clase clan
        Set Guilds(CANTIDADDECLANES) = New clsClan
        Call Guilds(CANTIDADDECLANES).Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
        
        'Damos de alta al clan como nuevo inicializando sus archivos
        Call Guilds(CANTIDADDECLANES).InicializarNuevoClan(UserList(FundadorIndex).name)
        
        'seteamos codex y descripcion
        For i = 1 To CantCodex
            Call Guilds(CANTIDADDECLANES).SetCodex(i, codex(i))
        Next i
        Call Guilds(CANTIDADDECLANES).SetDesc(Descripcion)
        Call Guilds(CANTIDADDECLANES).SetGuildNews("Clan creado con alineación : " & Alineacion2String(Alineacion))
        Call Guilds(CANTIDADDECLANES).SetLeader(UserList(FundadorIndex).name)
        Call Guilds(CANTIDADDECLANES).SetURL(URL)
        
        '"conectamos" al nuevo miembro a la lista de la clase
        Call Guilds(CANTIDADDECLANES).AceptarNuevoMiembro(UserList(FundadorIndex).name)
        Call Guilds(CANTIDADDECLANES).ConectarMiembro(FundadorIndex)
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call WarpUserChar(FundadorIndex, UserList(FundadorIndex).Pos.Map, UserList(FundadorIndex).Pos.X, UserList(FundadorIndex).Pos.Y, False)
        
        For i = 1 To CANTIDADDECLANES - 1
            Call Guilds(i).ProcesarFundacionDeOtroClan
        Next i
    Else
        refError = "No hay mas slots para fundar clanes. Consulte a un administrador."
        Exit Function
    End If
    
    CrearNuevoClan = True
    
End Function


Public Sub SendGuildNews(ByVal UserIndex As Integer)
Dim News            As String
Dim EnemiesCount    As Integer
Dim AlliesCount     As Integer
Dim GuildIndex      As Integer
Dim i               As Integer


    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex = 0 Then Exit Sub
    
    News = "GUILDNE" & Guilds(GuildIndex).GetGuildNews & "¬"

    EnemiesCount = Guilds(GuildIndex).CantidadEnemys
    News = News & CStr(EnemiesCount) & "¬"
    i = Guilds(GuildIndex).Iterador_ProximaRelacion(GUERRA)
    While i > 0
        News = News & Guilds(i).GuildName & "¬"
        i = Guilds(GuildIndex).Iterador_ProximaRelacion(GUERRA)
    Wend
    AlliesCount = Guilds(GuildIndex).CantidadAllies
    News = News & CStr(AlliesCount) & "¬"
    i = Guilds(GuildIndex).Iterador_ProximaRelacion(ALIADOS)
    While i > 0
        News = News & Guilds(i).GuildName & "¬"
        i = Guilds(GuildIndex).Iterador_ProximaRelacion(ALIADOS)
    Wend

    Call SendData(SendTarget.ToIndex, UserIndex, 0, News)

    If Guilds(GuildIndex).EleccionesAbiertas Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Hoy es la votacion para elegir un nuevo líder para el clan!!." & FONTTYPE_GUILD)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La eleccion durara 24 horas, se puede votar a cualquier miembro del clan." & FONTTYPE_GUILD)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Para votar escribe /VOTO NICKNAME." & FONTTYPE_GUILD)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo se computara un voto por miembro. Tu voto no puede ser cambiado." & FONTTYPE_GUILD)
    End If

End Sub

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
'sale solo si no es fundador del clan.

    m_PuedeSalirDeClan = False
    If GuildIndex = 0 Then Exit Function
    
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function
    End If

    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
    If UserList(QuienLoEchaUI).flags.Privilegios = PlayerType.User Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If

    m_PuedeSalirDeClan = UCase$(Guilds(GuildIndex).Fundador) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean

    PuedeFundarUnClan = False
    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.ELV < 25 Or UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 90 Then
        refError = "Para fundar un clan debes ser nivel 25 y tener 90 en liderazgo."
        Exit Function
    End If
    
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            If UserList(UserIndex).Faccion.ArmadaReal <> 1 Then
                refError = "Para fundar un clan real debes ser miembro de la armada."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            If Criminal(UserIndex) Then
                refError = "Para fundar un clan de ciudadanos no debes ser criminal."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            If Not Criminal(UserIndex) Then
                refError = "Para fundar un clan de criminales no debes ser ciudadano."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_LEGION
            If UserList(UserIndex).Faccion.FuerzasCaos <> 1 Then
                refError = "Para fundar un clan del mal debes pertenecer a la legión oscura"
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_MASTER
            If UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then
                refError = "Para fundar un clan sin alineación debes ser un dios."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            If UserList(UserIndex).Faccion.ArmadaReal <> 0 Or UserList(UserIndex).Faccion.FuerzasCaos <> 0 Then
                refError = "Para fundar un clan neutro no debes pertenecer a ninguna facción."
                Exit Function
            End If
    End Select
    
    PuedeFundarUnClan = True
    
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
Dim Promedio    As Long
Dim ELV         As Integer
Dim f           As Byte

    m_EstadoPermiteEntrarChar = False
    
    Personaje = Replace(Personaje, "\", vbNullString)
    Personaje = Replace(Personaje, "/", vbNullString)
    Personaje = Replace(Personaje, ".", vbNullString)
    
    If FileExist(CharPath & Personaje & ".chr") Then
        Promedio = CLng(GetVar(CharPath & Personaje & ".chr", "REP", "Promedio"))
        Select Case Guilds(GuildIndex).Alineacion
            Case ALINEACION_GUILD.ALINEACION_ARMADA
                If Promedio >= 0 Then
                    ELV = CInt(GetVar(CharPath & Personaje & ".chr", "Stats", "ELV"))
                    If ELV >= 25 Then
                        f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoReal"))
                    End If
                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f <> 0, True)
                End If
            Case ALINEACION_GUILD.ALINEACION_CIUDA
                m_EstadoPermiteEntrarChar = Promedio >= 0
            Case ALINEACION_GUILD.ALINEACION_CRIMINAL
                m_EstadoPermiteEntrarChar = Promedio < 0
            Case ALINEACION_GUILD.ALINEACION_NEUTRO
                m_EstadoPermiteEntrarChar = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoReal")) = 0
                m_EstadoPermiteEntrarChar = m_EstadoPermiteEntrarChar And (CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoCaos")) = 0)
            Case ALINEACION_GUILD.ALINEACION_LEGION
                If Promedio < 0 Then
                    ELV = CInt(GetVar(CharPath & Personaje & ".chr", "Stats", "ELV"))
                    If ELV >= 25 Then
                        f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoCaos"))
                    End If
                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f <> 0, True)
                End If
            Case Else
                m_EstadoPermiteEntrarChar = True
        End Select
    End If
End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
    Select Case Guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = Not Criminal(UserIndex) And _
                    IIf(UserList(UserIndex).Stats.ELV >= 25, UserList(UserIndex).Faccion.ArmadaReal <> 0, True)
        Case ALINEACION_GUILD.ALINEACION_LEGION
            m_EstadoPermiteEntrar = Criminal(UserIndex) And _
                    IIf(UserList(UserIndex).Stats.ELV >= 25, UserList(UserIndex).Faccion.FuerzasCaos <> 0, True)
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            m_EstadoPermiteEntrar = UserList(UserIndex).Faccion.ArmadaReal = 0 And UserList(UserIndex).Faccion.FuerzasCaos = 0
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            m_EstadoPermiteEntrar = Not Criminal(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            m_EstadoPermiteEntrar = Criminal(UserIndex)
        Case Else   'game masters
            m_EstadoPermiteEntrar = True
    End Select
End Function


Public Function String2Alineacion(ByRef s As String) As ALINEACION_GUILD
    Select Case s
        Case "Neutro"
            String2Alineacion = ALINEACION_NEUTRO
        Case "Legión oscura"
            String2Alineacion = ALINEACION_LEGION
        Case "Armada Real"
            String2Alineacion = ALINEACION_ARMADA
        Case "Game Masters"
            String2Alineacion = ALINEACION_MASTER
        Case "Legal"
            String2Alineacion = ALINEACION_CIUDA
        Case "Criminal"
            String2Alineacion = ALINEACION_CRIMINAL
    End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            Alineacion2String = "Neutro"
        Case ALINEACION_GUILD.ALINEACION_LEGION
            Alineacion2String = "Legión oscura"
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            Alineacion2String = "Armada Real"
        Case ALINEACION_GUILD.ALINEACION_MASTER
            Alineacion2String = "Game Masters"
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            Alineacion2String = "Legal"
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            Alineacion2String = "Criminal"
    End Select
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    Select Case Relacion
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "A"
        Case RELACIONES_GUILD.GUERRA
            Relacion2String = "G"
        Case RELACIONES_GUILD.PAZ
            Relacion2String = "P"
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "?"
    End Select
End Function

Public Function String2Relacion(ByVal s As String) As RELACIONES_GUILD
    Select Case UCase$(Trim$(s))
        Case vbNullString, "P"
            String2Relacion = PAZ
        Case "G"
            String2Relacion = GUERRA
        Case "A"
            String2Relacion = ALIADOS
        Case Else
            String2Relacion = PAZ
    End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
Dim car     As Byte
Dim i       As Integer

'old function by morgo

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))

    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        GuildNameValido = False
        Exit Function
    End If
    
Next i

GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
Dim i   As Integer

YaExiste = False
GuildName = UCase$(GuildName)

For i = 1 To CANTIDADDECLANES
    YaExiste = (UCase$(Guilds(i).GuildName) = GuildName)
    If YaExiste Then Exit Function
Next i



End Function

Public Function v_AbrirElecciones(ByVal UserIndex As Integer, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer

    v_AbrirElecciones = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GuildIndex) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya están abiertas"
        Exit Function
    End If
    
    v_AbrirElecciones = True
    Call Guilds(GuildIndex).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal UserIndex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer

    v_UsuarioVota = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningún clan"
        Exit Function
    End If

    If Not Guilds(GuildIndex).EleccionesAbiertas Then
        refError = "No hay elecciones abiertas en tu clan."
        Exit Function
    End If
    
    If InStr(1, Guilds(GuildIndex).GetMemberList(","), Votado, vbTextCompare) <= 0 Then
        refError = Votado & " no pertenece al clan"
        Exit Function
    End If

    If Guilds(GuildIndex).YaVoto(UserList(UserIndex).name) Then
        refError = "Ya has votado, no puedes cambiar tu voto"
        Exit Function
    End If
    
    Call Guilds(GuildIndex).ContabilizarVoto(UserList(UserIndex).name, Votado)
    v_UsuarioVota = True

End Function

Public Sub v_RutinaElecciones()
Dim i       As Integer

On Error GoTo errh
    Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Revisando elecciones" & FONTTYPE_SERVER)
    For i = 1 To CANTIDADDECLANES
        If Not Guilds(i) Is Nothing Then
            If Guilds(i).RevisarElecciones Then
                Call SendData(SendTarget.ToAll, 0, 0, "||        > " & Guilds(i).GetLeader & " es el nuevo lider de " & Guilds(i).GuildName & "!" & FONTTYPE_SERVER)
            End If
        End If
proximo:
    Next i
    Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Elecciones revisadas" & FONTTYPE_SERVER)
Exit Sub
errh:
    Call LogError("modGuilds.v_RutinaElecciones():" & Err.Description)
    Resume proximo
End Sub

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
'aca si que vamos a violar las capas deliveradamente ya que
'visual basic no permite declarar metodos de clase
Dim i       As Integer
Dim Temps   As String
    PlayerName = Replace(PlayerName, "\", vbNullString)
    PlayerName = Replace(PlayerName, "/", vbNullString)
    PlayerName = Replace(PlayerName, ".", vbNullString)
    Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "GUILDINDEX")
    If IsNumeric(Temps) Then
        GetGuildIndexFromChar = CInt(Temps)
    Else
        GetGuildIndexFromChar = 0
    End If
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
'me da el indice del guildname
Dim i As Integer

    GuildIndex = 0
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        If UCase$(Guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
Dim i As Integer
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = Guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> UserIndex And (UserList(i).flags.Privilegios < PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then _
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).name & ","
            i = Guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
End Function

Public Function SendGuildsList(ByVal UserIndex As Integer) As String
Dim tStr As String
Dim tInt As Integer

    tStr = CANTIDADDECLANES & ","
    For tInt = 1 To CANTIDADDECLANES
        tStr = tStr & Guilds(tInt).GuildName & ","
    Next tInt
    SendGuildsList = tStr
End Function

Public Function SendGuildDetails(ByRef GuildName As String) As String
Dim tStr    As String
Dim GI      As Integer
Dim i       As Integer

    GI = GuildIndex(GuildName)
    If GI = 0 Then Exit Function
    
    tStr = Guilds(GI).GuildName & "¬"
    tStr = tStr & Guilds(GI).Fundador & "¬"
    tStr = tStr & Guilds(GI).GetFechaFundacion & "¬"
    tStr = tStr & Guilds(GI).GetLeader & "¬"
    tStr = tStr & Guilds(GI).GetURL & "¬"
    tStr = tStr & CStr(Guilds(GI).CantidadDeMiembros) & "¬"
    tStr = tStr & IIf(Guilds(GI).EleccionesAbiertas, "Elecciones abiertas", "Elecciones cerradas") & "¬"
    tStr = tStr & Alineacion2String(Guilds(GI).Alineacion) & "¬"
    tStr = tStr & Guilds(GI).CantidadEnemys & "¬"
    tStr = tStr & Guilds(GI).CantidadAllies & "¬"
    tStr = tStr & Guilds(GI).PuntosAntifaccion & "/" & CStr(MAXANTIFACCION) & "¬"
    For i = 1 To CANTIDADMAXIMACODEX
        tStr = tStr & Guilds(GI).GetCodex(i) & "¬"
    Next i
    tStr = tStr & Guilds(GI).GetDesc
    
    SendGuildDetails = tStr
End Function


Public Function SendGuildLeaderInfo(ByVal UserIndex As Integer) As String
Dim tStr    As String
Dim tInt    As Integer
Dim CantAsp As Integer
Dim GI      As Integer

    SendGuildLeaderInfo = vbNullString
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then Exit Function
    
    '<-------Lista de guilds ---------->
    tStr = CANTIDADDECLANES & "¬"
    
    For tInt = 1 To CANTIDADDECLANES
        tStr = tStr & Guilds(tInt).GuildName & "¬"
    Next tInt
    
    '<-------Lista de miembros ---------->
    tStr = tStr & Guilds(GI).CantidadDeMiembros & "¬"
    tStr = tStr & Guilds(GI).GetMemberList("¬") & "¬"
    
    '<------- Guild News -------->
    tStr = tStr & Replace(Guilds(GI).GetGuildNews, vbCrLf, "º") & "¬"
    
    '<------- Solicitudes ------->
    CantAsp = Guilds(GI).CantidadAspirantes()
    tStr = tStr & CantAsp & "¬"
    If CantAsp > 0 Then
        tStr = tStr & Guilds(GI).GetAspirantes("¬") & "¬"
    End If

    SendGuildLeaderInfo = tStr

End Function


Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = Guilds(GuildIndex).m_Iterador_ProximoUserIndex()
    End If
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    'itera sobre los gms escuchando este clan
    Iterador_ProximoGM = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = Guilds(GuildIndex).Iterador_ProximoGM()
    End If
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
    'itera sobre las propuestas
    r_Iterador_ProximaPropuesta = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        r_Iterador_ProximaPropuesta = Guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)
    End If
End Function


Public Function GMEscuchaClan(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
Dim GI As Integer
'devuelve el guildindex
    GI = GuildIndex(GuildName)
    If GI > 0 Then
        Call Guilds(GI).ConectarGM(UserIndex)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Conectado a : " & GuildName & FONTTYPE_GUILD)
        GMEscuchaClan = GI
        UserList(UserIndex).EscucheClan = GI
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Error, el clan no existe" & FONTTYPE_GUILD)
        GMEscuchaClan = 0
    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
'el index lo tengo que tener de cuando me puse a escuchar
    UserList(UserIndex).EscucheClan = 0
    Call Guilds(GuildIndex).DesconectarGM(UserIndex)
End Sub
Public Function r_DeclararGuerra(ByVal UserIndex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
Dim GI  As Integer
Dim GIG As Integer

    r_DeclararGuerra = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildGuerra)
    
    If GI = GIG Then
        refError = "No puedes declarar la guerra a tu mismo clan"
        Exit Function
    End If

    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, GUERRA)
    Call Guilds(GIG).SetRelacion(GI, GUERRA)

    r_DeclararGuerra = GIG

End Function


Public Function r_AceptarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDePaz = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPaz)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If Guilds(GI).GetRelacion(GIG) <> GUERRA Then
        refError = "No estás en guerra con ese clan"
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar"
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, PAZ)
    Call Guilds(GIG).SetRelacion(GI, PAZ)
    
    r_AceptarPropuestaDePaz = GIG

End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay propuesta de alianza del clan " & GuildPro
        Exit Function
    End If
    
    Call Guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call Guilds(GIG).SetGuildNews(Guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & Guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG

End Function


Public Function r_RechazarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDePaz = 0
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, PAZ) Then
        refError = "No hay propuesta de paz del clan " & GuildPro
        Exit Function
    End If
    
    Call Guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call Guilds(GIG).SetGuildNews(Guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & Guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG

End Function


Public Function r_AceptarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildAllie)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If Guilds(GI).GetRelacion(GIG) <> PAZ Then
        refError = "No estás en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, ALIADOS)
    Call Guilds(GIG).SetRelacion(GI, ALIADOS)
    
    r_AceptarPropuestaDeAlianza = GIG

End Function


Public Function r_ClanGeneraPropuesta(ByVal UserIndex As Integer, ByRef OtroClan As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
Dim OtroClanGI      As Integer
Dim GI              As Integer

    r_ClanGeneraPropuesta = False
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroClan)
    
    If OtroClanGI = GI Then
        refError = "No puedes declarar relaciones con tu propio clan"
        Exit Function
    End If
    
    If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
        refError = "El sistema de clanes esta inconsistente, el otro clan no existe!"
        Exit Function
    End If
    
    If Guilds(OtroClanGI).HayPropuesta(GI, Tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If Tipo = PAZ Then
        If Guilds(GI).GetRelacion(OtroClanGI) <> GUERRA Then
            refError = "No estás en guerra con " & OtroClan
            Exit Function
        End If
    ElseIf Tipo = GUERRA Then
        'por ahora no hay propuestas de guerra
    ElseIf Tipo = ALIADOS Then
        If Guilds(GI).GetRelacion(OtroClanGI) <> PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
            Exit Function
        End If
    End If
    
    Call Guilds(OtroClanGI).SetPropuesta(Tipo, GI, Detalle)
    r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal UserIndex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
Dim OtroClanGI      As Integer
Dim GI              As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroGuild)
    
    If Not Guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
        refError = "No existe la propuesta solicitada"
        Exit Function
    End If
    
    r_VerPropuesta = Guilds(GI).GetPropuesta(OtroClanGI, Tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal UserIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As String
Dim GI  As Integer
Dim i   As Integer


    GI = UserList(UserIndex).GuildIndex
    If GI > 0 And GI <= CANTIDADDECLANES Then
        i = Guilds(GI).Iterador_ProximaPropuesta(Tipo)
        While i > 0
            r_ListaDePropuestas = r_ListaDePropuestas & Guilds(i).GuildName & ","
            i = Guilds(GI).Iterador_ProximaPropuesta(Tipo)
        Wend
        If Len(r_ListaDePropuestas) > 0 Then
            r_ListaDePropuestas = Left$(r_ListaDePropuestas, Len(r_ListaDePropuestas) - 1)
        End If
    End If

End Function

Public Function r_CantidadDePropuestas(ByVal UserIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
Dim GI As Integer
    GI = UserList(UserIndex).GuildIndex
    If GI > 0 And GI <= CANTIDADDECLANES Then
        r_CantidadDePropuestas = Guilds(GI).CantidadPropuestas(Tipo)
    End If
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal Guild As Integer, ByRef Detalles As String)
    Aspirante = Replace(Aspirante, "\", "")
    Aspirante = Replace(Aspirante, "/", "")
    Aspirante = Replace(Aspirante, ".", "")
    Call Guilds(Guild).InformarRechazoEnChar(Aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    Aspirante = Replace(Aspirante, "\", "")
    Aspirante = Replace(Aspirante, "/", "")
    Aspirante = Replace(Aspirante, ".", "")
    a_ObtenerRechazoDeChar = GetVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo")
    Call WriteVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo", vbNullString)
End Function

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, ByRef Nombre As String, ByRef motivo As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim UI              As Integer
Dim NroAspirante    As Integer

    a_RechazarAspirante = False
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If

    NroAspirante = Guilds(GI).NumeroDeAspirante(Nombre)

    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu clan"
        Exit Function
    End If

    Call Guilds(GI).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & Guilds(GI).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef Nombre As String) As String
Dim GI              As Integer
Dim NroAspirante    As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        Exit Function
    End If
    
    NroAspirante = Guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = Guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
    
End Function

Public Function a_DetallesPersonaje(ByVal UserIndex As Integer, ByRef Personaje As String, ByRef refError As String) As String
Dim GI          As Integer
Dim NroAsp      As Integer
Dim tStr        As String
Dim UserFile    As String
Dim Peticiones  As String
Dim Miembro     As String
Dim GuildActual As Integer



    a_DetallesPersonaje = vbNullString
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    Personaje = Replace(Personaje, "\", vbNullString)
    Personaje = Replace(Personaje, "/", vbNullString)
    Personaje = Replace(Personaje, ".", vbNullString)
    
    NroAsp = Guilds(GI).NumeroDeAspirante(Personaje)
    
    If NroAsp = 0 Then
        If InStr(1, Guilds(GI).GetMemberList("."), Personaje, vbTextCompare) <= 0 Then
            refError = "El personaje no es ni aspirante ni miembro del clan"
            Exit Function
        End If
    End If
    
    'ahora traemos la info
    
    UserFile = CharPath & Personaje & ".chr"

    tStr = Personaje & "¬"
    tStr = tStr & GetVar(UserFile, "INIT", "Raza") & "¬"
    tStr = tStr & GetVar(UserFile, "INIT", "Clase") & "¬"
    tStr = tStr & GetVar(UserFile, "INIT", "Genero") & "¬"
    tStr = tStr & GetVar(UserFile, "STATS", "ELV") & "¬"
    tStr = tStr & GetVar(UserFile, "STATS", "GLD") & "¬"
    tStr = tStr & GetVar(UserFile, "STATS", "Banco") & "¬"
    tStr = tStr & GetVar(UserFile, "REP", "Promedio") & "¬"
    
    Peticiones = GetVar(UserFile, "GUILD", "Pedidos")
    tStr = tStr & IIf(Len(Peticiones) > 400, ".." & Right$(Peticiones, 400), Peticiones) & "¬"
    
    Miembro = GetVar(UserFile, "GUILD", "Miembro")
    tStr = tStr & IIf(Len(Miembro) > 400, ".." & Right$(Miembro, 400), Miembro) & "¬"
    
    GuildActual = val(GetVar(UserFile, "GUILD", "GuildIndex"))
    If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
        tStr = tStr & "<" & Guilds(GuildActual).GuildName & ">" & "¬"
    Else
        tStr = tStr & "Ninguno" & "¬"
    End If

    tStr = tStr & GetVar(UserFile, "FACCIONES", "EjercitoReal") & "¬"
    tStr = tStr & GetVar(UserFile, "FACCIONES", "EjercitoCaos") & "¬"
    tStr = tStr & GetVar(UserFile, "FACCIONES", "CiudMatados") & "¬"
    tStr = tStr & GetVar(UserFile, "FACCIONES", "CrimMatados") & "¬"
    
    a_DetallesPersonaje = tStr
End Function

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, ByRef Clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
Dim ViejoSolicitado     As String
Dim ViejoGuildINdex     As Integer
Dim ViejoNroAspirante   As Integer
Dim NuevoGuildIndex     As Integer

    a_NuevoAspirante = False

    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro"
        Exit Function
    End If
    
    If EsNewbie(UserIndex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function
    End If

    NuevoGuildIndex = GuildIndex(Clan)
    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe! Avise a un administrador."
        Exit Function
    End If
    
    If Not m_EstadoPermiteEntrar(UserIndex, NuevoGuildIndex) Then
        refError = "Tu no puedes entrar a un clan de alineación " & Alineacion2String(Guilds(NuevoGuildIndex).Alineacion)
        Exit Function
    End If

    If Guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
        Exit Function
    End If

    ViejoSolicitado = GetVar(CharPath & UserList(UserIndex).name & ".chr", "GUILD", "ASPIRANTEA")

    If ViejoSolicitado <> vbNullString Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)
        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = Guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(UserIndex).name)
            If ViejoNroAspirante > 0 Then
                Call Guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).name, ViejoNroAspirante)
            End If
        Else
            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If
    End If
    
    Call Guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer
Dim AspiranteUI     As Integer

    'un pj ingresa al clan :D

    a_AceptarAspirante = False
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    NroAspirante = Guilds(GI).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan"
        Exit Function
    End If
    
    AspiranteUI = NameIndex(Aspirante)
    If AspiranteUI > 0 Then
        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(Guilds(GI).Alineacion)
            Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    Else
        If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(Guilds(GI).Alineacion)
            Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    End If
    'el pj es aspirante al clan y puede entrar
    
    Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
    Call Guilds(GI).AceptarNuevoMiembro(Aspirante)

    a_AceptarAspirante = True

End Function
