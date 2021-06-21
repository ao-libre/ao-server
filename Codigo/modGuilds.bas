Attribute VB_Name = "modGuilds"
'**************************************************************
' modGuilds.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Implemented by Mariano Barrou (El Oso)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE             As String
'archivo .\guilds\guildinfo.ini o similar

Private Const MAX_GUILDS          As Integer = 1000

'cantidad maxima de guilds en el servidor

Public CANTIDADDECLANES           As Integer
'cantidad actual de clanes en el servidor

Public guilds(1 To MAX_GUILDS)    As clsClan
'array global de guilds, se indexa por userlist().guildindex

Private Const CANTIDADMAXIMACODEX As Byte = 8

'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES        As Byte = 10

'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION      As Byte = 5

'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

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
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando guildsinfo.inf."

    Dim CantClanes As String

    Dim i          As Integer

    Dim TempStr    As String

    Dim Alin       As ALINEACION_GUILD
    
    GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"

    CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    
    If IsNumeric(CantClanes) Then
        CANTIDADDECLANES = CInt(CantClanes)
    Else
        CANTIDADDECLANES = 0

    End If
    
    For i = 1 To CANTIDADDECLANES
        Set guilds(i) = New clsClan
        TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
        Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
        Call guilds(i).Inicializar(TempStr, i, Alin)
    Next i

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo guildsinfo.inf."
    
End Sub

Public Function m_ConectarMiembroAClan(ByVal Userindex As Integer, _
                                       ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim NuevaA As Boolean

    Dim News   As String

    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(Userindex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(Userindex)
        UserList(Userindex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(Userindex, True, NuevaA)

        If NuevaA Then News = News & "El clan tiene nueva alineacion."

        'If NuevoL Or NuevaA Then Call guilds(GuildIndex).SetGuildNews(News)
    End If

End Function

Public Function m_ValidarPermanencia(ByVal Userindex As Integer, _
                                     ByVal SumaAntifaccion As Boolean, _
                                     ByRef CambioAlineacion As Boolean) As Boolean
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 14/12/2009
    '25/03/2009: ZaMa - Desequipo los items faccionarios que tenga el funda al abandonar la faccion
    '14/12/2009: ZaMa - La alineacion del clan depende del lider
    '14/02/2010: ZaMa - Ya no es necesario saber si el lider cambia, ya que no puede cambiar.
    '***************************************************

    Dim GuildIndex As Integer

    m_ValidarPermanencia = True
    
    GuildIndex = UserList(Userindex).GuildIndex

    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    
    If Not m_EstadoPermiteEntrar(Userindex, GuildIndex) Then
        
        ' Es el lider, bajamos 1 rango de alineacion
        If m_EsGuildLeader(UserList(Userindex).Name, GuildIndex) Then
            Call LogClanes(UserList(Userindex).Name & ", lider de " & guilds(GuildIndex).GuildName & " hizo bajar la alienacion de su clan.")
        
            CambioAlineacion = True
            
            ' Por si paso de ser armada/legion a pk/ciuda, chequeo de nuevo
            Do
                Call UpdateGuildMembers(GuildIndex)
            Loop Until m_EstadoPermiteEntrar(Userindex, GuildIndex)

        Else
            Call LogClanes(UserList(Userindex).Name & " de " & guilds(GuildIndex).GuildName & " es expulsado en validar permanencia.")
        
            m_ValidarPermanencia = False

            If SumaAntifaccion Then guilds(GuildIndex).PuntosAntifaccion = guilds(GuildIndex).PuntosAntifaccion + 1
            
            CambioAlineacion = guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
            
            Call LogClanes(UserList(Userindex).Name & " de " & guilds(GuildIndex).GuildName & IIf(CambioAlineacion, " SI ", " NO ") & "provoca cambio de alineacion. MAXANT:" & CambioAlineacion)
            
            Call m_EcharMiembroDeClan(-1, UserList(Userindex).Name)
            
            ' Llegamos a la maxima cantidad de antifacciones permitidas, bajamos un grado de alineacion
            If CambioAlineacion Then
                Call UpdateGuildMembers(GuildIndex)

            End If

        End If

    End If

End Function

Private Sub UpdateGuildMembers(ByVal GuildIndex As Integer)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Pulo detalles en el funcionamiento general.
    '***************************************************
    Dim GuildMembers() As String

    Dim TotalMembers   As Integer

    Dim MemberIndex    As Long

    Dim Sale           As Boolean

    Dim MemberName     As String

    Dim Userindex      As Integer

    Dim Reenlistadas   As Integer
    
    ' Si devuelve true, cambio a neutro y echamos a todos los que esten de mas, sino no echamos a nadie
    If guilds(GuildIndex).CambiarAlineacion(BajarGrado(GuildIndex)) Then 'ALINEACION_NEUTRO)
        
        'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo alterar
        'internamente al iterador en el proceso
        GuildMembers = guilds(GuildIndex).GetMemberList()
        TotalMembers = UBound(GuildMembers)
        
        For MemberIndex = 0 To TotalMembers
            MemberName = GuildMembers(MemberIndex)
            
            'vamos a violar un poco de capas..
            Userindex = NameIndex(MemberName)

            If Userindex > 0 Then
                Sale = Not m_EstadoPermiteEntrar(Userindex, GuildIndex)
            Else
                Sale = Not m_EstadoPermiteEntrarChar(MemberName, GuildIndex)

            End If

            If Sale Then
                If m_EsGuildLeader(MemberName, GuildIndex) Then  'hay que sacarlo de las facciones. (CHOTS: no cuenta como reenlistada.)
                 
                    If Userindex > 0 Then
                        If UserList(Userindex).Faccion.ArmadaReal <> 0 Then
                            Call ExpulsarFaccionReal(Userindex)
                            UserList(Userindex).Faccion.Reenlistadas = UserList(Userindex).Faccion.Reenlistadas - 1
                        ElseIf UserList(Userindex).Faccion.FuerzasCaos <> 0 Then
                            Call ExpulsarFaccionCaos(Userindex)
                            UserList(Userindex).Faccion.Reenlistadas = UserList(Userindex).Faccion.Reenlistadas - 1

                        End If

                    Else

                        If PersonajeExiste(MemberName) Then
                            Call KickUserFacciones(MemberName)
                            Reenlistadas = GetUserReenlists(MemberName)
                            Call SaveUserReenlists(MemberName, IIf(Reenlistadas > 1, Reenlistadas - 1, Reenlistadas))

                        End If

                    End If

                Else    'sale si no es guildLeader
                    Call m_EcharMiembroDeClan(-1, MemberName)

                End If

            End If

        Next MemberIndex

    Else
        ' Resetea los puntos de antifaccion
        guilds(GuildIndex).PuntosAntifaccion = 0

    End If

End Sub

Private Function BajarGrado(ByVal GuildIndex As Integer) As ALINEACION_GUILD
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 27/11/2009
    'Reduce el grado de la alineacion a partir de la alineacion dada
    '***************************************************

    Select Case guilds(GuildIndex).Alineacion

        Case ALINEACION_ARMADA
            BajarGrado = ALINEACION_CIUDA

        Case ALINEACION_LEGION
            BajarGrado = ALINEACION_CRIMINAL

        Case Else
            BajarGrado = ALINEACION_NEUTRO

    End Select

End Function

Public Sub m_DesconectarMiembroDelClan(ByVal Userindex As Integer, _
                                       ByVal GuildIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If UserList(Userindex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(Userindex)

End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, _
                                 ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))

End Function

Private Function m_EsGuildFounder(ByRef PJ As String, _
                                  ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))

End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, _
                                     ByVal Expulsado As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'UI echa a Expulsado del clan de Expulsado
    Dim Userindex As Integer

    Dim GI        As Integer
    
    m_EcharMiembroDeClan = 0

    Userindex = NameIndex(Expulsado)

    If Userindex > 0 Then
        'pj online
        GI = UserList(Userindex).GuildIndex

        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                Call guilds(GI).DesConectarMiembro(Userindex)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(Userindex).GuildIndex = 0
                Call RefreshCharStatus(Userindex)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0

            End If

        Else
            m_EcharMiembroDeClan = 0

        End If

    Else
        'pj offline
        GI = GetUserGuildIndex(Expulsado)

        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0

            End If

        Else
            m_EcharMiembroDeClan = 0

        End If

    End If

End Function

Public Sub ActualizarWebSite(ByVal Userindex As Integer, ByRef Web As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GI As Integer

    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then Exit Sub
    
    Call guilds(GI).SetURL(Web)
    
End Sub

Public Sub ChangeCodexAndDesc(ByRef desc As String, _
                              ByRef codex() As String, _
                              ByVal GuildIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Long
    
    If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
    
    With guilds(GuildIndex)
        Call .SetDesc(desc)
        
        For i = 0 To UBound(codex())
            Call .SetCodex(i, codex(i))
        Next i
        
        For i = i To CANTIDADMAXIMACODEX
            Call .SetCodex(i, vbNullString)
        Next i

    End With

End Sub

Public Sub ActualizarNoticias(ByVal Userindex As Integer, ByRef Datos As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 21/02/2010
    '21/02/2010: ZaMa - Ahora le avisa a los miembros que cambio el guildnews.
    '***************************************************

    Dim GI As Integer

    With UserList(Userindex)
        GI = .GuildIndex
        
        If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
        
        If Not m_EsGuildLeader(.Name, GI) Then Exit Sub
        
        Call guilds(GI).SetGuildNews(Datos)
        
        Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & " ha actualizado las noticias del clan!"))

    End With

End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, _
                               ByRef desc As String, _
                               ByRef GuildName As String, _
                               ByRef URL As String, _
                               ByRef codex() As String, _
                               ByVal Alineacion As ALINEACION_GUILD, _
                               ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim CantCodex   As Integer

    Dim i           As Integer

    Dim DummyString As String

    CrearNuevoClan = False

    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function

    End If

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan invalido."
        Exit Function

    End If
    
    If YaExiste(GuildName) Then
        refError = "Ya existe un clan con ese nombre."
        Exit Function

    End If

    CantCodex = UBound(codex()) + 1

    'tenemos todo para fundar ya
    If CANTIDADDECLANES < UBound(guilds) Then
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan

        'constructor custom de la clase clan
        Set guilds(CANTIDADDECLANES) = New clsClan
        
        With guilds(CANTIDADDECLANES)
            Call .Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
            
            'Damos de alta al clan como nuevo inicializando sus archivos
            Call .InicializarNuevoClan(UserList(FundadorIndex).Name)
            
            'seteamos codex y descripcion
            For i = 1 To CantCodex
                Call .SetCodex(i, codex(i - 1))
            Next i

            Call .SetDesc(desc)
            Call .SetGuildNews("Clan creado con alineacion: " & Alineacion2String(Alineacion))
            Call .SetLeader(UserList(FundadorIndex).Name)
            Call .SetURL(URL)
            
            '"conectamos" al nuevo miembro a la lista de la clase
            Call .AceptarNuevoMiembro(UserList(FundadorIndex).Name)
            Call .ConectarMiembro(FundadorIndex)

        End With
        
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call RefreshCharStatus(FundadorIndex)
        
        For i = 1 To CANTIDADDECLANES - 1
            Call guilds(i).ProcesarFundacionDeOtroClan
        Next i

    Else
        refError = "No hay mas slots para fundar clanes. Consulte a un administrador."
        Exit Function

    End If
    
    CrearNuevoClan = True

End Function

Public Sub SendGuildNews(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GuildIndex As Integer

    Dim i          As Integer

    Dim go         As Integer

    GuildIndex = UserList(Userindex).GuildIndex

    If GuildIndex = 0 Then Exit Sub

    Dim enemies() As String
    
    With guilds(GuildIndex)

        If .CantidadEnemys Then
            ReDim enemies(0 To .CantidadEnemys - 1) As String
        Else
            ReDim enemies(0)

        End If
        
        Dim allies() As String
        
        If .CantidadAllies Then
            ReDim allies(0 To .CantidadAllies - 1) As String
        Else
            ReDim allies(0)

        End If
        
        i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
        go = 0
        
        While i > 0

            enemies(go) = guilds(i).GuildName
            i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
            go = go + 1
        Wend
        
        i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
        go = 0
        
        While i > 0

            allies(go) = guilds(i).GuildName
            i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
        Wend
    
        Call WriteGuildNews(Userindex, .GetGuildNews, enemies, allies)
    
        If .EleccionesAbiertas Then
            Call WriteConsoleMsg(Userindex, "Hoy es la votacion para elegir un nuevo lider para el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(Userindex, "La eleccion durara 24 horas, se puede votar a cualquier miembro del clan.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(Userindex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(Userindex, "Solo se computara un voto por miembro. Tu voto no puede ser cambiado.", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, _
                                   ByVal GuildIndex As Integer, _
                                   ByVal QuienLoEchaUI As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

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
    If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.User Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function

            End If

        End If

    End If

    ' Ahora el lider es el unico que no puede salir del clan
    m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).GetLeader) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnClan(ByVal Userindex As Integer, _
                                  ByVal Alineacion As ALINEACION_GUILD, _
                                  ByRef refError As String) As Boolean
    '***************************************************
    'Autor: Unknown
    'Last Modification: 27/11/2009
    'Returns true if can Found a guild
    '27/11/2009: ZaMa - Ahora valida si ya fundo clan o no.
    '***************************************************
    
    If UserList(Userindex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function

    End If
    
    If UserList(Userindex).Stats.ELV < 25 Or UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) < 90 Then
        refError = "Para fundar un clan debes ser nivel 25 y tener 90 skills en liderazgo."
        Exit Function

    End If
    
    Select Case Alineacion

        Case ALINEACION_GUILD.ALINEACION_ARMADA

            If UserList(Userindex).Faccion.ArmadaReal <> 1 Then
                refError = "Para fundar un clan real debes ser miembro del ejercito real."
                Exit Function

            End If

        Case ALINEACION_GUILD.ALINEACION_CIUDA

            If criminal(Userindex) Then
                refError = "Para fundar un clan de ciudadanos no debes ser criminal."
                Exit Function

            End If

        Case ALINEACION_GUILD.ALINEACION_CRIMINAL

            If Not criminal(Userindex) Then
                refError = "Para fundar un clan de criminales no debes ser ciudadano."
                Exit Function

            End If

        Case ALINEACION_GUILD.ALINEACION_LEGION

            If UserList(Userindex).Faccion.FuerzasCaos <> 1 Then
                refError = "Para fundar un clan del mal debes pertenecer a la legion oscura."
                Exit Function

            End If

        Case ALINEACION_GUILD.ALINEACION_MASTER

            If UserList(Userindex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                refError = "Para fundar un clan sin alineacion debes ser un dios."
                Exit Function

            End If

        Case ALINEACION_GUILD.ALINEACION_NEUTRO

            If UserList(Userindex).Faccion.ArmadaReal <> 0 Or UserList(Userindex).Faccion.FuerzasCaos <> 0 Then
                refError = "Para fundar un clan neutro no debes pertenecer a ninguna faccion."
                Exit Function

            End If

    End Select
    
    PuedeFundarUnClan = True
    
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, _
                                           ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Promedio As Long

    Dim ELV      As Integer

    Dim f        As Byte

    m_EstadoPermiteEntrarChar = False
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace(Personaje, "\", vbNullString)

    End If

    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace(Personaje, "/", vbNullString)

    End If

    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace(Personaje, ".", vbNullString)

    End If
    
    If PersonajeExiste(Personaje) Then
        Promedio = GetUserPromedio(Personaje)

        Select Case guilds(GuildIndex).Alineacion

            Case ALINEACION_GUILD.ALINEACION_ARMADA

                If Promedio >= 0 Then
                    ELV = GetUserLevel(Personaje)

                    If ELV >= 25 Then
                        f = UserBelongsToRoyalArmy(Personaje)

                    End If

                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f, True)

                End If

            Case ALINEACION_GUILD.ALINEACION_CIUDA
                m_EstadoPermiteEntrarChar = Promedio >= 0

            Case ALINEACION_GUILD.ALINEACION_CRIMINAL
                m_EstadoPermiteEntrarChar = Promedio < 0

            Case ALINEACION_GUILD.ALINEACION_NEUTRO
                m_EstadoPermiteEntrarChar = (UserBelongsToRoyalArmy(Personaje) = False And UserBelongsToChaosLegion(Personaje) = False)

            Case ALINEACION_GUILD.ALINEACION_LEGION

                If Promedio < 0 Then
                    ELV = GetUserLevel(Personaje)

                    If ELV >= 25 Then
                        f = UserBelongsToChaosLegion(Personaje)

                    End If

                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f, True)

                End If

            Case Else
                m_EstadoPermiteEntrarChar = True

        End Select

    End If

End Function

Private Function m_EstadoPermiteEntrar(ByVal Userindex As Integer, _
                                       ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case guilds(GuildIndex).Alineacion

        Case ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = Not criminal(Userindex) And IIf(UserList(Userindex).Stats.ELV >= 25, UserList(Userindex).Faccion.ArmadaReal <> 0, True)
        
        Case ALINEACION_GUILD.ALINEACION_LEGION
            m_EstadoPermiteEntrar = criminal(Userindex) And IIf(UserList(Userindex).Stats.ELV >= 25, UserList(Userindex).Faccion.FuerzasCaos <> 0, True)
        
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            m_EstadoPermiteEntrar = UserList(Userindex).Faccion.ArmadaReal = 0 And UserList(Userindex).Faccion.FuerzasCaos = 0
        
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            m_EstadoPermiteEntrar = Not criminal(Userindex)
        
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            m_EstadoPermiteEntrar = criminal(Userindex)
        
        Case Else   'game masters
            m_EstadoPermiteEntrar = True

    End Select

End Function

Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case S

        Case "Neutral"
            String2Alineacion = ALINEACION_NEUTRO

        Case "Del Mal"
            String2Alineacion = ALINEACION_LEGION

        Case "Real"
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
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case Alineacion

        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            Alineacion2String = "Neutral"

        Case ALINEACION_GUILD.ALINEACION_LEGION
            Alineacion2String = "Del Mal"

        Case ALINEACION_GUILD.ALINEACION_ARMADA
            Alineacion2String = "Real"

        Case ALINEACION_GUILD.ALINEACION_MASTER
            Alineacion2String = "Game Masters"

        Case ALINEACION_GUILD.ALINEACION_CIUDA
            Alineacion2String = "Legal"

        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            Alineacion2String = "Criminal"

    End Select

End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

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

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case UCase$(Trim$(S))

        Case vbNullString, "P"
            String2Relacion = RELACIONES_GUILD.PAZ

        Case "G"
            String2Relacion = RELACIONES_GUILD.GUERRA

        Case "A"
            String2Relacion = RELACIONES_GUILD.ALIADOS

        Case Else
            String2Relacion = RELACIONES_GUILD.PAZ

    End Select

End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim car As Byte

    Dim i   As Integer

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
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    YaExiste = False
    GuildName = UCase$(GuildName)

    For i = 1 To CANTIDADDECLANES
        YaExiste = (UCase$(guilds(i).GuildName) = GuildName)

        If YaExiste Then Exit Function
    Next i

End Function

Public Function HasFound(ByRef UserName As String) As Boolean

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 27/11/2009
    'Returns true if it's already the founder of other guild
    '***************************************************
    Dim i    As Long

    Dim Name As String

    Name = UCase$(UserName)

    For i = 1 To CANTIDADDECLANES
        HasFound = (UCase$(guilds(i).Fundador) = Name)

        If HasFound Then Exit Function
    Next i

End Function

Public Function v_AbrirElecciones(ByVal Userindex As Integer, _
                                  ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GuildIndex As Integer

    v_AbrirElecciones = False
    GuildIndex = UserList(Userindex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningun clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GuildIndex) Then
        refError = "No eres el lider de tu clan"
        Exit Function

    End If
    
    If guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya estan abiertas."
        Exit Function

    End If
    
    v_AbrirElecciones = True
    Call guilds(GuildIndex).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal Userindex As Integer, _
                              ByRef Votado As String, _
                              ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GuildIndex As Integer

    Dim list()     As String

    Dim i          As Long

    v_UsuarioVota = False
    GuildIndex = UserList(Userindex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningun clan."
        Exit Function

    End If

    With guilds(GuildIndex)

        If Not .EleccionesAbiertas Then
            refError = "No hay elecciones abiertas en tu clan."
            Exit Function

        End If
        
        list = .GetMemberList()

        For i = 0 To UBound(list())

            If UCase$(Votado) = list(i) Then Exit For
        Next i
        
        If i > UBound(list()) Then
            refError = Votado & " no pertenece al clan."
            Exit Function

        End If
        
        If .YaVoto(UserList(Userindex).Name) Then
            refError = "Ya has votado, no puedes cambiar tu voto."
            Exit Function

        End If
        
        Call .ContabilizarVoto(UserList(Userindex).Name, Votado)
        v_UsuarioVota = True

    End With

End Function

Public Sub v_RutinaElecciones()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    On Error GoTo errh

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Revisando elecciones", FontTypeNames.FONTTYPE_SERVER))

    For i = 1 To CANTIDADDECLANES

        If Not guilds(i) Is Nothing Then
            If guilds(i).RevisarElecciones Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & guilds(i).GetLeader & " es el nuevo lider de " & guilds(i).GuildName & ".", FontTypeNames.FONTTYPE_SERVER))

            End If

        End If

proximo:
    Next i

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Elecciones revisadas.", FontTypeNames.FONTTYPE_SERVER))
    Exit Sub
errh:
    Call LogError("modGuilds.v_RutinaElecciones():" & Err.description)

    Resume proximo

End Sub

Public Function GuildIndex(ByRef GuildName As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'me da el indice del guildname
    Dim i As Integer

    GuildIndex = 0
    GuildName = UCase$(GuildName)

    For i = 1 To CANTIDADDECLANES

        If UCase$(guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function

        End If

    Next i

End Function

Public Function m_ListaDeMiembrosOnline(ByVal Userindex As Integer, _
                                        ByVal GuildIndex As Integer) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Solo dioses pueden ver otros dioses online.
    '***************************************************

    Dim i    As Integer

    Dim priv As PlayerType

    priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
    
    ' Solo dioses pueden ver otros dioses online
    If UserList(Userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
        priv = priv Or PlayerType.Dios Or PlayerType.Admin

    End If
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        
        ' Horrible, tengo que decirlo..
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        
        While i > 0
        
            'No mostramos dioses y admins
            If i <> Userindex And (UserList(i).flags.Privilegios And priv) Then
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","

            End If
            
            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend

    End If
    
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)

    End If

End Function

Public Function PrepareGuildsList() As String()
    '***************************************************
    'Author: Unknown
    'Last Modification: 08/09/2019
    '08/09/2019: Jopi - Ignoro los clanes cerrados.
    '***************************************************

    Dim tStr() As String

    Dim i      As Long
    
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        
        For i = 1 To CANTIDADDECLANES
        
            If LenB(guilds(i).GuildName) <> 0 Then
                If guilds(i).GuildName <> "CLAN CERRADO" Then
                    tStr(i - 1) = guilds(i).GuildName
                End If
            End If
            
        Next i

    End If
    
    PrepareGuildsList = tStr

End Function

Public Sub SendGuildDetails(ByVal Userindex As Integer, ByRef GuildName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim codex(CANTIDADMAXIMACODEX - 1) As String

    Dim GI                             As Integer

    Dim i                              As Long

    GI = GuildIndex(GuildName)

    If GI = 0 Then Exit Sub
    
    With guilds(GI)

        For i = 1 To CANTIDADMAXIMACODEX
            codex(i - 1) = .GetCodex(i)
        Next i
        
        Call Protocol.WriteGuildDetails(Userindex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, .GetURL, .CantidadDeMiembros, .EleccionesAbiertas, Alineacion2String(.Alineacion), .CantidadEnemys, .CantidadAllies, .PuntosAntifaccion & "/" & CStr(MAXANTIFACCION), codex, .GetDesc)

    End With

End Sub

Public Sub SendGuildLeaderInfo(ByVal Userindex As Integer)

    '***************************************************
    'Autor: Mariano Barrou (El Oso)
    'Last Modification: 12/10/06
    'Las Modified By: Juan Martin Sotuyo Dodero (Maraxus)
    '***************************************************
    Dim GI              As Integer

    Dim guildList()     As String

    Dim MemberList()    As String

    Dim aspirantsList() As String

    With UserList(Userindex)
        GI = .GuildIndex
        
        guildList = PrepareGuildsList()
        
        If GI <= 0 Or GI > CANTIDADDECLANES Then
            'Send the guild list instead
            Call WriteGuildList(Userindex, guildList)
            Exit Sub

        End If
        
        MemberList = guilds(GI).GetMemberList()
        
        If Not m_EsGuildLeader(.Name, GI) Then
            'Send the guild list instead
            Call WriteGuildMemberInfo(Userindex, guildList, MemberList)
            Exit Sub

        End If
        
        aspirantsList = guilds(GI).GetAspirantes()
        
        Call WriteGuildLeaderInfo(Userindex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList)

    End With

End Sub

Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0

    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()

    End If

End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'itera sobre los gms escuchando este clan
    Iterador_ProximoGM = 0

    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()

    End If

End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, _
                                            ByVal Tipo As RELACIONES_GUILD) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'itera sobre las propuestas
    r_Iterador_ProximaPropuesta = 0

    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        r_Iterador_ProximaPropuesta = guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)

    End If

End Function

Public Function GMEscuchaClan(ByVal Userindex As Integer, _
                              ByVal GuildName As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GI As Integer

    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(Userindex).EscucheClan <> 0 Then
        'Quit listening to previous guild!!
        Call WriteConsoleMsg(Userindex, "Dejas de escuchar a : " & guilds(UserList(Userindex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
        guilds(UserList(Userindex).EscucheClan).DesconectarGM (Userindex)
        Exit Function

    End If
    
    'devuelve el guildindex
    GI = GuildIndex(GuildName)

    If GI > 0 Then
        If UserList(Userindex).EscucheClan <> 0 Then
            If UserList(Userindex).EscucheClan = GI Then
                'Already listening to them...
                Call WriteConsoleMsg(Userindex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaClan = GI
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(Userindex, "Dejas de escuchar a : " & guilds(UserList(Userindex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
                guilds(UserList(Userindex).EscucheClan).DesconectarGM (Userindex)

            End If

        End If
        
        Call guilds(GI).ConectarGM(Userindex)
        Call WriteConsoleMsg(Userindex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = GI
        UserList(Userindex).EscucheClan = GI
    Else
        Call WriteConsoleMsg(Userindex, "Error, el clan no existe.", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = 0

    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal Userindex As Integer, ByVal GuildIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'el index lo tengo que tener de cuando me puse a escuchar
    UserList(Userindex).EscucheClan = 0
    Call guilds(GuildIndex).DesconectarGM(Userindex)

End Sub

Public Function r_DeclararGuerra(ByVal Userindex As Integer, _
                                 ByRef GuildGuerra As String, _
                                 ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GI  As Integer

    Dim GIG As Integer

    r_DeclararGuerra = 0
    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function

    End If
    
    GIG = GuildIndex(GuildGuerra)

    If guilds(GI).GetRelacion(GIG) = GUERRA Then
        refError = "Tu clan ya esta en guerra con " & GuildGuerra & "."
        Exit Function

    End If
        
    If GI = GIG Then
        refError = "No puedes declarar la guerra a tu mismo clan."
        Exit Function

    End If

    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function

    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.GUERRA)
    
    r_DeclararGuerra = GIG

End Function

Public Function r_AceptarPropuestaDePaz(ByVal Userindex As Integer, _
                                        ByRef GuildPaz As String, _
                                        ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
    Dim GI  As Integer

    Dim GIG As Integer

    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function

    End If

    GIG = GuildIndex(GuildPaz)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function

    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
        refError = "No estas en guerra con ese clan."
        Exit Function

    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar."
        Exit Function

    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.PAZ)
    
    r_AceptarPropuestaDePaz = GIG

End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal Userindex As Integer, _
                                             ByRef GuildPro As String, _
                                             ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'devuelve el index al clan guildPro
    Dim GI  As Integer

    Dim GIG As Integer

    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(Userindex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function

    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function

    End If
    
    If Not guilds(GI).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay propuesta de alianza del clan " & GuildPro
        Exit Function

    End If
    
    Call guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG

End Function

Public Function r_RechazarPropuestaDePaz(ByVal Userindex As Integer, _
                                         ByRef GuildPro As String, _
                                         ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'devuelve el index al clan guildPro
    Dim GI  As Integer

    Dim GIG As Integer

    r_RechazarPropuestaDePaz = 0
    GI = UserList(Userindex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function

    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function

    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay propuesta de paz del clan " & GuildPro
        Exit Function

    End If
    
    Call guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG

End Function

Public Function r_AceptarPropuestaDeAlianza(ByVal Userindex As Integer, _
                                            ByRef GuildAllie As String, _
                                            ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
    Dim GI  As Integer

    Dim GIG As Integer

    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function

    End If

    GIG = GuildIndex(GuildAllie)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function

    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
        refError = "No estas en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function

    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function

    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.ALIADOS)
    
    r_AceptarPropuestaDeAlianza = GIG

End Function

Public Function r_ClanGeneraPropuesta(ByVal Userindex As Integer, _
                                      ByRef OtroClan As String, _
                                      ByVal Tipo As RELACIONES_GUILD, _
                                      ByRef Detalle As String, _
                                      ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim OtroClanGI As Integer

    Dim GI         As Integer

    r_ClanGeneraPropuesta = False
    
    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function

    End If
    
    OtroClanGI = GuildIndex(OtroClan)
    
    If OtroClanGI = GI Then
        refError = "No puedes declarar relaciones con tu propio clan."
        Exit Function

    End If
    
    If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
        refError = "El sistema de clanes esta inconsistente, el otro clan no existe."
        Exit Function

    End If
    
    If guilds(OtroClanGI).HayPropuesta(GI, Tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function

    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If Tipo = RELACIONES_GUILD.PAZ Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
            refError = "No estas en guerra con " & OtroClan
            Exit Function

        End If

    ElseIf Tipo = RELACIONES_GUILD.GUERRA Then
        'por ahora no hay propuestas de guerra
    ElseIf Tipo = RELACIONES_GUILD.ALIADOS Then

        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
            Exit Function

        End If

    End If
    
    Call guilds(OtroClanGI).SetPropuesta(Tipo, GI, Detalle)
    r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal Userindex As Integer, _
                               ByRef OtroGuild As String, _
                               ByVal Tipo As RELACIONES_GUILD, _
                               ByRef refError As String) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim OtroClanGI As Integer

    Dim GI         As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function

    End If
    
    OtroClanGI = GuildIndex(OtroGuild)
    
    If Not guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
        refError = "No existe la propuesta solicitada."
        Exit Function

    End If
    
    r_VerPropuesta = guilds(GI).GetPropuesta(OtroClanGI, Tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal Userindex As Integer, _
                                    ByVal Tipo As RELACIONES_GUILD) As String()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GI            As Integer

    Dim i             As Integer

    Dim proposalCount As Integer

    Dim proposals()   As String
    
    GI = UserList(Userindex).GuildIndex
    
    If GI > 0 And GI <= CANTIDADDECLANES Then

        With guilds(GI)
            proposalCount = .CantidadPropuestas(Tipo)
            
            'Resize array to contain all proposals
            If proposalCount > 0 Then
                ReDim proposals(proposalCount - 1) As String
            Else
                ReDim proposals(0) As String

            End If
            
            'Store each guild name
            For i = 0 To proposalCount - 1
                proposals(i) = guilds(.Iterador_ProximaPropuesta(Tipo)).GuildName
            Next i

        End With

    End If
    
    r_ListaDePropuestas = proposals

End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, _
                                   ByVal Guild As Integer, _
                                   ByRef Detalles As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")

    End If

    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")

    End If

    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")

    End If

    Call SaveUserGuildRejectionReason(Aspirante, Detalles)

End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")

    End If

    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")

    End If

    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")

    End If

    a_ObtenerRechazoDeChar = GetUserGuildRejectionReason(Aspirante)
    Call SaveUserGuildRejectionReason(Aspirante, vbNullString)

End Function

Public Function a_RechazarAspirante(ByVal Userindex As Integer, _
                                    ByRef Nombre As String, _
                                    ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GI           As Integer

    Dim NroAspirante As Integer

    a_RechazarAspirante = False
    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningun clan"
        Exit Function

    End If

    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)

    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu clan."
        Exit Function

    End If

    Call guilds(GI).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal Userindex As Integer, _
                                    ByRef Nombre As String) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GI           As Integer

    Dim NroAspirante As Integer

    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        Exit Function

    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)

    If NroAspirante > 0 Then
        a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)

    End If
    
End Function

Public Sub SendDetallesPersonaje(ByVal Userindex As Integer, ByVal Personaje As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GI     As Integer

    Dim NroAsp As Integer

    Dim list() As String

    Dim i      As Long
    
    On Error GoTo Error

    GI = UserList(Userindex).GuildIndex
    
    Personaje = UCase$(Personaje)
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Call Protocol.WriteConsoleMsg(Userindex, "No perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        Call Protocol.WriteConsoleMsg(Userindex, "No eres el lider de tu clan.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace$(Personaje, "\", vbNullString)

    End If

    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)

    End If

    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)

    End If
    
    NroAsp = guilds(GI).NumeroDeAspirante(Personaje)
    
    If NroAsp = 0 Then
        list = guilds(GI).GetMemberList()
        
        For i = 0 To UBound(list())

            If Personaje = list(i) Then Exit For
        Next i
        
        If i > UBound(list()) Then
            Call Protocol.WriteConsoleMsg(Userindex, "El personaje no es ni aspirante ni miembro del clan.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

    End If

    If Not Database_Enabled Then
        Call SendCharacterInfoCharfile(Userindex, Personaje)
    Else
        Call SendCharacterInfoDatabase(Userindex, Personaje)

    End If

    Exit Sub
Error:

    If Not PersonajeExiste(Personaje) Then
        Call LogError("El usuario " & UserList(Userindex).Name & " (" & Userindex & " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
    Else
        Call LogError("[" & Err.Number & "] " & Err.description & " En la rutina SendDetallesPersonaje, por el usuario " & UserList(Userindex).Name & " (" & Userindex & " ), pidiendo informacion sobre el personaje " & Personaje)

    End If

End Sub

Public Function a_NuevoAspirante(ByVal Userindex As Integer, _
                                 ByRef clan As String, _
                                 ByRef Solicitud As String, _
                                 ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ViejoSolicitado   As String

    Dim ViejoGuildINdex   As Integer

    Dim ViejoNroAspirante As Integer

    Dim NuevoGuildIndex   As Integer

    a_NuevoAspirante = False

    If UserList(Userindex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro."
        Exit Function

    End If
    
    If EsNewbie(Userindex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function

    End If

    NuevoGuildIndex = GuildIndex(clan)

    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe, avise a un administrador."
        Exit Function

    End If
    
    If Not m_EstadoPermiteEntrar(Userindex, NuevoGuildIndex) Then
        refError = "Tu no puedes entrar a un clan de alineacion " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
        Exit Function

    End If

    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contactate con un miembro para que procese las solicitudes."
        Exit Function

    End If

    ViejoSolicitado = GetUserGuildAspirant(UserList(Userindex).Name)

    If LenB(ViejoSolicitado) <> 0 Then
        'borramos la vieja solicitud
        ViejoGuildINdex = ViejoSolicitado

        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(Userindex).Name)

            If ViejoNroAspirante > 0 Then
                Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(Userindex).Name, ViejoNroAspirante)

            End If

        Else

            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If

    End If
    
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(Userindex).Name, Solicitud)
    a_NuevoAspirante = True

End Function

Public Function a_AceptarAspirante(ByVal Userindex As Integer, _
                                   ByRef Aspirante As String, _
                                   ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GI           As Integer

    Dim NroAspirante As Integer

    Dim AspiranteUI  As Integer

    'un pj ingresa al clan :D

    a_AceptarAspirante = False
    
    GI = UserList(Userindex).GuildIndex

    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningun clan"
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan"
        Exit Function

    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan."
        Exit Function

    End If
    
    AspiranteUI = NameIndex(Aspirante)

    If AspiranteUI > 0 Then

        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = Aspirante & " no puede entrar a un clan de alineacion " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function

        End If

    Else

        If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
            refError = Aspirante & " no puede entrar a un clan de alineacion " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf GetUserGuildIndex(Aspirante) Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function

        End If

    End If

    'el pj es aspirante al clan y puede entrar
    
    Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
    Call guilds(GI).AceptarNuevoMiembro(Aspirante)
    
    ' If player is online, update tag
    If AspiranteUI > 0 Then
        Call RefreshCharStatus(AspiranteUI)

    End If
    
    a_AceptarAspirante = True

End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
    GuildName = guilds(GuildIndex).GuildName

End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
    GuildLeader = guilds(GuildIndex).GetLeader

End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)

End Function

Public Function GuildFounder(ByVal GuildIndex As Integer) As String

    '***************************************************
    'Autor: ZaMa
    'Returns the guild founder's name
    'Last Modification: 25/03/2009
    '***************************************************
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
    GuildFounder = guilds(GuildIndex).Fundador

End Function

Public Function GetUserGuildMember(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso
    'Returns the guilds the user has been member of
    '***************************************************
    If Not Database_Enabled Then
        GetUserGuildMember = GetUserGuildMemberCharfile(UserName)
    Else
        GetUserGuildMember = GetUserGuildMemberDatabase(UserName)

    End If

End Function

Public Function GetUserGuildAspirant(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Returns the guilds the user has been member of
    '***************************************************
    If Not Database_Enabled Then
        GetUserGuildAspirant = GetUserGuildAspirantCharfile(UserName)
    Else
        GetUserGuildAspirant = GetUserGuildAspirantDatabase(UserName)

    End If

End Function

Public Function GetUserGuildRejectionReason(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Returns the reason why the user has not been accepted to the guild
    '***************************************************
    If Not Database_Enabled Then
        GetUserGuildRejectionReason = GetUserGuildRejectionReasonCharfile(UserName)
    Else
        GetUserGuildRejectionReason = GetUserGuildRejectionReasonDatabase(UserName)

    End If

End Function

Public Function GetUserGuildPedidos(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Returns the guilds the user asked to be a member of
    '***************************************************
    If Not Database_Enabled Then
        GetUserGuildPedidos = GetUserGuildPedidosCharfile(UserName)
    Else
        GetUserGuildPedidos = GetUserGuildPedidosDatabase(UserName)

    End If

End Function

Public Sub SaveUserGuildRejectionReason(ByVal UserName As String, ByVal Reason As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the rection reason for the user
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserGuildRejectionReasonCharfile(UserName, Reason)
    Else
        Call SaveUserGuildRejectionReasonDatabase(UserName, Reason)

    End If

End Sub

Public Sub SaveUserGuildIndex(ByVal UserName As String, ByVal GuildIndex As Integer)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the guild index
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserGuildIndexCharfile(UserName, GuildIndex)
    Else
        Call SaveUserGuildIndexDatabase(UserName, GuildIndex)

    End If

End Sub

Public Sub SaveUserGuildAspirant(ByVal UserName As String, ByVal AspirantIndex As Integer)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the guild Aspirant index
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserGuildAspirantCharfile(UserName, AspirantIndex)
    Else
        Call SaveUserGuildAspirantDatabase(UserName, AspirantIndex)

    End If

End Sub

Public Sub SaveUserGuildMember(ByVal UserName As String, ByVal guilds As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the guilds the user has been member of
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserGuildMemberCharfile(UserName, guilds)
    Else
        Call SaveUserGuildMemberDatabase(UserName, guilds)

    End If

End Sub

Public Sub SaveUserGuildPedidos(ByVal UserName As String, ByVal Pedidos As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the guilds the user has asked to be a member of
    '***************************************************
    If Not Database_Enabled Then
        Call SaveUserGuildPedidosCharfile(UserName, Pedidos)
    Else
        Call SaveUserGuildPedidosDatabase(UserName, Pedidos)

    End If

End Sub
