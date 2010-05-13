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
'º¬

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE   As String
'archivo .\guilds\guildinfo.ini o similar

Private Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

Private guilds(1 To MAX_GUILDS) As clsClan
'array global de guilds, se indexa por userlist().guildindex

Private Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION As Byte = 5
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
        Set guilds(i) = New clsClan
        TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
        Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
        Call guilds(i).Inicializar(TempStr, i, Alin)
    Next i
    
End Sub

Public Function m_ConectarMiembroAClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************


Dim NuevaA  As Boolean
Dim News    As String

    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(UserIndex)
        UserList(UserIndex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(UserIndex, True, NuevaA)
        If NuevaA Then News = News & "El clan tiene nueva alineación."
        'If NuevoL Or NuevaA Then Call guilds(GuildIndex).SetGuildNews(News)
    End If

End Function


Public Function m_ValidarPermanencia(ByVal UserIndex As Integer, ByVal SumaAntifaccion As Boolean, _
                            ByRef CambioAlineacion As Boolean) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 14/12/2009
'25/03/2009: ZaMa - Desequipo los items faccionarios que tenga el funda al abandonar la faccion
'14/12/2009: ZaMa - La alineacion del clan depende del lider
'14/02/2010: ZaMa - Ya no es necesario saber si el lider cambia, ya que no puede cambiar.
'***************************************************

Dim GuildIndex  As Integer

    m_ValidarPermanencia = True
    
    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    
    If Not m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        
        ' Es el lider, bajamos 1 rango de alineacion
        If GuildLeader(GuildIndex) = UserList(UserIndex).name Then
            Call LogClanes(UserList(UserIndex).name & ", líder de " & guilds(GuildIndex).GuildName & " hizo bajar la alienación de su clan.")
        
            CambioAlineacion = True
            
            ' Por si paso de ser armada/legion a pk/ciuda, chequeo de nuevo
            Do
                Call UpdateGuildMembers(GuildIndex)
            Loop Until m_EstadoPermiteEntrar(UserIndex, GuildIndex)
        Else
            Call LogClanes(UserList(UserIndex).name & " de " & guilds(GuildIndex).GuildName & " es expulsado en validar permanencia.")
        
            m_ValidarPermanencia = False
            If SumaAntifaccion Then guilds(GuildIndex).PuntosAntifaccion = guilds(GuildIndex).PuntosAntifaccion + 1
            
            CambioAlineacion = guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
            
            Call LogClanes(UserList(UserIndex).name & " de " & guilds(GuildIndex).GuildName & _
                IIf(CambioAlineacion, " SI ", " NO ") & "provoca cambio de alineación. MAXANT:" & CambioAlineacion)
            
            Call m_EcharMiembroDeClan(-1, UserList(UserIndex).name)
            
            ' Llegamos a la maxima cantidad de antifacciones permitidas, bajamos un grado de alineación
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
    Dim TotalMembers As Integer
    Dim MemberIndex As Long
    Dim Sale As Boolean
    Dim MemberName As String
    Dim UserIndex As Integer
    Dim Reenlistadas As Integer
    
    ' Si devuelve true, cambio a neutro y echamos a todos los que estén de mas, sino no echamos a nadie
    If guilds(GuildIndex).CambiarAlineacion(BajarGrado(GuildIndex)) Then 'ALINEACION_NEUTRO)
        
        'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo alterar
        'internamente al iterador en el proceso
        GuildMembers = guilds(GuildIndex).GetMemberList()
        TotalMembers = UBound(GuildMembers)
        
        For MemberIndex = 0 To TotalMembers
            MemberName = GuildMembers(MemberIndex)
            
            'vamos a violar un poco de capas..
            UserIndex = NameIndex(MemberName)
            If UserIndex > 0 Then
                Sale = Not m_EstadoPermiteEntrar(UserIndex, GuildIndex)
            Else
                Sale = Not m_EstadoPermiteEntrarChar(MemberName, GuildIndex)
            End If

            If Sale Then
                If m_EsGuildLeader(MemberName, GuildIndex) Then  'hay que sacarlo de las facciones
                 
                    If UserIndex > 0 Then
                        If UserList(UserIndex).Faccion.ArmadaReal <> 0 Then
                            Call ExpulsarFaccionReal(UserIndex)
                            ' No cuenta como reenlistada :p.
                            UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas - 1
                        ElseIf UserList(UserIndex).Faccion.FuerzasCaos <> 0 Then
                            Call ExpulsarFaccionCaos(UserIndex)
                            ' No cuenta como reenlistada :p.
                            UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas - 1
                        End If
                    Else
                        If FileExist(CharPath & MemberName & ".chr") Then
                            Call WriteVar(CharPath & MemberName & ".chr", "FACCIONES", "EjercitoCaos", 0)
                            Call WriteVar(CharPath & MemberName & ".chr", "FACCIONES", "EjercitoReal", 0)
                            Reenlistadas = GetVar(CharPath & MemberName & ".chr", "FACCIONES", "Reenlistadas")
                            Call WriteVar(CharPath & MemberName & ".chr", "FACCIONES", "Reenlistadas", _
                                    IIf(Reenlistadas > 1, Reenlistadas - 1, Reenlistadas))
                        End If
                    End If
                Else    'sale si no es guildLeader
                    Call m_EcharMiembroDeClan(-1, MemberName)
                End If
            End If
        Next MemberIndex
    Else
        ' Resetea los puntos de antifacción
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

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If UserList(UserIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(UserIndex)
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

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
                Call guilds(GI).DesConectarMiembro(UserIndex)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(UserIndex).GuildIndex = 0
                Call RefreshCharStatus(UserIndex)
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

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim GI As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then Exit Sub
    
    Call guilds(GI).SetURL(Web)
    
End Sub


Public Sub ChangeCodexAndDesc(ByRef desc As String, ByRef codex() As String, ByVal GuildIndex As Integer)
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

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
'***************************************************
'Author: Unknown
'Last Modification: 21/02/2010
'21/02/2010: ZaMa - Ahora le avisa a los miembros que cambio el guildnews.
'***************************************************

    Dim GI As Integer

    With UserList(UserIndex)
        GI = .GuildIndex
        
        If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
        
        If Not m_EsGuildLeader(.name, GI) Then Exit Sub
        
        Call guilds(GI).SetGuildNews(Datos)
        
        Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & " ha actualizado las noticias del clan!"))
    End With
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef desc As String, ByRef GuildName As String, ByRef URL As String, ByRef codex() As String, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim CantCodex       As Integer
Dim i               As Integer
Dim DummyString     As String

    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan inválido."
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
            Call .InicializarNuevoClan(UserList(FundadorIndex).name)
            
            'seteamos codex y descripcion
            For i = 1 To CantCodex
                Call .SetCodex(i, codex(i - 1))
            Next i
            Call .SetDesc(desc)
            Call .SetGuildNews("Clan creado con alineación: " & Alineacion2String(Alineacion))
            Call .SetLeader(UserList(FundadorIndex).name)
            Call .SetURL(URL)
            
            '"conectamos" al nuevo miembro a la lista de la clase
            Call .AceptarNuevoMiembro(UserList(FundadorIndex).name)
            Call .ConectarMiembro(FundadorIndex)
        End With
        
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call RefreshCharStatus(FundadorIndex)
        
        For i = 1 To CANTIDADDECLANES - 1
            Call guilds(i).ProcesarFundacionDeOtroClan
        Next i
    Else
        refError = "No hay más slots para fundar clanes. Consulte a un administrador."
        Exit Function
    End If
    
    CrearNuevoClan = True
End Function

Public Sub SendGuildNews(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim GuildIndex  As Integer
Dim i               As Integer
Dim go As Integer

    GuildIndex = UserList(UserIndex).GuildIndex
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
    
        Call WriteGuildNews(UserIndex, .GetGuildNews, enemies, allies)
    
        If .EleccionesAbiertas Then
            Call WriteConsoleMsg(UserIndex, "Hoy es la votación para elegir un nuevo líder para el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(UserIndex, "La elección durará 24 horas, se puede votar a cualquier miembro del clan.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(UserIndex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(UserIndex, "Sólo se computará un voto por miembro. Tu voto no puede ser cambiado.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With

End Sub

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
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
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If

    ' Ahora el lider es el unico que no puede salir del clan
    m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).GetLeader) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 27/11/2009
'Returns true if can Found a guild
'27/11/2009: ZaMa - Ahora valida si ya fundo clan o no.
'***************************************************
    
    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.ELV < 25 Or UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 90 Then
        refError = "Para fundar un clan debes ser nivel 25 y tener 90 skills en liderazgo."
        Exit Function
    End If
    
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            If UserList(UserIndex).Faccion.ArmadaReal <> 1 Then
                refError = "Para fundar un clan real debes ser miembro del ejército real."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            If criminal(UserIndex) Then
                refError = "Para fundar un clan de ciudadanos no debes ser criminal."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            If Not criminal(UserIndex) Then
                refError = "Para fundar un clan de criminales no debes ser ciudadano."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_LEGION
            If UserList(UserIndex).Faccion.FuerzasCaos <> 1 Then
                refError = "Para fundar un clan del mal debes pertenecer a la legión oscura."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_MASTER
            If UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
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
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim Promedio    As Long
Dim ELV         As Integer
Dim f           As Byte

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
    
    If FileExist(CharPath & Personaje & ".chr") Then
        Promedio = CLng(GetVar(CharPath & Personaje & ".chr", "REP", "Promedio"))
        Select Case guilds(GuildIndex).Alineacion
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
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = Not criminal(UserIndex) And _
                    IIf(UserList(UserIndex).Stats.ELV >= 25, UserList(UserIndex).Faccion.ArmadaReal <> 0, True)
        
        Case ALINEACION_GUILD.ALINEACION_LEGION
            m_EstadoPermiteEntrar = criminal(UserIndex) And _
                    IIf(UserList(UserIndex).Stats.ELV >= 25, UserList(UserIndex).Faccion.FuerzasCaos <> 0, True)
        
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            m_EstadoPermiteEntrar = UserList(UserIndex).Faccion.ArmadaReal = 0 And UserList(UserIndex).Faccion.FuerzasCaos = 0
        
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            m_EstadoPermiteEntrar = Not criminal(UserIndex)
        
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            m_EstadoPermiteEntrar = criminal(UserIndex)
        
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
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i   As Integer

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
Dim i As Long
Dim name As String

name = UCase$(UserName)

For i = 1 To CANTIDADDECLANES
    HasFound = (UCase$(guilds(i).Fundador) = name)
    If HasFound Then Exit Function
Next i

End Function

Public Function v_AbrirElecciones(ByVal UserIndex As Integer, ByRef refError As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim GuildIndex      As Integer

    v_AbrirElecciones = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tú no perteneces a ningún clan."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GuildIndex) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya están abiertas."
        Exit Function
    End If
    
    v_AbrirElecciones = True
    Call guilds(GuildIndex).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal UserIndex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim GuildIndex      As Integer
Dim list()          As String
Dim i As Long

    v_UsuarioVota = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tú no perteneces a ningún clan."
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
        
        
        If .YaVoto(UserList(UserIndex).name) Then
            refError = "Ya has votado, no puedes cambiar tu voto."
            Exit Function
        End If
        
        Call .ContabilizarVoto(UserList(UserIndex).name, Votado)
        v_UsuarioVota = True
    End With

End Function

Public Sub v_RutinaElecciones()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i       As Integer

On Error GoTo errh
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Revisando elecciones", FontTypeNames.FONTTYPE_SERVER))
    For i = 1 To CANTIDADDECLANES
        If Not guilds(i) Is Nothing Then
            If guilds(i).RevisarElecciones Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & guilds(i).GetLeader & " es el nuevo líder de " & guilds(i).GuildName & ".", FontTypeNames.FONTTYPE_SERVER))
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

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'aca si que vamos a violar las capas deliveradamente ya que
'visual basic no permite declarar metodos de clase
Dim Temps   As String
    If InStrB(PlayerName, "\") <> 0 Then
        PlayerName = Replace(PlayerName, "\", vbNullString)
    End If
    If InStrB(PlayerName, "/") <> 0 Then
        PlayerName = Replace(PlayerName, "/", vbNullString)
    End If
    If InStrB(PlayerName, ".") <> 0 Then
        PlayerName = Replace(PlayerName, ".", vbNullString)
    End If
    Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "GUILDINDEX")
    If IsNumeric(Temps) Then
        GetGuildIndexFromChar = CInt(Temps)
    Else
        GetGuildIndexFromChar = 0
    End If
End Function

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

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i As Integer
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> UserIndex And ((UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0)) Then _
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).name & ","
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
'Last Modification: -
'
'***************************************************

    Dim tStr() As String
    Dim i As Long
    
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        
        For i = 1 To CANTIDADDECLANES
            tStr(i - 1) = guilds(i).GuildName
        Next i
    End If
    
    PrepareGuildsList = tStr
End Function

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim codex(CANTIDADMAXIMACODEX - 1)  As String
    Dim GI      As Integer
    Dim i       As Long

    GI = GuildIndex(GuildName)
    If GI = 0 Then Exit Sub
    
    With guilds(GI)
        For i = 1 To CANTIDADMAXIMACODEX
            codex(i - 1) = .GetCodex(i)
        Next i
        
        Call Protocol.WriteGuildDetails(UserIndex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, _
                                    .GetURL, .CantidadDeMiembros, .EleccionesAbiertas, Alineacion2String(.Alineacion), _
                                    .CantidadEnemys, .CantidadAllies, .PuntosAntifaccion & "/" & CStr(MAXANTIFACCION), _
                                    codex, .GetDesc)
    End With
End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
'***************************************************
'Autor: Mariano Barrou (El Oso)
'Last Modification: 12/10/06
'Las Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************
    Dim GI      As Integer
    Dim guildList() As String
    Dim MemberList() As String
    Dim aspirantsList() As String

    With UserList(UserIndex)
        GI = .GuildIndex
        
        guildList = PrepareGuildsList()
        
        If GI <= 0 Or GI > CANTIDADDECLANES Then
            'Send the guild list instead
            Call WriteGuildList(UserIndex, guildList)
            Exit Sub
        End If
        
        MemberList = guilds(GI).GetMemberList()
        
        If Not m_EsGuildLeader(.name, GI) Then
            'Send the guild list instead
            Call WriteGuildMemberInfo(UserIndex, guildList, MemberList)
            Exit Sub
        End If
        
        aspirantsList = guilds(GI).GetAspirantes()
        
        Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList)
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

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
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

Public Function GMEscuchaClan(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim GI As Integer

    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(UserIndex).EscucheClan <> 0 Then
        'Quit listening to previous guild!!
        Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
        guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
        Exit Function
    End If
    
'devuelve el guildindex
    GI = GuildIndex(GuildName)
    If GI > 0 Then
        If UserList(UserIndex).EscucheClan <> 0 Then
            If UserList(UserIndex).EscucheClan = GI Then
                'Already listening to them...
                Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaClan = GI
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
                guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
            End If
        End If
        
        Call guilds(GI).ConectarGM(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = GI
        UserList(UserIndex).EscucheClan = GI
    Else
        Call WriteConsoleMsg(UserIndex, "Error, el clan no existe.", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = 0
    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'el index lo tengo que tener de cuando me puse a escuchar
    UserList(UserIndex).EscucheClan = 0
    Call guilds(GuildIndex).DesconectarGM(UserIndex)
End Sub
Public Function r_DeclararGuerra(ByVal UserIndex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim GI  As Integer
Dim GIG As Integer

    r_DeclararGuerra = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan."
        Exit Function
    End If
    
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ningún clan."
        Exit Function
    End If
    
    GIG = GuildIndex(GuildGuerra)
    If guilds(GI).GetRelacion(GIG) = GUERRA Then
        refError = "Tu clan ya está en guerra con " & GuildGuerra & "."
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


Public Function r_AceptarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan."
        Exit Function
    End If
    
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ningún clan."
        Exit Function
    End If

    GIG = GuildIndex(GuildPaz)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function
    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
        refError = "No estás en guerra con ese clan."
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

Public Function r_RechazarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan."
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan."
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


Public Function r_RechazarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDePaz = 0
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan."
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan."
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

Public Function r_AceptarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan."
        Exit Function
    End If
    
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ningún clan."
        Exit Function
    End If

    GIG = GuildIndex(GuildAllie)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function
    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
        refError = "No estás en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
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


Public Function r_ClanGeneraPropuesta(ByVal UserIndex As Integer, ByRef OtroClan As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim OtroClanGI      As Integer
Dim GI              As Integer

    r_ClanGeneraPropuesta = False
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
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
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan."
        Exit Function
    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If Tipo = RELACIONES_GUILD.PAZ Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
            refError = "No estás en guerra con " & OtroClan
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

Public Function r_VerPropuesta(ByVal UserIndex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim OtroClanGI      As Integer
Dim GI              As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        refError = "No eres el líder de tu clan."
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroGuild)
    
    If Not guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
        refError = "No existe la propuesta solicitada."
        Exit Function
    End If
    
    r_VerPropuesta = guilds(GI).GetPropuesta(OtroClanGI, Tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal UserIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As String()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim GI  As Integer
    Dim i   As Integer
    Dim proposalCount As Integer
    Dim proposals() As String
    
    GI = UserList(UserIndex).GuildIndex
    
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

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal guild As Integer, ByRef Detalles As String)
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
    Call guilds(guild).InformarRechazoEnChar(Aspirante, Detalles)
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
    a_ObtenerRechazoDeChar = GetVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo")
    Call WriteVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo", vbNullString)
End Function

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, ByRef Nombre As String, ByRef refError As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim GI              As Integer
Dim NroAspirante    As Integer

    a_RechazarAspirante = False
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
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

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef Nombre As String) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim GI              As Integer
    Dim NroAspirante    As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
    
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)
 '***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim GI          As Integer
    Dim NroAsp      As Integer
    Dim GuildName   As String
    Dim UserFile    As clsIniReader
    Dim Miembro     As String
    Dim GuildActual As Integer
    Dim list()      As String
    Dim i           As Long
    
    On Error GoTo error
    GI = UserList(UserIndex).GuildIndex
    
    Personaje = UCase$(Personaje)
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No eres el líder de tu clan.", FontTypeNames.FONTTYPE_INFO)
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
            Call Protocol.WriteConsoleMsg(UserIndex, "El personaje no es ni aspirante ni miembro del clan.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    'ahora traemos la info
    
    Set UserFile = New clsIniReader
    
    With UserFile
        .Initialize (CharPath & Personaje & ".chr")
        
        ' Get the character's current guild
        GuildActual = val(.GetValue("GUILD", "GuildIndex"))
        If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
            GuildName = "<" & guilds(GuildActual).GuildName & ">"
        Else
            GuildName = "Ninguno"
        End If
        
        'Get previous guilds
        Miembro = .GetValue("GUILD", "Miembro")
        If Len(Miembro) > 400 Then
            Miembro = ".." & Right$(Miembro, 400)
        End If
        
        Call Protocol.WriteCharacterInfo(UserIndex, Personaje, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), _
                                .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), _
                                .GetValue("STATS", "Banco"), .GetValue("REP", "Promedio"), .GetValue("GUILD", "Pedidos"), _
                                GuildName, Miembro, .GetValue("FACCIONES", "EjercitoReal"), .GetValue("FACCIONES", "EjercitoCaos"), _
                                .GetValue("FACCIONES", "CiudMatados"), .GetValue("FACCIONES", "CrimMatados"))
    End With
    
    Set UserFile = Nothing
    
    Exit Sub
error:
    Set UserFile = Nothing
    If Not (FileExist(CharPath & Personaje & ".chr", vbArchive)) Then
        Call LogError("El usuario " & UserList(UserIndex).name & " (" & UserIndex & _
                    " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
    Else
        Call LogError("[" & Err.Number & "] " & Err.description & " En la rutina SendDetallesPersonaje, por el usuario " & _
                    UserList(UserIndex).name & " (" & UserIndex & " ), pidiendo información sobre el personaje " & Personaje)
    End If
End Sub

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, ByRef clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim ViejoSolicitado     As String
Dim ViejoGuildINdex     As Integer
Dim ViejoNroAspirante   As Integer
Dim NuevoGuildIndex     As Integer

    a_NuevoAspirante = False

    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro."
        Exit Function
    End If
    
    If EsNewbie(UserIndex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function
    End If

    NuevoGuildIndex = GuildIndex(clan)
    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe, avise a un administrador."
        Exit Function
    End If
    
    If Not m_EstadoPermiteEntrar(UserIndex, NuevoGuildIndex) Then
        refError = "Tú no puedes entrar a un clan de alineación " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
        Exit Function
    End If

    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
        Exit Function
    End If

    ViejoSolicitado = GetVar(CharPath & UserList(UserIndex).name & ".chr", "GUILD", "ASPIRANTEA")

    If LenB(ViejoSolicitado) <> 0 Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)
        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(UserIndex).name)
            If ViejoNroAspirante > 0 Then
                Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).name, ViejoNroAspirante)
            End If
        Else
            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If
    End If
    
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

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
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan."
        Exit Function
    End If
    
    AspiranteUI = NameIndex(Aspirante)
    If AspiranteUI > 0 Then
        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = Aspirante & " no puede entrar a un clan de alineación " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    Else
        If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
            refError = Aspirante & " no puede entrar a un clan de alineación " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf GetGuildIndexFromChar(Aspirante) Then
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

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildName = guilds(GuildIndex).GuildName
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildLeader = guilds(GuildIndex).GetLeader
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)
End Function

Public Function GuildFounder(ByVal GuildIndex As Integer) As String
'***************************************************
'Autor: ZaMa
'Returns the guild founder's name
'Last Modification: 25/03/2009
'***************************************************
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildFounder = guilds(GuildIndex).Fundador
End Function
