VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsClan.cls
'
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

''
' clase clan
'
' Es el "ADO" de los clanes. La interfaz entre el disco y
' el juego. Los datos no se guardan en memoria
' para evitar problemas de sincronizacion, y considerando
' que la performance de estas rutinas NO es critica.
' by el oso :p

Private p_GuildName                 As String
Private p_Alineacion                As ALINEACION_GUILD
Private p_OnlineMembers             As Collection   'Array de UserIndexes!
Private p_GMsOnline                 As Collection
Private p_PropuestasDePaz           As Collection
Private p_PropuestasDeAlianza       As Collection
Private p_IteradorRelaciones        As Integer
Private p_IteradorOnlineMembers     As Integer
Private p_IteradorPropuesta         As Integer
Private p_IteradorOnlineGMs         As Integer
Private p_GuildNumber               As Integer      'Numero de guild en el mundo
Private p_Relaciones()              As RELACIONES_GUILD 'array de relaciones con los otros clanes
Private GUILDINFOFILE               As String
Private GUILDPATH                   As String       'aca pq me es mas comodo setearlo y pq en ningun disenio
Private MEMBERSFILE                 As String       'decente la capa de arriba se entera donde estan
Private SOLICITUDESFILE             As String       'los datos fisicamente
Private PROPUESTASFILE              As String
Private RELACIONESFILE              As String
Private VOTACIONESFILE              As String

Private Const NEWSLENGTH = 1024
Private Const DESCLENGTH = 256
Private Const CODEXLENGTH = 256

Public Property Get GuildName() As String
    GuildName = p_GuildName
End Property


'
'ALINEACION Y ANTIFACCION
'

Public Property Get Alineacion() As ALINEACION_GUILD
    Alineacion = p_Alineacion
End Property

Public Property Get PuntosAntifaccion() As Integer
    PuntosAntifaccion = val(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Antifaccion"))
End Property

Public Property Let PuntosAntifaccion(ByVal p As Integer)
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Antifaccion", CStr(p))
End Property

Public Function CambiarAlineacion(ByVal NuevaAlineacion As ALINEACION_GUILD) As Boolean
    p_Alineacion = NuevaAlineacion
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Alineacion", Alineacion2String(p_Alineacion))
    
    If p_Alineacion = ALINEACION_NEUTRO Then CambiarAlineacion = True
End Function

'
'INICIALIZADORES
'

Private Sub Class_Initialize()
    GUILDPATH = App.Path & "\GUILDS\"
    GUILDINFOFILE = GUILDPATH & "guildsinfo.inf"
End Sub

Private Sub Class_Terminate()
    Set p_OnlineMembers = Nothing
    Set p_GMsOnline = Nothing
    Set p_PropuestasDePaz = Nothing
    Set p_PropuestasDeAlianza = Nothing
End Sub



Public Sub Inicializar(ByVal GuildName As String, ByVal GuildNumber As Integer, ByVal Alineacion As ALINEACION_GUILD)
Dim i As Integer

    p_GuildName = GuildName
    p_GuildNumber = GuildNumber
    p_Alineacion = Alineacion
    Set p_OnlineMembers = New Collection
    Set p_GMsOnline = New Collection
    Set p_PropuestasDePaz = New Collection
    Set p_PropuestasDeAlianza = New Collection
    'ALLIESFILE = GUILDPATH & p_GuildName & "-Allied.all"
    'ENEMIESFILE = GUILDPATH & p_GuildName & "-enemys.ene"
    RELACIONESFILE = GUILDPATH & p_GuildName & "-relaciones.rel"
    MEMBERSFILE = GUILDPATH & p_GuildName & "-members.mem"
    PROPUESTASFILE = GUILDPATH & p_GuildName & "-propositions.pro"
    SOLICITUDESFILE = GUILDPATH & p_GuildName & "-solicitudes.sol"
    VOTACIONESFILE = GUILDPATH & p_GuildName & "-votaciones.vot"
    p_IteradorOnlineMembers = 0
    p_IteradorPropuesta = 0
    p_IteradorOnlineGMs = 0
    p_IteradorRelaciones = 0
    ReDim Preserve p_Relaciones(1 To CANTIDADDECLANES) As RELACIONES_GUILD
    For i = 1 To CANTIDADDECLANES
        p_Relaciones(i) = String2Relacion(GetVar(RELACIONESFILE, "RELACIONES", CStr(i)))
    Next i
    For i = 1 To CANTIDADDECLANES
        If Trim$(GetVar(PROPUESTASFILE, CStr(i), "Pendiente")) = "1" Then
            Select Case String2Relacion(Trim$(GetVar(PROPUESTASFILE, CStr(i), "Tipo")))
                Case RELACIONES_GUILD.ALIADOS
                    p_PropuestasDeAlianza.Add i
                Case RELACIONES_GUILD.PAZ
                    p_PropuestasDePaz.Add i
            End Select
        End If
    Next i
End Sub

''
' esta TIENE QUE LLAMARSE LUEGO DE INICIALIZAR()
'
' @param Fundador Nombre del fundador del clan
'
Public Sub InicializarNuevoClan(ByRef Fundador As String)
Dim OldQ    As String  'string pq al comienzo quizas no hay archivo guildinfo.ini y oldq es ""
Dim NewQ    As Integer
    'para que genere los archivos
    Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", "0")
    Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", "0")


    OldQ = GetVar(GUILDINFOFILE, "INIT", "nroguilds")
    If IsNumeric(OldQ) Then
        NewQ = CInt(Trim$(OldQ)) + 1
    Else
        NewQ = 1
    End If

    Call WriteVar(GUILDINFOFILE, "INIT", "NroGuilds", NewQ)

    Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Founder", Fundador)
    Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "GuildName", p_GuildName)
    Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Date", Date)
    Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Antifaccion", "0")
    Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Alineacion", Alineacion2String(p_Alineacion))

End Sub

Public Sub ProcesarFundacionDeOtroClan()
    ReDim Preserve p_Relaciones(1 To CANTIDADDECLANES) As RELACIONES_GUILD
    p_Relaciones(CANTIDADDECLANES) = PAZ
End Sub


'
'MEMBRESIAS
'

Public Property Get Fundador() As String
    Fundador = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Founder")
End Property

'Public Property Get JugadoresOnline() As String
'Dim i As Integer
'    'leve violacion de capas x aqui, je
'    For i = 1 To p_OnlineMembers.Count
'        JugadoresOnline = UserList(p_OnlineMembers.Item(i)).Name & "," & JugadoresOnline
'    Next i
'End Property

Public Property Get CantidadDeMiembros() As Integer
Dim OldQ As String
    OldQ = GetVar(MEMBERSFILE, "INIT", "NroMembers")
    CantidadDeMiembros = IIf(IsNumeric(OldQ), CInt(OldQ), 0)
End Property

Public Sub SetLeader(ByRef leader As String)
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Leader", leader)
End Sub

Public Function GetLeader() As String
    GetLeader = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Leader")
End Function

Public Function GetMemberList() As String()
    Dim OldQ    As Integer
    Dim list() As String
    Dim i As Long
    
    OldQ = Me.CantidadDeMiembros
    
    ReDim list(OldQ - 1) As String
    
    For i = 1 To OldQ
        list(i - 1) = UCase$(GetVar(MEMBERSFILE, "Members", "Member" & i))
    Next i
    
    GetMemberList = list
End Function

Public Sub ConectarMiembro(ByVal UserIndex As Integer)
    p_OnlineMembers.Add UserIndex

    With UserList(UserIndex)
        Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & " se ha conectado."))
    End With
End Sub

Public Sub DesConectarMiembro(ByVal UserIndex As Integer)
Dim i   As Integer
    For i = 1 To p_OnlineMembers.Count
        If p_OnlineMembers.Item(i) = UserIndex Then
            p_OnlineMembers.Remove i
            
            With UserList(UserIndex)
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & " se ha desconectado."))
            End With

            Exit Sub
        End If
    Next i
    

End Sub

Public Sub AceptarNuevoMiembro(ByRef Nombre As String)
Dim OldQ    As Integer
Dim OldQs   As String
Dim ruta    As String

ruta = CharPath & Nombre & ".chr"
    If FileExist(ruta) Then
        Call WriteVar(ruta, "GUILD", "GUILDINDEX", p_GuildNumber)
        Call WriteVar(ruta, "GUILD", "AspiranteA", "0")
        'CantPs = GetVar(CharPath & Nombre & ".chr", "GUILD", "ClanesParticipo")
        'If IsNumeric(CantPs) Then
        '    CantP = CInt(CantPs)
        'Else
        '    CantP = 0
        'End If
        'Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "ClanesParticipo", CantP + 1)
        OldQs = GetVar(MEMBERSFILE, "INIT", "NroMembers")
        If IsNumeric(OldQs) Then
            OldQ = CInt(OldQs)
        Else
            OldQ = 0
        End If
        Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", OldQ + 1)
        Call WriteVar(MEMBERSFILE, "Members", "Member" & OldQ + 1, Nombre)
    End If

End Sub

Public Sub ExpulsarMiembro(ByRef Nombre As String)
Dim OldQ        As Integer
Dim Temps       As String
Dim i           As Integer
Dim EsMiembro   As Boolean
Dim MiembroDe   As String

    If LenB(dir$(CharPath & Nombre & ".chr")) <> 0 Then
        OldQ = CInt(GetVar(MEMBERSFILE, "INIT", "NroMembers"))
        i = 1
        Nombre = UCase$(Nombre)
        While i <= OldQ And UCase$(Trim$(GetVar(MEMBERSFILE, "Members", "Member" & i))) <> Nombre
            i = i + 1
        Wend
        EsMiembro = i <= OldQ
        
        If EsMiembro Then
            Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "GuildIndex", vbNullString)
            While i < OldQ
                Temps = GetVar(MEMBERSFILE, "Members", "Member" & i + 1)
                Call WriteVar(MEMBERSFILE, "Members", "Member" & i, Temps)
                i = i + 1
            Wend
            Call WriteVar(MEMBERSFILE, "Members", "Member" & OldQ, vbNullString)
            'seteo la cantidad de miembros nueva
            Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", OldQ - 1)
            'lo echo a el
            MiembroDe = GetVar(CharPath & Nombre & ".chr", "GUILD", "Miembro")
            If Not InStr(1, MiembroDe, p_GuildName, vbTextCompare) > 0 Then
                If LenB(MiembroDe) <> 0 Then
                    MiembroDe = MiembroDe & ","
                End If
                MiembroDe = MiembroDe & p_GuildName
                Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "Miembro", MiembroDe)
            End If
        End If

    End If

End Sub

'
'ASPIRANTES
'

Public Function GetAspirantes() As String()
    Dim OldQ    As Integer
    Dim list() As String
    Dim i As Long
    
    OldQ = Me.CantidadAspirantes()
    
    If OldQ > 1 Then
        ReDim list(OldQ - 1) As String
    Else
        ReDim list(0) As String
    End If
    
    For i = 1 To OldQ
        list(i - 1) = GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre")
    Next i
    
    GetAspirantes = list
End Function

Public Function CantidadAspirantes() As Integer
Dim Temps As String

    CantidadAspirantes = 0
    Temps = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")
    If Not IsNumeric(Temps) Then
        Exit Function
    End If
    CantidadAspirantes = CInt(Temps)

End Function

Public Function DetallesSolicitudAspirante(ByVal NroAspirante As Integer) As String
    DetallesSolicitudAspirante = GetVar(SOLICITUDESFILE, "SOLICITUD" & NroAspirante, "Detalle")
End Function

Public Function NumeroDeAspirante(ByRef Nombre As String) As Integer
Dim i       As Integer
    
    NumeroDeAspirante = 0
    
    For i = 1 To MAXASPIRANTES
        If UCase$(Trim$(GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre"))) = UCase$(Nombre) Then
            NumeroDeAspirante = i
            Exit Function
        End If
    Next i
End Function

Public Sub NuevoAspirante(ByRef Nombre As String, ByRef Peticion As String)
Dim i       As Integer
Dim OldQ    As String
Dim OldQI   As Integer

    OldQ = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")
    If IsNumeric(OldQ) Then
        OldQI = CInt(OldQ)
    Else
        OldQI = 0
    End If
    For i = 1 To MAXASPIRANTES
        If GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre") = vbNullString Then
            Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre", Nombre)
            Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Detalle", IIf(Trim$(Peticion) = vbNullString, "Peticion vacia", Peticion))
            Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", OldQI + 1)
            Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "ASPIRANTEA", p_GuildNumber)
            Exit Sub
        End If
    Next i
End Sub

Public Sub RetirarAspirante(ByRef Nombre As String, ByRef NroAspirante As Integer)
Dim OldQ    As String
Dim OldQI   As String
Dim Pedidos As String
Dim i       As Integer

    OldQ = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")
    If IsNumeric(OldQ) Then
        OldQI = CInt(OldQ)
    Else
        OldQI = 1
    End If
    'Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & NroAspirante, "Nombre", vbNullString)
    'Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & NroAspirante, "Detalle", vbNullString)
    Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "ASPIRANTEA", "0")
    Pedidos = GetVar(CharPath & Nombre & ".chr", "GUILD", "Pedidos")
    If Not InStr(1, Pedidos, p_GuildName, vbTextCompare) > 0 Then
        If LenB(Pedidos) <> 0 Then
            Pedidos = Pedidos & ","
        End If
        Pedidos = Pedidos & p_GuildName
        Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "Pedidos", Pedidos)
    End If
    
    Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", OldQI - 1)
    For i = NroAspirante To modGuilds.MAXASPIRANTES - 1
        Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre", GetVar(SOLICITUDESFILE, "SOLICITUD" & (i + 1), "Nombre"))
        Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Detalle", GetVar(SOLICITUDESFILE, "SOLICITUD" & (i + 1), "Detalle"))
    Next i
    
    Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & modGuilds.MAXASPIRANTES, "Nombre", vbNullString)
    Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & modGuilds.MAXASPIRANTES, "Detalle", vbNullString)
    
End Sub

Public Sub InformarRechazoEnChar(ByRef Nombre As String, ByRef Detalles As String)
    Call WriteVar(CharPath & Nombre & ".chr", "GUILD", "MotivoRechazo", Detalles)
End Sub

'
'DEFINICION DEL CLAN (CODEX Y NOTICIAS)
'

Public Function GetFechaFundacion() As String
    GetFechaFundacion = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Date")
End Function

Public Sub SetCodex(ByVal CodexNumber As Integer, ByRef codex As String)
    Call ReplaceInvalidChars(codex)
    codex = Left$(codex, CODEXLENGTH)
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Codex" & CodexNumber, codex)
End Sub

Public Function GetCodex(ByVal CodexNumber As Integer) As String
    GetCodex = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Codex" & CodexNumber)
End Function


Public Sub SetURL(ByRef URL As String)
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "URL", Left$(URL, 40))
End Sub

Public Function GetURL() As String
    GetURL = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "URL")
End Function

Public Sub SetGuildNews(ByRef News As String)
    Call ReplaceInvalidChars(News)
    
    News = Left$(News, NEWSLENGTH)
    
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "GuildNews", News)
End Sub

Public Function GetGuildNews() As String
    GetGuildNews = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "GuildNews")
End Function

Public Sub SetDesc(ByRef desc As String)
    Call ReplaceInvalidChars(desc)
    desc = Left$(desc, DESCLENGTH)
    
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Desc", desc)
End Sub

Public Function GetDesc() As String
    GetDesc = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Desc")
End Function

'
'
'ELECCIONES
'
'

Public Function EleccionesAbiertas() As Boolean
Dim ee As String
    ee = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas")
    EleccionesAbiertas = (ee = "1")     'cualquier otra cosa da falso
End Function

Public Sub AbrirElecciones()
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas", "1")
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan", DateAdd("d", 1, Now))
    Call WriteVar(VOTACIONESFILE, "INIT", "NumVotos", "0")
End Sub

Private Sub CerrarElecciones()  'solo pueden cerrarse mediante recuento de votos
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas", "0")
    Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan", vbNullString)
    Call Kill(VOTACIONESFILE)   'borramos toda la evidencia ;-)
End Sub

Public Sub ContabilizarVoto(ByRef Votante As String, ByRef Votado As String)
Dim q       As Integer
Dim Temps   As String

    Temps = GetVar(VOTACIONESFILE, "INIT", "NumVotos")
    q = IIf(IsNumeric(Temps), CInt(Temps), 0)
    Call WriteVar(VOTACIONESFILE, "VOTOS", Votante, Votado)
    Call WriteVar(VOTACIONESFILE, "INIT", "NumVotos", CStr(q + 1))
End Sub

Public Function YaVoto(ByRef Votante) As Boolean
    YaVoto = ((LenB(Trim$(GetVar(VOTACIONESFILE, "VOTOS", Votante)))) <> 0)
End Function

Private Function ContarVotos(ByRef CantGanadores As Integer) As String
Dim q           As Integer
Dim i           As Integer
Dim Temps       As String
Dim tempV       As String
Dim d           As diccionario

On Error GoTo errh
    ContarVotos = vbNullString
    CantGanadores = 0
    Temps = GetVar(MEMBERSFILE, "INIT", "NroMembers")
    q = IIf(IsNumeric(Temps), CInt(Temps), 0)
    If q > 0 Then
        'el diccionario tiene clave el elegido y valor la #votos
        Set d = New diccionario
        
        For i = 1 To q
            'miembro del clan
            Temps = GetVar(MEMBERSFILE, "MEMBERS", "Member" & i)
            
            'a quienvoto
            tempV = GetVar(VOTACIONESFILE, "VOTOS", Temps)
            
            'si voto a alguien contabilizamos el voto
            If LenB(tempV) <> 0 Then
                If Not IsNull(d.At(tempV)) Then  'cuantos votos tiene?
                    Call d.AtPut(tempV, CInt(d.At(tempV)) + 1)
                Else
                    Call d.AtPut(tempV, 1)
                End If
            End If
        Next i
    
        'quien quedo con mas votos, y cuantos tuvieron esos votos?
        ContarVotos = d.MayorValor(CantGanadores)
    
        Set d = Nothing
    End If
    
Exit Function
errh:
    LogError ("clsClan.Contarvotos: " & Err.description)
    If Not d Is Nothing Then Set d = Nothing
    ContarVotos = vbNullString
End Function

Public Function RevisarElecciones() As Boolean
Dim FechaSufragio   As Date
Dim Temps           As String
Dim Ganador         As String
Dim CantGanadores   As Integer
Dim list()          As String
Dim i               As Long

    RevisarElecciones = False
    Temps = Trim$(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan"))
    
    If Temps = vbNullString Then Exit Function
    
    If IsDate(Temps) Then
        FechaSufragio = CDate(Temps)
        If FechaSufragio < Now Then     'toca!
            Ganador = ContarVotos(CantGanadores)

            If CantGanadores > 1 Then
                'empate en la votacion
                Call SetGuildNews("*Empate en la votación. " & Ganador & " con " & CantGanadores & " votos ganaron las elecciones del clan.")
            ElseIf CantGanadores = 1 Then
                list = Me.GetMemberList()
                
                For i = 0 To UBound(list())
                    If Ganador = list(i) Then Exit For
                Next i
                
                If i <= UBound(list()) Then
                    Call SetGuildNews("*" & Ganador & " ganó la elección del clan*")
                    Call Me.SetLeader(Ganador)
                    RevisarElecciones = True
                Else
                    Call SetGuildNews("*" & Ganador & " ganó la elección del clan pero abandonó las filas por lo que la votación queda desierta*")
                End If
            Else
                Call SetGuildNews("*El período de votación se cerró sin votos*")
            End If
            
            Call CerrarElecciones
            
        End If
    Else
        Call LogError("clsClan.RevisarElecciones: tempS is not Date")
    End If

End Function

'/VOTACIONES


'
'RELACIONES
'

Public Property Get CantidadPropuestas(ByVal Tipo As RELACIONES_GUILD) As Integer
    Select Case Tipo
        Case RELACIONES_GUILD.ALIADOS
            CantidadPropuestas = p_PropuestasDeAlianza.Count
        Case RELACIONES_GUILD.GUERRA
        
        Case RELACIONES_GUILD.PAZ
            CantidadPropuestas = p_PropuestasDePaz.Count
    End Select
End Property

Public Property Get CantidadEnemys() As Integer
Dim i As Integer
    For i = 1 To CANTIDADDECLANES
        CantidadEnemys = CantidadEnemys + IIf(p_Relaciones(i) = GUERRA, 1, 0)
    Next i
End Property

Public Property Get CantidadAllies() As Integer
Dim i As Integer
    For i = 1 To CANTIDADDECLANES
        CantidadAllies = CantidadAllies + IIf(p_Relaciones(i) = ALIADOS, 1, 0)
    Next i
End Property

Public Function GetRelacion(ByVal OtroGuild As Integer) As RELACIONES_GUILD
    GetRelacion = p_Relaciones(OtroGuild)
End Function

Public Sub SetRelacion(ByVal GuildIndex As Integer, ByVal Relacion As RELACIONES_GUILD)
    p_Relaciones(GuildIndex) = Relacion
    Call WriteVar(RELACIONESFILE, "RELACIONES", GuildIndex, Relacion2String(Relacion))
End Sub

Public Sub SetPropuesta(ByVal Tipo As RELACIONES_GUILD, ByVal OtroGuild As Integer, ByRef Detalle As String)
    Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle", Detalle)
    Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Tipo", Relacion2String(Tipo))
    Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Pendiente", "1")
    Select Case Tipo
        Case RELACIONES_GUILD.ALIADOS
            p_PropuestasDeAlianza.Add OtroGuild
        Case RELACIONES_GUILD.PAZ
            p_PropuestasDePaz.Add OtroGuild
    End Select
End Sub

Public Sub AnularPropuestas(ByVal OtroGuild As Integer)
Dim i As Integer

    Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle", vbNullString)
    Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Pendiente", "0")
    For i = 1 To p_PropuestasDePaz.Count
        If p_PropuestasDePaz.Item(i) = OtroGuild Then p_PropuestasDePaz.Remove (i)
        Exit Sub
    Next i
    For i = 1 To p_PropuestasDeAlianza.Count
        If p_PropuestasDeAlianza.Item(i) = OtroGuild Then p_PropuestasDeAlianza.Remove (i)
        Exit Sub
    Next i
    
End Sub

Public Function GetPropuesta(ByVal OtroGuild As Integer, ByRef Tipo As RELACIONES_GUILD) As String
'trae la solicitd que haya, no valida si es actual o de que tipo es
    GetPropuesta = GetVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle")
    Tipo = String2Relacion(GetVar(PROPUESTASFILE, CStr(OtroGuild), "Tipo"))
End Function

Public Function HayPropuesta(ByVal OtroGuild As Integer, ByRef Tipo As RELACIONES_GUILD) As Boolean
Dim i As Integer

    HayPropuesta = False
    Select Case Tipo
        Case RELACIONES_GUILD.ALIADOS
            For i = 1 To p_PropuestasDeAlianza.Count
                If p_PropuestasDeAlianza.Item(i) = OtroGuild Then
                    HayPropuesta = True
                End If
            Next i
        Case RELACIONES_GUILD.PAZ
            For i = 1 To p_PropuestasDePaz.Count
                If p_PropuestasDePaz.Item(i) = OtroGuild Then
                    HayPropuesta = True
                End If
            Next i
        Case RELACIONES_GUILD.GUERRA

    End Select

End Function

'Public Function GetEnemy(ByVal EnemyIndex As Integer) As String
'    GetEnemy = GetVar(ENEMIESFILE, "ENEMYS", "ENEMY" & EnemyIndex)
'End Function

'Public Function GetAllie(ByVal AllieIndex As Integer) As String
'    GetAllie = GetVar(ALLIESFILE, "ALLIES", "ALLIE" & AllieIndex)
'End Function


'
'ITERADORES
'

Public Function Iterador_ProximaPropuesta(ByVal Tipo As RELACIONES_GUILD) As Integer
    
    Iterador_ProximaPropuesta = 0
    Select Case Tipo
        Case RELACIONES_GUILD.ALIADOS
            If p_IteradorPropuesta < p_PropuestasDeAlianza.Count Then
                p_IteradorPropuesta = p_IteradorPropuesta + 1
                Iterador_ProximaPropuesta = p_PropuestasDeAlianza.Item(p_IteradorPropuesta)
            End If
            
            If p_IteradorPropuesta >= p_PropuestasDeAlianza.Count Then
                p_IteradorPropuesta = 0
            End If
        Case RELACIONES_GUILD.PAZ
            If p_IteradorPropuesta < p_PropuestasDePaz.Count Then
                p_IteradorPropuesta = p_IteradorPropuesta + 1
                Iterador_ProximaPropuesta = p_PropuestasDePaz.Item(p_IteradorPropuesta)
            End If
            
            If p_IteradorPropuesta >= p_PropuestasDePaz.Count Then
                p_IteradorPropuesta = 0
            End If
    End Select
    
End Function

Public Function m_Iterador_ProximoUserIndex() As Integer
    
    If p_IteradorOnlineMembers < p_OnlineMembers.Count Then
        p_IteradorOnlineMembers = p_IteradorOnlineMembers + 1
        m_Iterador_ProximoUserIndex = p_OnlineMembers.Item(p_IteradorOnlineMembers)
    Else
        p_IteradorOnlineMembers = 0
        m_Iterador_ProximoUserIndex = 0
    End If
End Function

Public Function Iterador_ProximoGM() As Integer
    
    If p_IteradorOnlineGMs < p_GMsOnline.Count Then
        p_IteradorOnlineGMs = p_IteradorOnlineGMs + 1
        Iterador_ProximoGM = p_GMsOnline.Item(p_IteradorOnlineGMs)
    Else
        p_IteradorOnlineGMs = 0
        Iterador_ProximoGM = 0
    End If
End Function

Public Function Iterador_ProximaRelacion(ByVal r As RELACIONES_GUILD) As Integer
    
    While p_IteradorRelaciones < UBound(p_Relaciones)
        
        p_IteradorRelaciones = p_IteradorRelaciones + 1
        If p_Relaciones(p_IteradorRelaciones) = r Then
            Iterador_ProximaRelacion = p_IteradorRelaciones
            Exit Function
        End If
    Wend

    If p_IteradorRelaciones >= UBound(p_Relaciones) Then
        p_IteradorRelaciones = 0
    End If
End Function
'
'
'


'
'ADMINISTRATIVAS
'

Public Sub ConectarGM(ByVal UserIndex As Integer)
    p_GMsOnline.Add UserIndex
End Sub

Public Sub DesconectarGM(ByVal UserIndex As Integer)
Dim i As Integer
    For i = 1 To p_GMsOnline.Count
        If p_GMsOnline.Item(i) = UserIndex Then
            p_GMsOnline.Remove (i)
        End If
    Next i
End Sub



'
'VARIAS, EXTRAS Y DEMASES
'

Private Sub ReplaceInvalidChars(ByRef S As String)
    If InStrB(S, Chr$(13)) <> 0 Then
        S = Replace(S, Chr$(13), vbNullString)
    End If
    If InStrB(S, Chr$(10)) <> 0 Then
        S = Replace(S, Chr$(10), vbNullString)
    End If
    If InStrB(S, "¬") <> 0 Then
        S = Replace(S, "¬", vbNullString)   'morgo usaba esto como "separador"
    End If
End Sub
