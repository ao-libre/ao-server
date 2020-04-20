Attribute VB_Name = "Areas"
'************************************************
'*   Sistema de areas refactorizado por WyroX   *
'*       Menos brujeria, mas comentarios.       *
'************************************************

'************************************************
'*                Como funciona?                *
'************************************************
' El mapa se divide en una cuadricula con secto-
' res de igual tamanio. El objetivo es que cada
' usuario reciba solo los paquetes de otros usua-
' rios, NPCs y objetos correspondientes a su area,
' es decir, su cuadricula y las 8 adyacentes.
' Los arrays de usuarios por cada mapa tienen un
' tamanio inicial predeterminado que se calcula
' en base a la cantidad de usuarios que frecuenta
' el mapa.
' El servidor solo se encarga de enviar los nuevos
' elementos del area y el cliente se encarga de
' eliminar los que ya no correspondan.
 
Option Explicit



'************************************************
'*            Valores modificables              *
'************************************************

' Tamanio del mapa
Public Const XMaxMapSize        As Byte = 100
Public Const XMinMapSize        As Byte = 1
Public Const YMaxMapSize        As Byte = 100
Public Const YMinMapSize        As Byte = 1

' Tamanio en tiles de la pantalla.
'ADVERTENCIA: TIENEN QUE SER IMPAR!
Public Const XWindow            As Byte = 23
Public Const YWindow            As Byte = 19

' Cantidad de tiles buffer
' (para que graficos grandes se vean desde fuera de la pantalla)
Private Const TileBufferSize    As Byte = 5

'************************************************
'*      Valores calculados automaticamente      *
'************************************************

' Rangos de vision
Public Const RANGO_VISION_X     As Byte = XWindow \ 2
Public Const RANGO_VISION_Y     As Byte = YWindow \ 2

' Tamanio de las areas
Public Const AREAS_X            As Byte = RANGO_VISION_X + TileBufferSize
Public Const AREAS_Y            As Byte = RANGO_VISION_Y + TileBufferSize



'************************************************
'*               Otras variables                *
'************************************************

Private AreasIO As clsIniManager

Private FILE_AREAS As String

Private Const USER_NUEVO        As Byte = 255

Public ConnGroups()             As Collection

Public Type AreaInfo
    AreaPerteneceX              As Integer
    AreaPerteneceY              As Integer
End Type

Public Sub InitializeAreas()
    Dim i As Long
    ReDim ConnGroups(1 To NumMaps) As Collection
    For i = 1 To NumMaps
        Set ConnGroups(i) = New Collection
    Next i
End Sub

'*****************************************************************************************
'* AgregarUser: Agrega el usuario al mapa, enviando los datos correspondientes a su area *
'               y notificando al resto de usuarios.                                      *
'*****************************************************************************************
Public Sub AgregarUser(ByVal Userindex As Integer, ByVal Map As Integer, Optional ByVal ButIndex As Boolean = False)
    If Not MapaValido(Map) Then Exit Sub

    Dim EsNuevo As Boolean
        EsNuevo = True
    
    ' Evitamos agregar usuarios repetidos
    Dim i As Integer
    For i = 1 To ConnGroups(Map).Count()
        If ConnGroups(Map).Item(i) = Userindex Then
            EsNuevo = False
            Exit For
        End If
    Next i

    ' Si es nuevo en el mapa
    If EsNuevo Then
        Call ConnGroups(Map).Add(Userindex)
    End If
    
    With UserList(Userindex)
        .AreasInfo.AreaPerteneceX = -1
        .AreasInfo.AreaPerteneceY = -1
    End With
    Call CheckUpdateNeededUser(Userindex, USER_NUEVO, ButIndex)
    
End Sub

'************************************************************************************
'* AgregarNpc: agrega el npc al mapa, notificando a los usuarios dentro de su area. *
'************************************************************************************
Public Sub AgregarNpc(ByVal NpcIndex As Integer)

    With Npclist(NpcIndex)
        .AreasInfo.AreaPerteneceX = -1
        .AreasInfo.AreaPerteneceY = -1
    End With

    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
    
End Sub

'*****************************************************************************
'* QuitarUser: remueve el usuario del array del mapa en el que se encuentra. *
'*****************************************************************************
Public Sub QuitarUser(ByVal Userindex As Integer, ByVal Map As Integer)

    ' Buscamos el index dentro del array
    Dim LoopA As Long
    For LoopA = 1 To ConnGroups(Map).Count()
        If ConnGroups(Map).Item(LoopA) = Userindex Then
            Call ConnGroups(Map).Remove(LoopA)
            Exit For
        End If
    Next LoopA

End Sub

'***************************************************************************************************************
'* CheckUpdateNeededUser: Comprueba si es necesario modificar el area del usuario,                             *
'                         de ser asi, le envia todos los datos nuevos y avisa a los demas usuarios de la zona. *
'***************************************************************************************************************
Public Sub CheckUpdateNeededUser(ByVal Userindex As Integer, ByVal heading As Byte, Optional ByVal ButIndex As Boolean = False, Optional verInvis As Byte = 0)

    With UserList(Userindex)

        ' Comprobamos si cambio de area
        If .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X And _
           .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y Then _
                Exit Sub

        Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer, CurUser As Long, Map As Long

        ' Calculamos segun la direccion del usuario el area nueva que tenemos que mandarle
        Call CalcularNuevaArea(.Pos.X, .Pos.Y, heading, MinX, MaxX, MinY, MaxY)

        ' Avisamos al cliente para que borre todo lo que esta fuera del area
        Call WriteAreaChanged(Userindex)

        Map = .Pos.Map

        For X = MinX To MaxX
            For Y = MinY To MaxY

                '<<< User >>>
                If MapData(Map, X, Y).Userindex Then

                    CurUser = MapData(Map, X, Y).Userindex

                    ' No nos enviamos a nosotros mismos...
                    If Userindex <> CurUser Then

                        ' No vemos admins invisibles
                        If Not (UserList(CurUser).flags.AdminInvisible = 1) Then
                            ' Creamos el char del usuario
                            Call MakeUserChar(False, Userindex, CurUser, Map, X, Y)

                            ' Enviamos la invisibilidad de ser necesario
                            If UserList(CurUser).flags.Navegando = 0 Then
                                If UserList(CurUser).flags.invisible Or UserList(CurUser).flags.Oculto Then
                                    If UserList(Userindex).flags.Privilegios And PlayerType.User Then
                                        Call WriteSetInvisible(Userindex, UserList(CurUser).Char.CharIndex, True)
                                    End If
                                End If
                            End If
                        End If

                        ' Si no somos un admin invisible
                        If Not (.flags.AdminInvisible = 1) Then
                            ' Enviamos nuestro char al usuario
                            Call MakeUserChar(False, CurUser, Userindex, .Pos.Map, .Pos.X, .Pos.Y)
                            
                            If .flags.Navegando = 0 Then
                                ' Enviamos la invisibilidad de ser necesario
                                If .flags.invisible Or .flags.Oculto Then
                                    If UserList(CurUser).flags.Privilegios And PlayerType.User Then
                                        Call WriteSetInvisible(CurUser, .Char.CharIndex, True)
                                    End If
                                End If
                            
                            Else
                                Call WriteConsoleMsg(CurUser, "No podes hacerte invisible navegando.", FONTTYPE_INFO)
                                
                            End If
                            
                        End If
                        
                    '... excepto que nos hayamos warpeado al mapa y ButIndex = false
                    ElseIf heading = USER_NUEVO And Not ButIndex Then
                        Call MakeUserChar(False, Userindex, Userindex, Map, X, Y)
                        
                        If .flags.AdminInvisible = 1 Or .flags.Navegando = 0 And (.flags.invisible Or .flags.Oculto) Then
                            Call WriteSetInvisible(Userindex, .Char.CharIndex, True)
                        End If
                    End If
                    
                End If

                '<<< Npc >>>
                If MapData(Map, X, Y).NpcIndex Then
                    Call MakeNPCChar(False, Userindex, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                End If

                'Objs
                If MapData(Map, X, Y).ObjInfo.ObjIndex Then
                    CurUser = MapData(Map, X, Y).ObjInfo.ObjIndex
                    If Not EsObjetoFijo(ObjData(CurUser).OBJType) Then
                        Call WriteObjectCreate(Userindex, ObjData(CurUser).GrhIndex, X, Y)

                        If ObjData(CurUser).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, Userindex, X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(False, Userindex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                        End If
                    End If
                End If

            Next Y
        Next X

        .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X
        .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y

    End With
End Sub


'***************************************************************************************************************
'* CheckUpdateNeededNpc: comprueba si el NPC cambio de area y le avisa a todos los usuarios que sea necesario. *
'***************************************************************************************************************
Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal heading As Byte)
    With Npclist(NpcIndex)

        ' Comprobamos si cambio de area
        If .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X And _
           .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y Then _
                Exit Sub

        Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer, Userindex As Long

        ' Calculamos el area nueva segun la direccion del NPC
        Call CalcularNuevaArea(.Pos.X, .Pos.Y, heading, MinX, MaxX, MinY, MaxY)

        ' Si no hay usuarios en el mapa ahorramos tiempo y salimos
        If MapInfo(.Pos.Map).NumUsers <> 0 Then
        
            For X = MinX To MaxX
                For Y = MinY To MaxY
                    ' Si hay un usuario le enviamos el NPC
                    If MapData(.Pos.Map, X, Y).Userindex Then _
                        Call MakeNPCChar(False, MapData(.Pos.Map, X, Y).Userindex, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)
                Next Y
            Next X
        
        End If

        .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X
        .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y

    End With
End Sub

'**************************************************************************************************************************
'* CalcularNuevaArea: segun la posicion actual y la direccion dada, se calcula el area en tiles que debe ser actualizada. *
'**************************************************************************************************************************
Private Sub CalcularNuevaArea(ByVal X As Integer, ByVal Y As Integer, ByVal heading As Byte, ByRef MinX As Integer, ByRef MaxX As Integer, ByRef MinY As Integer, ByRef MaxY As Integer)

    Dim AreaX As Integer, AreaY As Integer
    Dim MinAreaX As Integer, MaxAreaX As Integer, MinAreaY As Integer, MaxAreaY As Integer

    ' Calculamos el area actual al que pertenece
    AreaX = X \ AREAS_X
    AreaY = Y \ AREAS_Y

    ' Calculamos el conjunto de areas nuevas
    Select Case heading
        Case eHeading.NORTH
            ' 3 areas nuevas arriba
            MinAreaX = AreaX - 1
            MinAreaY = AreaY - 1
            MaxAreaX = AreaX + 1
            MaxAreaY = AreaY - 1

        Case eHeading.EAST
            ' 3 areas nuevas a la derecha
            MinAreaX = AreaX + 1
            MinAreaY = AreaY - 1
            MaxAreaX = AreaX + 1
            MaxAreaY = AreaY + 1

        Case eHeading.SOUTH
            ' 3 areas nuevas abajo
            MinAreaX = AreaX - 1
            MinAreaY = AreaY + 1
            MaxAreaX = AreaX + 1
            MaxAreaY = AreaY + 1

        Case eHeading.WEST
            ' 3 areas nuevas a la izquierda
            MinAreaX = AreaX - 1
            MinAreaY = AreaY - 1
            MaxAreaX = AreaX - 1
            MaxAreaY = AreaY + 1

        Case Else ' Heading = USER_NUEVO (cambio de mapa, recien logueado, etc.)
            ' 9 areas nuevas alrededor del usuario (3x3)
            MinAreaX = AreaX - 1
            MinAreaY = AreaY - 1
            MaxAreaX = AreaX + 1
            MaxAreaY = AreaY + 1
    End Select

    ' Convertimos de areas a tiles
    MinX = MinAreaX * AREAS_X
    MinY = MinAreaY * AREAS_Y
    MaxX = (MaxAreaX + 1) * AREAS_X - 1
    MaxY = (MaxAreaY + 1) * AREAS_Y - 1

    ' Comprobamos que este dentro del mapa
    If MinX < XMinMapSize Then MinX = XMinMapSize
    If MinY < YMinMapSize Then MinY = YMinMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize

End Sub

'******************************************************************************************
'* EstanMismoArea: devuelve verdadero si los usuarios deben enviarse paquetes mutuamente. *
'******************************************************************************************
Public Function EstanMismoArea(ByVal UserA As Integer, ByVal UserB As Integer) As Boolean
    EstanMismoArea = Abs(UserList(UserA).AreasInfo.AreaPerteneceX - UserList(UserB).AreasInfo.AreaPerteneceX) <= 1 And _
                     Abs(UserList(UserA).AreasInfo.AreaPerteneceY - UserList(UserB).AreasInfo.AreaPerteneceY) <= 1
End Function

'*****************************************************************************************
'* EstanMismoAreaNPC: devuelve verdadero si el usuario esta en el mismo area que el NPC. *
'*****************************************************************************************
Public Function EstanMismoAreaNPC(ByVal NpcIndex As Integer, ByVal Userindex As Integer) As Boolean
    EstanMismoAreaNPC = Abs(UserList(Userindex).AreasInfo.AreaPerteneceX - Npclist(NpcIndex).AreasInfo.AreaPerteneceX) <= 1 And _
                        Abs(UserList(Userindex).AreasInfo.AreaPerteneceY - Npclist(NpcIndex).AreasInfo.AreaPerteneceY) <= 1
End Function

'**********************************************************************************************
'* EstanMismoAreaPos: devuelve verdadero si el usuario esta en el mismo area que la posicion. *
'**********************************************************************************************
Public Function EstanMismoAreaPos(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    EstanMismoAreaPos = Abs(UserList(Userindex).AreasInfo.AreaPerteneceX - X \ AREAS_X) <= 1 And _
                        Abs(UserList(Userindex).AreasInfo.AreaPerteneceY - Y \ AREAS_Y) <= 1
End Function
