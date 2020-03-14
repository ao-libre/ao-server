Attribute VB_Name = "Areas"
'Como podrán ver la estructura del modulo es la misma, pero algunos procedimientos cambiaron. _
   Con este nuevo módulo, pueden usarse mapas de 100x100, 200x200, 300x300 ,etc (puede hacerse un mundo contínuo si se quiere). _
    ¿Y cómo?, eso es tan facil como modificar las constantes XMAXMAPSIZE e YMAXMAPSIZE. El resto se calcula automáticamente. _
      Cabe mencionar que este módulo trae solucionado el problema de los clones en resoluciones mayores a 800x600!. _
       Seguramente pueden observar que las funciones para enviar datos al area son "raras" pq devuelven un valor q _
        puede variar entre 0, 1, o 2. Eso pueden cambiarlo a gusto, yo lo tengo así por una modificación q lo hice al _
         sistema de caminata (si tocan eso, posteriormente, van a tener que ajustar los subs del mod senddata).
 
Option Explicit

'Tamano del mapa
Public Const XMaxMapSize        As Byte = 100
Public Const XMinMapSize        As Byte = 1
Public Const YMaxMapSize        As Byte = 100
Public Const YMinMapSize        As Byte = 1

'Tamano del tileset
Public Const TileSizeX          As Byte = 32
Public Const TileSizeY          As Byte = 32

'Tamano en Tiles de la pantalla de visualizacion
Public Const XWindow            As Byte = 23
Public Const YWindow            As Byte = 17

Public Const RANGO_VISION_X     As Byte = 11
Public Const RANGO_VISION_Y     As Byte = 9

Private Const USER_NUEVO         As Byte = 255

Private Const tiles_Area_X      As Byte = 11
Private Const tiles_Area_Y      As Byte = 11

Public ConnGroups()             As ConnGroup
Public idAreaX()                As Integer
Public idAreaY()                As Integer

Private CurDay                  As Byte
Private CurHour                 As Byte

Public Type AreaInfo
    AreaPerteneceX              As Integer
    AreaPerteneceY              As Integer
End Type

Public Type ConnGroup
    CountEntrys                 As Long
    OptValue                    As Long
    UserEntrys()                As Long
End Type

'Generamos areas específicas tanto para X como para Y
Public Sub generarIDAreas()
    
    Dim i As Long, count As Integer, tempArea As Integer

    ReDim idAreaX(0 To XMaxMapSize) As Integer
    ReDim idAreaY(0 To YMaxMapSize) As Integer

    For i = 0 To (XMaxMapSize)

        If (i / tiles_Area_X) = CInt(i / tiles_Area_X) Then
            tempArea = tempArea + 1
            count = 0
        End If
        
        idAreaX(i) = tempArea
        count = count + 1
    
    Next i

    tempArea = 0
    count = 0
   
    For i = 0 To (YMaxMapSize)

        If (i / tiles_Area_Y) = CInt(i / tiles_Area_Y) Then
            tempArea = tempArea + 1
            count = 0
        End If
        
        idAreaY(i) = tempArea
        count = count + 1
    
    Next i
  
    'Esto lo mantenemos del sistema de areas anterior
    CurDay = IIf(Weekday(Date) > 6, 1, 2)
    CurHour = Fix(Hour(time) \ 3)
          
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
 
    For i = 1 To NumMaps
        ConnGroups(i).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & i, CurDay & "-" & CurHour))
    
        If ConnGroups(i).OptValue = 0 Then ConnGroups(i).OptValue = 1
       
        ReDim ConnGroups(i).UserEntrys(1 To ConnGroups(i).OptValue) As Long

    Next i
  
End Sub

'JAO; 20/12/17; el siguiente procedimiento calcula los mínimos y máximos tiles para enviar al cliente
Private Sub direccionadoAreas(ByVal X As Integer, _
                             ByVal Y As Integer, _
                             ByVal Head As Byte, _
                             ByRef stX As Integer, _
                             ByRef lastX As Integer, _
                             ByRef stY As Integer, _
                             ByRef lastY As Integer)
  
    Dim subAreaX As Single, subAreaY As Single, pAreaX As Integer, pAreaY As Integer, mArraysX() As String, mArraysY() As String
    
    If Head = eHeading.WEST Then X = X + 1 'lo pasamos directamente al area vecina para obtener enteros
    If Head = eHeading.NORTH Then Y = Y + 1
    
    subAreaX = X / tiles_Area_X 'Con esto obtenemos el ID del area en que estamos, junto con la posición en que estamos dentro del area (ID - > nº entero, posición nuestra dentro del area -> decimales)
    mArraysX = Split(subAreaX, ",") 'Separamos el AreaID de la posición del AreaIndex
     
    subAreaY = Y / tiles_Area_Y
    mArraysY = Split(subAreaY, ",")
     
    If CInt(subAreaX) <> subAreaX Then 'Si el AreaID no es entero, es porque no estamos en el 1er tile del area
        mArraysX(1) = mArraysX(1) * ((0.1 ^ Len(mArraysX(1)))) 'Cálculos del Sr Contador
        pAreaX = mArraysX(1) * tiles_Area_X 'JAO: ¿En qué parte del area estamos situados? (en términos de tiles)
    Else
        pAreaX = 0 'JAO: Si el número es entero, es porque estamos en la posición inicial del area
    End If
     
    If CInt(subAreaY) <> subAreaY Then 'Idem con respecto a subAreaX
        mArraysY(1) = mArraysY(1) * ((0.1 ^ Len(mArraysY(1))))
        pAreaY = mArraysY(1) * tiles_Area_Y
    Else
        pAreaY = 0
    End If
     
    'De aquí en adelante se fijan minimos y maximos en x e y para enviar la informacion correspondiente
    If Head = eHeading.WEST Then ' Vuelta hacia la izquierda
        stX = X - (tiles_Area_X - pAreaX) - (tiles_Area_X)
        lastX = X + (tiles_Area_X - pAreaX) + (tiles_Area_X)
      
        stY = Y - (tiles_Area_Y - pAreaY) - (tiles_Area_Y)
        lastY = Y + (tiles_Area_Y - pAreaY) + (tiles_Area_Y)
      
    ElseIf Head = eHeading.EAST Then 'Vuelta hacia la derecha
        stX = X - (tiles_Area_X - pAreaX) - (tiles_Area_X)
        lastX = X + (tiles_Area_X - pAreaX) + (tiles_Area_X)
      
        stY = Y - (tiles_Area_Y - pAreaY) - (tiles_Area_Y)
        lastY = Y + (tiles_Area_Y - pAreaY) + (tiles_Area_Y)
      
    ElseIf Head = eHeading.NORTH Then 'Vuelta hacia arriba
        stX = X - (tiles_Area_X - pAreaX) - (tiles_Area_X)
        lastX = X + (tiles_Area_X - pAreaX) + (tiles_Area_X)
      
        stY = Y - (tiles_Area_Y - pAreaY) - (tiles_Area_Y)
        lastY = Y + (tiles_Area_Y - pAreaY) + (tiles_Area_Y)
                    
    ElseIf Head = eHeading.SOUTH Then 'Vuelta hacia abajo
        stX = X - (tiles_Area_X - pAreaX) - (tiles_Area_X)
        lastX = X + (tiles_Area_X - pAreaX) + (tiles_Area_X)
      
        stY = Y - (tiles_Area_Y - pAreaY) - (tiles_Area_Y)
        lastY = Y + (tiles_Area_Y - pAreaY) + (tiles_Area_Y)
      
    Else 'Esto ocurre cuando head = user_nuevo (cambio de mapa o logueo de un pj, entonces, enviamos más info)
        stX = X - (tiles_Area_X * 2)
        lastX = X + (tiles_Area_X * 2)
      
        stY = Y - (tiles_Area_Y * 2)
        lastY = Y + (tiles_Area_Y * 2)
     
    End If

      'verificamos que todo este dentro de los parámetros posibles...
       If stX < tiles_Area_X Then stX = tiles_Area_X
       If lastX > XMaxMapSize Then lastX = XMaxMapSize
       If stY < tiles_Area_Y Then stY = tiles_Area_Y
       If lastY > YMaxMapSize Then lastY = YMaxMapSize
                         
                         
End Sub

'Esto queda prácticamente igual, solo que se eliminó el hardcode feo que había
Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, Optional ByVal ButIndex As Boolean = False, Optional verInvis As Byte = 0)

    If UserList(UserIndex).AreasInfo.AreaPerteneceX = idAreaX(UserList(UserIndex).Pos.X) And UserList(UserIndex).AreasInfo.AreaPerteneceY = idAreaY(UserList(UserIndex).Pos.Y) Then Exit Sub

    Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer, TempInt As Long, Map As Long
 
    With UserList(UserIndex)
 
        Call direccionadoAreas(.Pos.X, .Pos.Y, Head, MinX, MaxX, MinY, MaxY)
    
        Map = .Pos.Map
        Call WriteAreaChanged(UserIndex)
    
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY
                
                '<<< User >>>
                If MapData(Map, X, Y).UserIndex Then
                    
                    TempInt = MapData(Map, X, Y).UserIndex
                    
                    If UserIndex <> TempInt Then
                        
                        ' Solo avisa al otro cliente si no es un admin invisible
                        If Not (UserList(TempInt).flags.AdminInvisible = 1) Then
                            Call MakeUserChar(False, UserIndex, TempInt, Map, X, Y)
                            
                            'Si el user estaba invisible le avisamos al nuevo cliente de eso
                            If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
                                If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
                                    Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                                End If
                            End If
                        End If
                        
                        ' Solo avisa al otro cliente si no es un admin invisible
                        If Not (UserList(UserIndex).flags.AdminInvisible = 1) Then
                            Call MakeUserChar(False, TempInt, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
                            
                            If UserList(UserIndex).flags.invisible Or UserList(UserIndex).flags.Oculto Then
                                If UserList(TempInt).flags.Privilegios And PlayerType.User Then
                                    Call WriteSetInvisible(TempInt, UserList(UserIndex).Char.CharIndex, True)
                                End If
                            End If
                        End If
                        
                        Call FlushBuffer(TempInt)
                    
                    ElseIf Head = USER_NUEVO Then
                        Call MakeUserChar(False, UserIndex, UserIndex, Map, X, Y)
                    End If
                End If
                
                '<<< Npc >>>
                If MapData(Map, X, Y).NpcIndex Then
                    Call MakeNPCChar(False, UserIndex, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                End If
             
                'Objs
                If MapData(Map, X, Y).ObjInfo.ObjIndex Then
                    TempInt = MapData(Map, X, Y).ObjInfo.ObjIndex
                    If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                        Call WriteObjectCreate(UserIndex, ObjData(TempInt).GrhIndex, X, Y)
                    
                        If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, UserIndex, X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(False, UserIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                        End If
                    End If
                End If
        
            Next Y
        Next X
    
        .AreasInfo.AreaPerteneceX = idAreaX(.Pos.X)
        .AreasInfo.AreaPerteneceY = idAreaY(.Pos.Y)
    
    End With
End Sub

'IMPORTANTE --!--, dejé esto sin with porque con los hilos se crashea todo al carajo
Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
    If Npclist(NpcIndex).AreasInfo.AreaPerteneceX = idAreaX(Npclist(NpcIndex).Pos.X) And Npclist(NpcIndex).AreasInfo.AreaPerteneceY = idAreaY(Npclist(NpcIndex).Pos.Y) Then Exit Sub

    Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer, TempInt As Long
 
    Call direccionadoAreas(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Head, MinX, MaxX, MinY, MaxY)

         If MapInfo(Npclist(NpcIndex).Pos.Map).NumUsers <> 0 Then
        
            For X = MinX To MaxX
                For Y = MinY To MaxY
                    If MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex Then _
                       Call MakeNPCChar(False, MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex, NpcIndex, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
                Next Y
            Next X
         End If
    
    Npclist(NpcIndex).AreasInfo.AreaPerteneceX = idAreaX(Npclist(NpcIndex).Pos.X)
    Npclist(NpcIndex).AreasInfo.AreaPerteneceY = idAreaY(Npclist(NpcIndex).Pos.Y)
 
End Sub

Public Sub AreasOptimizacion()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
'**************************************************************
    Dim LoopC As Long
    Dim tCurDay As Byte
    Dim tCurHour As Byte
    Dim EntryValue As Long
 
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(time) \ 3)) Then
    
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
    
        For LoopC = 1 To NumMaps
            EntryValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) \ 2))
        
            ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, tCurDay & "-" & tCurHour))
            If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
            If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).NumUsers Then ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
        Next LoopC
    
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub

Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
On Error GoTo ErrorHandler

    Dim TempVal As Long
    Dim LoopC As Long
 
    'Search for the user
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(LoopC) = UserIndex Then Exit For
    Next LoopC
 
    'Char not found
    If LoopC > ConnGroups(Map).CountEntrys Then Exit Sub
 
    'Remove from old map
    ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1
    TempVal = ConnGroups(Map).CountEntrys
 
    'Move list back
    For LoopC = LoopC To TempVal
        ConnGroups(Map).UserEntrys(LoopC) = ConnGroups(Map).UserEntrys(LoopC + 1)
    Next LoopC
 
    If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
    End If
 
    Exit Sub
 
ErrorHandler:
 
    Dim UserName As String
    If UserIndex > 0 Then UserName = UserList(UserIndex).Name

    Call LogError("Error en QuitarUser " & Err.Number & ": " & Err.description & _
                  ". User: " & UserName & "(" & UserIndex & ")")

End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer, Optional ByVal ButIndex As Boolean = False)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 04/01/2007
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'   - Now the method checks for repetead users instead of trusting parameters.
'   - If the character is new to the map, update it
'**************************************************************

    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i As Long
 
    If Not MapaValido(Map) Then Exit Sub
 
    EsNuevo = True
 
    'Prevent adding repeated users
    For i = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(i) = UserIndex Then
            EsNuevo = False
            Exit For
        End If
    Next i
 
    If EsNuevo Then
        'Update map and connection groups data
        ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1
        TempVal = ConnGroups(Map).CountEntrys
 
        If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
        End If
    
        ConnGroups(Map).UserEntrys(TempVal) = UserIndex
    End If
 
    With UserList(UserIndex)
        .AreasInfo.AreaPerteneceX = 0
        .AreasInfo.AreaPerteneceY = 0
    End With
 
    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, ButIndex)
End Sub

Public Sub AgregarNpc(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    With Npclist(NpcIndex)
        .AreasInfo.AreaPerteneceX = 0
        .AreasInfo.AreaPerteneceY = 0
    End With
 
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub

'JAO: De aquí en adelante, se verifica si otro user o npc está en nuestra area o area vecina para enviar datos
Public Function estoyAreaUser(ByVal U1 As Integer, ByVal U2 As Integer) As Long

    If UserList(U1).AreasInfo.AreaPerteneceX = UserList(U2).AreasInfo.AreaPerteneceX Or _
       UserList(U1).AreasInfo.AreaPerteneceX = UserList(U2).AreasInfo.AreaPerteneceX + 1 Or _
       UserList(U1).AreasInfo.AreaPerteneceX = UserList(U2).AreasInfo.AreaPerteneceX - 1 Then estoyAreaUser = estoyAreaUser + 1
    
    If UserList(U1).AreasInfo.AreaPerteneceY = UserList(U2).AreasInfo.AreaPerteneceY Or _
       UserList(U1).AreasInfo.AreaPerteneceY = UserList(U2).AreasInfo.AreaPerteneceY + 1 Or _
       UserList(U1).AreasInfo.AreaPerteneceY = UserList(U2).AreasInfo.AreaPerteneceY - 1 Then estoyAreaUser = estoyAreaUser + 1

End Function

Public Function npcAreaNpc(ByVal U1 As Integer, ByVal U2 As Integer) As Long

    If Npclist(U1).AreasInfo.AreaPerteneceX = Npclist(U2).AreasInfo.AreaPerteneceX Or _
       Npclist(U1).AreasInfo.AreaPerteneceX = Npclist(U2).AreasInfo.AreaPerteneceX + 1 Or _
       Npclist(U1).AreasInfo.AreaPerteneceX = Npclist(U2).AreasInfo.AreaPerteneceX - 1 Then npcAreaNpc = npcAreaNpc + 1
     
    If Npclist(U1).AreasInfo.AreaPerteneceY = Npclist(U2).AreasInfo.AreaPerteneceY Or _
       Npclist(U1).AreasInfo.AreaPerteneceY = Npclist(U2).AreasInfo.AreaPerteneceY + 1 Or _
       Npclist(U1).AreasInfo.AreaPerteneceY = Npclist(U2).AreasInfo.AreaPerteneceY - 1 Then npcAreaNpc = npcAreaNpc + 1

End Function

Public Function npcAreaUser(ByVal U1 As Integer, ByVal U2 As Integer) As Long

    If Npclist(U1).AreasInfo.AreaPerteneceX = UserList(U2).AreasInfo.AreaPerteneceX Or _
       Npclist(U1).AreasInfo.AreaPerteneceX = UserList(U2).AreasInfo.AreaPerteneceX + 1 Or _
       Npclist(U1).AreasInfo.AreaPerteneceX = UserList(U2).AreasInfo.AreaPerteneceX - 1 Then npcAreaUser = npcAreaUser + 1
     
    If Npclist(U1).AreasInfo.AreaPerteneceY = UserList(U2).AreasInfo.AreaPerteneceY Or _
       Npclist(U1).AreasInfo.AreaPerteneceY = UserList(U2).AreasInfo.AreaPerteneceY + 1 Or _
       Npclist(U1).AreasInfo.AreaPerteneceY = UserList(U2).AreasInfo.AreaPerteneceY - 1 Then npcAreaUser = npcAreaUser + 1

End Function

Public Function toUserArea(ByVal U1 As Integer, ByVal X As Integer, ByVal Y As Integer) As Long
    
    If UserList(U1).AreasInfo.AreaPerteneceX = X Or _
       UserList(U1).AreasInfo.AreaPerteneceX = X + 1 Or _
       UserList(U1).AreasInfo.AreaPerteneceX = X - 1 Then toUserArea = toUserArea + 1
    
    If UserList(U1).AreasInfo.AreaPerteneceY = Y Or _
       UserList(U1).AreasInfo.AreaPerteneceY = Y + 1 Or _
       UserList(U1).AreasInfo.AreaPerteneceY = Y - 1 Then toUserArea = toUserArea + 1

Debug.Print "toUserArea: " & toUserArea

End Function

