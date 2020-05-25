Attribute VB_Name = "InvUsuario"
'Argentum Online 0.12.2
'Copyright (C) 2002 Marquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Public Function TieneObjetosRobables(ByVal Userindex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' 22/05/2010: Los items newbies ya no son robables.
    '***************************************************

    '17/09/02
    'Agregue que la funcion se asegure que el objeto no es un barco

    On Error GoTo ErrHandler

    Dim i        As Integer

    Dim ObjIndex As Integer
    
    For i = 1 To UserList(Userindex).CurrentInventorySlots
        ObjIndex = UserList(Userindex).Invent.Object(i).ObjIndex

        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And Not ItemNewbie(ObjIndex)) Then
                TieneObjetosRobables = True
                Exit Function

            End If

        End If

    Next i
    
    Exit Function

ErrHandler:
    Call LogError("Error en TieneObjetosRobables. Error: " & Err.Number & " - " & Err.description)

End Function

Function ClasePuedeUsarItem(ByVal Userindex As Integer, _
                            ByVal ObjIndex As Integer, _
                            Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 01/04/2019
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '08/08/2015: Shak - Hechizos por clase
    '01/04/2019: Recox - Se arreglo la prohibicion de hechizos por clase
    '***************************************************

    On Error GoTo manejador
  
    'Admins can use ANYTHING!
    If UserList(Userindex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then

            Dim i As Integer

            For i = 1 To NUMCLASES

                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(Userindex).Clase Then
              
                    '//Si es un hechizo
                    If ObjData(ObjIndex).OBJType = eOBJType.otPergaminos Then
                        sMotivo = "Tu clase no tiene la habilidad de aprender este hechizo."
                        ClasePuedeUsarItem = False
                        Exit Function
                    Else
                        sMotivo = "Tu clase no puede usar este objeto."
                        ClasePuedeUsarItem = False
                        Exit Function
                    End If

                End If
                
            Next i

        End If

    End If
  
    ClasePuedeUsarItem = True

    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")

End Function

Sub QuitarNewbieObj(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Integer

    With UserList(Userindex)

        For j = 1 To UserList(Userindex).CurrentInventorySlots

            If .Invent.Object(j).ObjIndex > 0 Then
             
                If ObjData(.Invent.Object(j).ObjIndex).Newbie = 1 Then Call QuitarUserInvItem(Userindex, j, MAX_INVENTORY_OBJS)
                Call UpdateUserInv(False, Userindex, j)
        
            End If

        Next j
    
        '[Barrin 17-12-03] Si el usuario dejo de ser Newbie, y estaba en el Newbie Dungeon
        'es transportado a su hogar de origen ;)
        If MapInfo(.Pos.Map).Restringir = eRestrict.restrict_newbie Then
        
            Dim DeDonde As WorldPos
        
            Select Case .Hogar

                Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                    DeDonde = Lindos

                Case eCiudad.cUllathorpe
                    DeDonde = Ullathorpe

                Case eCiudad.cBanderbill
                    DeDonde = Banderbill

                Case Else
                    DeDonde = Nix

            End Select
        
            Call WarpUserChar(Userindex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
    
        End If

        '[/Barrin]
    End With

End Sub

Sub LimpiarInventario(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Integer

    With UserList(Userindex)

        For j = 1 To .CurrentInventorySlots
            .Invent.Object(j).ObjIndex = 0
            .Invent.Object(j).Amount = 0
            .Invent.Object(j).Equipped = 0
        Next j
    
        .Invent.NroItems = 0
    
        .Invent.ArmourEqpObjIndex = 0
        .Invent.ArmourEqpSlot = 0
    
        .Invent.WeaponEqpObjIndex = 0
        .Invent.WeaponEqpSlot = 0
    
        .Invent.CascoEqpObjIndex = 0
        .Invent.CascoEqpSlot = 0
    
        .Invent.EscudoEqpObjIndex = 0
        .Invent.EscudoEqpSlot = 0
    
        .Invent.AnilloEqpObjIndex = 0
        .Invent.AnilloEqpSlot = 0
    
        .Invent.MunicionEqpObjIndex = 0
        .Invent.MunicionEqpSlot = 0
    
        .Invent.BarcoObjIndex = 0
        .Invent.BarcoSlot = 0
    
        .Invent.MochilaEqpObjIndex = 0
        .Invent.MochilaEqpSlot = 0
        
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaEqpSlot = 0
    End With

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal Userindex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
    '***************************************************
    On Error GoTo ErrHandler

    'If Cantidad > 100000 Then Exit Sub

    With UserList(Userindex)

        'SI EL Pjta TIENE ORO LO TIRAMOS
        If (Cantidad > 0) And (Cantidad <= .Stats.Gld) Then

            Dim MiObj As obj

            'info debug
            Dim loops As Integer
            
            'Seguridad Alkon (guardo el oro tirado si supera los 50k)
            If Cantidad > 50000 Then

                Dim j        As Integer

                Dim K        As Integer

                Dim M        As Integer

                Dim Cercanos As String

                M = .Pos.Map

                For j = .Pos.X - 10 To .Pos.X + 10
                    For K = .Pos.Y - 10 To .Pos.Y + 10

                        If InMapBounds(M, j, K) Then
                            If MapData(M, j, K).Userindex > 0 Then
                                Cercanos = Cercanos & UserList(MapData(M, j, K).Userindex).Name & ","

                            End If

                        End If

                    Next K
                Next j

                Call LogDesarrollo(.Name & " tira oro. Cercanos: " & Cercanos)

            End If

            '/Seguridad
            Dim Extra    As Long

            Dim TeniaOro As Long

            TeniaOro = .Stats.Gld

            If Cantidad > 500000 Then 'Para evitar explotar demasiado
                Extra = Cantidad - 500000
                Cantidad = 500000

            End If
            
            Do While (Cantidad > 0)
                
                If Cantidad > MAX_INVENTORY_OBJS And .Stats.Gld > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    Cantidad = Cantidad - MiObj.Amount
                Else
                    MiObj.Amount = Cantidad
                    Cantidad = Cantidad - MiObj.Amount

                End If
    
                MiObj.ObjIndex = iORO
                
                If EsGm(Userindex) Then Call LogGM(.Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)

                Dim AuxPos As WorldPos
                
                If .Clase = eClass.Pirat And .Invent.BarcoObjIndex = 476 Then
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, False)

                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.Gld = .Stats.Gld - MiObj.Amount

                    End If

                Else
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, True)

                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.Gld = .Stats.Gld - MiObj.Amount

                    End If

                End If
                
                'info debug
                loops = loops + 1

                If loops > 100 Then
                    LogError ("Error en tiraroro")
                    Exit Sub

                End If
                
            Loop

            If TeniaOro = .Stats.Gld Then Extra = 0
            If Extra > 0 Then
                .Stats.Gld = .Stats.Gld - Extra

            End If
        
        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)

End Sub

Sub QuitarUserInvItem(ByVal Userindex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    If Slot < 1 Or Slot > UserList(Userindex).CurrentInventorySlots Then Exit Sub
    
    With UserList(Userindex).Invent.Object(Slot)

        If .Amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(Userindex, Slot)

        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad

        'Quedan mas?
        If .Amount <= 0 Then
            UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0

        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, _
                  ByVal Userindex As Integer, _
                  ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim NullObj As UserObj

    Dim LoopC   As Long

    With UserList(Userindex)

        'Actualiza un solo slot
        If Not UpdateAll Then
    
            'Actualiza el inventario
            If .Invent.Object(Slot).ObjIndex > 0 Then
                Call ChangeUserInv(Userindex, Slot, .Invent.Object(Slot))
            Else
                Call ChangeUserInv(Userindex, Slot, NullObj)

            End If
    
        Else
    
            'Actualiza todos los slots
            For LoopC = 1 To .CurrentInventorySlots

                'Actualiza el inventario
                If .Invent.Object(LoopC).ObjIndex > 0 Then
                    Call ChangeUserInv(Userindex, LoopC, .Invent.Object(LoopC))
                Else
                    Call ChangeUserInv(Userindex, LoopC, NullObj)

                End If

            Next LoopC

        End If
    
        Exit Sub

    End With

ErrHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.description)

End Sub

Sub DropObj(ByVal Userindex As Integer, _
            ByVal Slot As Byte, _
            ByVal Num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 11/5/2010
    '11/5/2010 - ZaMa: Arreglo bug que permitia apilar mas de 10k de items.
    '***************************************************

    Dim DropObj As obj

    Dim MapObj  As obj

    With UserList(Userindex)

        If Num > 0 Then
            
            'Validacion para que no podamos tirar nuestra monturas mientras la usamos.
            If .flags.Equitando = 1 And .Invent.MonturaEqpSlot = Slot Then
                Call WriteConsoleMsg(Userindex, "No podes tirar tu montura mientras la estas usando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            'Validacion para que no podamos tirar nuestra mochila mientras la usamos.
            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = Slot Then
                    Call WriteConsoleMsg(Userindex, "No puedes tirar tu alforja o mochila mientras la estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
        
            DropObj.ObjIndex = .Invent.Object(Slot).ObjIndex
        
            If (ItemNewbie(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
                Call WriteConsoleMsg(Userindex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
        
            DropObj.Amount = MinimoInt(Num, .Invent.Object(Slot).Amount)

            'Check objeto en el suelo
            MapObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
            MapObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
        
            If MapObj.ObjIndex = 0 Or MapObj.ObjIndex = DropObj.ObjIndex Then
        
                If MapObj.Amount = MAX_INVENTORY_OBJS Then
                    Call WriteConsoleMsg(Userindex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
            
                If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
                    DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount

                End If
            
                Call MakeObj(DropObj, Map, X, Y)
                Call QuitarUserInvItem(Userindex, Slot, DropObj.Amount)
                Call UpdateUserInv(False, Userindex, Slot)
            
                If ObjData(DropObj.ObjIndex).OBJType = eOBJType.otBarcos Then
                    Call WriteConsoleMsg(Userindex, "ATENCION!! ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_WARNING)

                End If
            
                If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Tiro cantidad:" & Num & " Objeto:" & ObjData(DropObj.ObjIndex).Name)
            
                'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
                'Es un Objeto que tenemos que loguear?
                If ObjData(DropObj.ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.Name & " tiro al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                ElseIf DropObj.Amount > 5000 Then 'Es mucha cantidad? > Subi a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)

                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(DropObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " tiro al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)

                    End If

                End If

            Else
                Call WriteConsoleMsg(Userindex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

End Sub

Sub EraseObj(ByVal Num As Integer, _
             ByVal Map As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With MapData(Map, X, Y)
        .ObjInfo.Amount = .ObjInfo.Amount - Num
    
        If .ObjInfo.Amount <= 0 Then
            .ObjInfo.ObjIndex = 0
            .ObjInfo.Amount = 0
        
            Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))

        End If

    End With

End Sub

Sub MakeObj(ByRef obj As obj, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
    
        With MapData(Map, X, Y)

            If .ObjInfo.ObjIndex = obj.ObjIndex Then
                .ObjInfo.Amount = .ObjInfo.Amount + obj.Amount
            Else
                .ObjInfo = obj
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(obj.ObjIndex).GrhIndex, X, Y))

            End If
            
            'Validamos que sea un objeto valido para borrar del mapa
            Dim IsNotObjFogata As Boolean
            Dim IsNotObjTeleport As Boolean
            Dim IsNotFragua As Boolean
            Dim IsNotYacimientoPez As Boolean
            Dim IsNotYacimiento As Boolean
            Dim IsNotMueble As Boolean
            Dim IsNotArbolElfico As Boolean
            Dim IsNotArbol As Boolean
            Dim IsNotCartel As Boolean
            Dim IsNotBarco As Boolean
            Dim IsNotMontura As Boolean
            Dim IsNotYunque As Boolean
            Dim IsNotManual As Boolean
            Dim IsNotForo As Boolean
            Dim IsNotPuerta As Boolean
            Dim IsValidObjectToClean As Boolean

            IsNotObjFogata = ObjData(obj.ObjIndex).OBJType <> otFogata 
            IsNotObjTeleport = ObjData(obj.ObjIndex).OBJType <> otTeleport 
            IsNotFragua = ObjData(obj.ObjIndex).OBJType <> otFragua 
            IsNotYacimientoPez = ObjData(obj.ObjIndex).OBJType <> otYacimientoPez 
            IsNotYacimiento = ObjData(obj.ObjIndex).OBJType <> otYacimiento 
            IsNotMueble = ObjData(obj.ObjIndex).OBJType <> otMuebles 
            IsNotArbolElfico = ObjData(obj.ObjIndex).OBJType <> otArbolElfico 
            IsNotArbol = ObjData(obj.ObjIndex).OBJType <> otArboles 
            IsNotCartel = ObjData(obj.ObjIndex).OBJType <> otCarteles 
            IsNotBarco = ObjData(obj.ObjIndex).OBJType <> otBarcos 
            IsNotMontura = ObjData(obj.ObjIndex).OBJType <> otMonturas 
            IsNotYunque = ObjData(obj.ObjIndex).OBJType <> otYunque 
            IsNotManual = ObjData(obj.ObjIndex).OBJType <> otManuales 
            IsNotForo = ObjData(obj.ObjIndex).OBJType <> otForos 
            IsNotPuerta = ObjData(obj.ObjIndex).OBJType <> otPuertas 

            If IsNotObjFogata And IsNotObjTeleport And IsNotFragua And IsNotYacimientoPez And IsNotYacimiento And IsNotMueble And IsNotArbolElfico And IsNotArbol And IsNotCartel And IsNotBarco And IsNotMontura And IsNotYunque And IsNotManual And IsNotForo And IsNotPuerta Then 
                IsValidObjectToClean = True
            Else
                IsValidObjectToClean = False
            End If

            '//Agregamos las pos de los objetos
            If IsValidObjectToClean And ItemNoEsDeMapa(ObjData(obj.ObjIndex).OBJType) Then
                Dim xPos As WorldPos

                xPos.Map = Map
                xPos.X = X
                xPos.Y = Y

                Dim IsNotTileCasaTrigger As Boolean
                Dim IsNotTileBajoTecho As Boolean
                Dim IsNotTileBlocked As Boolean

                IsNotTileCasaTrigger = MapData(xPos.Map, xPos.X, xPos.Y).trigger <> eTrigger.CASA 
                IsNotTileBajoTecho = MapData(xPos.Map, xPos.X, xPos.Y).trigger <> eTrigger.BAJOTECHO
                IsNotTileBlocked = MapData(xPos.Map, xPos.X, xPos.Y).Blocked <> 1

                If (IsNotTileCasaTrigger Or IsNotTileBajoTecho) And IsNotTileBlocked Then AgregarObjetoLimpieza xPos

            End If

        End With

    End If

End Sub

Function MeterItemEnInventario(ByVal Userindex As Integer, ByRef MiObj As obj) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 17/12/2019
    'Obtengo el numero de getMaxInventorySlots antes que nada para que funcionen correctamente las mochilas y alforjas (Recox)
    '***************************************************

    On Error GoTo ErrHandler

    Dim Slot As Byte

    With UserList(Userindex)
        'Chequeamos la cantidad maxima de items en inventario para poder activar/desactivar
        'la mochila sin tener que reiniciar el juego
        'Ya que en en el unico lugar donde se setea este valor es en ConnectUser (Recox)
        .CurrentInventorySlots = getMaxInventorySlots(Userindex)
        
        'el user ya tiene un objeto del mismo tipo?
        Slot = 1
        
        Do Until .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then
                Exit Do

            End If

        Loop
            
        'Sino busca un slot vacio
        If Slot > .CurrentInventorySlots Then
            Slot = 1

            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > .CurrentInventorySlots Then
                    Call WriteConsoleMsg(Userindex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_WARNING)
                    MeterItemEnInventario = False
                    Exit Function

                End If

            Loop
            .Invent.NroItems = .Invent.NroItems + 1

        End If
    
        If Slot > MAX_NORMAL_INVENTORY_SLOTS And Slot <= MAX_INVENTORY_SLOTS Then
            If Not ItemSeCae(MiObj.ObjIndex) Then
                Call WriteConsoleMsg(Userindex, "No puedes agarrar objetos especiales y ponerlos directamente en tu " & ObjData(.Invent.MochilaEqpObjIndex).Name & ". Puedes acomodar los items en tu inventario con drag and drop y asi si poder moverlos dentro de tu alforja. Recuerda que items que no se caen normalmente si estan dentro de la mochila caeran", FontTypeNames.FONTTYPE_WARNING)
                MeterItemEnInventario = False
                Exit Function

            End If

        End If

        'Mete el objeto
        If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
        Else
            .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS

        End If

    End With
    
    MeterItemEnInventario = True
           
    Call UpdateUserInv(False, Userindex, Slot)
    
    Exit Function
ErrHandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.Number & " : " & Err.description)

End Function

Sub GetObj(ByVal Userindex As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 18/12/2009
    '18/12/2009: ZaMa - Oro directo a la billetera.
    '***************************************************

    Dim obj    As ObjData

    Dim MiObj  As obj

    Dim ObjPos As String
    
    With UserList(Userindex)

        'Hay algun obj?
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then

            'Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

                Dim X As Integer

                Dim Y As Integer
                
                X = .Pos.X
                Y = .Pos.Y
                
                obj = ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
                MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                MiObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                ' Oro directo a la billetera!
                If obj.OBJType = otOro Then

                    'Calculamos la diferencia con el maximo de oro permitido el cual es el valor de LONG
                    Dim RemainingAmountToMaximumGold As Long
                    RemainingAmountToMaximumGold = 2147483647 - .Stats.Gld

                    If Not .Stats.Gld > 2147483647 And RemainingAmountToMaximumGold >= MiObj.Amount Then
                        .Stats.Gld = .Stats.Gld + MiObj.Amount
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                            
                        Call WriteUpdateGold(Userindex)
                    Else
                        Call WriteConsoleMsg(Userindex, "No puedes juntar este oro por que tendrias mas del maximo disponible (2147483647)", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

                    If MeterItemEnInventario(Userindex, MiObj) Then
                    
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)

                        If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
        
                        'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                        'Es un Objeto que tenemos que loguear?
                        If ObjData(MiObj.ObjIndex).Log = 1 Then
                            ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                            Call LogDesarrollo(.Name & " junto del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                        
                        ElseIf MiObj.Amount > 5000 Then 'Es mucha cantidad?

                            'Si no es de los prohibidos de loguear, lo logueamos.
                            If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                                ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                                Call LogDesarrollo(.Name & " junto del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)

                            End If

                        End If

                    End If

                End If

            End If

        Else
            Call WriteConsoleMsg(Userindex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

Public Sub Desequipar(ByVal Userindex As Integer, ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    'Desequipa el item slot del inventario
    Dim obj As ObjData
    
    With UserList(Userindex)
        With .Invent

            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
                Exit Sub
            ElseIf .Object(Slot).ObjIndex = 0 Then
                Exit Sub

            End If
            
            obj = ObjData(.Object(Slot).ObjIndex)

        End With
        
        Select Case obj.OBJType

            Case eOBJType.otWeapon

                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0

                End With
                
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .WeaponAnim = NingunArma
                        Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                    End With

                End If
            
            Case eOBJType.otFlechas

                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0

                End With
            
            Case eOBJType.otAnillo

                With .Invent
                    .Object(Slot).Equipped = 0
                    .AnilloEqpObjIndex = 0
                    .AnilloEqpSlot = 0

                End With
            
            Case eOBJType.otArmadura

                With .Invent
                    .Object(Slot).Equipped = 0
                    .ArmourEqpObjIndex = 0
                    .ArmourEqpSlot = 0

                End With
                
                Call DarCuerpoDesnudo(Userindex, .flags.Mimetizado = 1)

                With .Char
                    Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                End With
                 
            Case eOBJType.otCasco

                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0

                End With
                
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .CascoAnim = NingunCasco
                        Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                    End With

                End If
            
            Case eOBJType.otEscudo

                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0

                End With
                
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .ShieldAnim = NingunEscudo
                        Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)

                    End With

                End If
            
            Case eOBJType.otMochilas

                With .Invent
                    .Object(Slot).Equipped = 0
                    .MochilaEqpObjIndex = 0
                    .MochilaEqpSlot = 0

                End With
                
                Call InvUsuario.TirarTodosLosItemsEnMochila(Userindex)
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS

        End Select

    End With
    
    Call WriteUpdateUserStats(Userindex)
    Call UpdateUserInv(False, Userindex, Slot)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.description)

End Sub

Function EsUsable(ByVal ObjIndex As Integer)
'*************************************************
'Author: Jopi
'Revisa si el objeto puede ser equipado/usado.
'Todas las opciones no usables estan en un solo case en ves de muchos diferentes (Recox)
'*************************************************
    Dim obj As ObjData
        obj = ObjData(ObjIndex)
    
    Select Case obj.OBJType
    
         Case eOBJType.otArbolElfico, _
              eOBJType.otArboles, _
              eOBJType.otCarteles, _
              eOBJType.otForos, _
              eOBJType.otFragua, _
              eOBJType.otMuebles, _
              eOBJType.otPuertas, _
              eOBJType.otTeleport, _
              eOBJType.otYacimiento, _
              eOBJType.otYacimientoPez, _
              eOBJType.otYunque
         
            EsUsable = False
        
        Case Else
            EsUsable = True
    
    End Select

End Function

Function SexoPuedeUsarItem(ByVal Userindex As Integer, _
                           ByVal ObjIndex As Integer, _
                           Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo ErrHandler
    
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(Userindex).Genero <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(Userindex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True

    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = "Tu genero no puede usar este objeto."
    
    Exit Function
ErrHandler:
    Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal Userindex As Integer, _
                              ByVal ObjIndex As Integer, _
                              Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    If ObjData(ObjIndex).Real = 1 Then
        If Not criminal(Userindex) Then
            FaccionPuedeUsarItem = esArmada(Userindex)
        Else
            FaccionPuedeUsarItem = False

        End If

    ElseIf ObjData(ObjIndex).Caos = 1 Then

        If criminal(Userindex) Then
            FaccionPuedeUsarItem = esCaos(Userindex)
        Else
            FaccionPuedeUsarItem = False

        End If

    Else
        FaccionPuedeUsarItem = True

    End If
    
    If Not FaccionPuedeUsarItem Then sMotivo = "Tu alineacion no puede usar este objeto."

End Function

Sub EquiparInvItem(ByVal Userindex As Integer, ByVal Slot As Byte)
    '*************************************************
    'Author: Unknown
    'Last modified: 03/02/2020
    '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
    '14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
    '03/02/2020: WyroX - Nivel minimo y skill minimo para poder equipar
    '*************************************************

    On Error GoTo ErrHandler

    'Equipa un item del inventario
    Dim obj      As ObjData

    Dim ObjIndex As Integer

    Dim sMotivo  As String
    
    With UserList(Userindex)
        ObjIndex = .Invent.Object(Slot).ObjIndex
        obj = ObjData(ObjIndex)
        
        ' No se pueden usar muebles.
        If Not EsUsable(ObjIndex) Then Exit Sub

        If .flags.Equitando = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes equiparte o desequiparte mientras estas en tu montura!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If obj.Newbie = 1 And Not EsNewbie(Userindex) Then
            Call WriteConsoleMsg(Userindex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Nivel minimo
        If .Stats.ELV < obj.MinLevel Then
            Call WriteConsoleMsg(Userindex, "Necesitas ser nivel " & obj.MinLevel & " para poder equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Skills minimos
        If obj.SkillRequerido Then
            If .Stats.UserSkills(obj.SkillRequerido) < obj.SkillCantidad Then
                Call WriteConsoleMsg(Userindex, "Necesitas " & obj.SkillCantidad & " puntos en " & SkillsNames(obj.SkillRequerido) & " para poder equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
                
        Select Case obj.OBJType

            Case eOBJType.otWeapon

                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(Userindex, Slot)

                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.WeaponAnim = NingunArma
                        Else
                            .Char.WeaponAnim = NingunArma
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                        Exit Sub

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.WeaponEqpSlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.WeaponEqpObjIndex = ObjIndex
                    .Invent.WeaponEqpSlot = Slot
                    
                    'El sonido solo se envia si no lo produce un admin invisible
                    If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = GetWeaponAnim(Userindex, ObjIndex)
                    Else
                        .Char.WeaponAnim = GetWeaponAnim(Userindex, ObjIndex)
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otAnillo

                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(Userindex, Slot)
                        Exit Sub

                    End If
                        
                    'Quitamos el elemento anterior
                    If .Invent.AnilloEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.AnilloEqpSlot)

                    End If
                
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.AnilloEqpObjIndex = ObjIndex
                    .Invent.AnilloEqpSlot = Slot
                        
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otFlechas

                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                        
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(Userindex, Slot)
                        Exit Sub

                    End If
                        
                    'Quitamos el elemento anterior
                    If .Invent.MunicionEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.MunicionEqpSlot)

                    End If
                
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.MunicionEqpObjIndex = ObjIndex
                    .Invent.MunicionEqpSlot = Slot
                        
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otArmadura

                If .flags.Navegando = 1 Then
                    Call WriteConsoleMsg(Userindex, "No podes equiparte o desequiparte vestimentas o armaduras mientras estas navegando.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'Parchesin para que no se saquen una armadura mientras estan en montura y dsp les queda el cuerpo de la armadura y velocidad de montura (Recox)
                If .flags.Equitando = 1 Then
                    Call WriteConsoleMsg(Userindex, "No podes equiparte o desequiparte vestimentas o armaduras mientras estas en tu montura.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And SexoPuedeUsarItem(Userindex, ObjIndex, sMotivo) And CheckRazaUsaRopa(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)
                        Call DarCuerpoDesnudo(Userindex, .flags.Mimetizado = 1)

                        If Not .flags.Mimetizado = 1 Then
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                        Exit Sub

                    End If
            
                    'Quita el anterior
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.ArmourEqpSlot)

                    End If
            
                    'Lo equipa
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ArmourEqpObjIndex = ObjIndex
                    .Invent.ArmourEqpSlot = Slot
                        
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.body = obj.Ropaje
                    Else
                        .Char.body = obj.Ropaje
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                    End If

                    .flags.Desnudo = 0
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otCasco

                If .flags.Navegando = 1 Then Exit Sub
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)

                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.CascoAnim = NingunCasco
                        Else
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                        Exit Sub

                    End If
            
                    'Quita el anterior
                    If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.CascoEqpSlot)

                    End If
            
                    'Lo equipa
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = ObjIndex
                    .Invent.CascoEqpSlot = Slot

                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.CascoAnim = obj.CascoAnim
                    Else
                        .Char.CascoAnim = obj.CascoAnim
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otEscudo

                If .flags.Navegando = 1 Then Exit Sub
                
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)

                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.ShieldAnim = NingunEscudo
                        Else
                            .Char.ShieldAnim = NingunEscudo
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                        Exit Sub

                    End If
             
                    'Quita el anterior
                    If .Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.EscudoEqpSlot)

                    End If
             
                    'Lo equipa
                     
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.EscudoEqpObjIndex = ObjIndex
                    .Invent.EscudoEqpSlot = Slot
                     
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.ShieldAnim = obj.ShieldAnim
                    Else
                        .Char.ShieldAnim = obj.ShieldAnim
                         
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
                 
            Case eOBJType.otMochilas

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If

                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(Userindex, Slot)
                    Exit Sub

                End If

                If .Invent.MochilaEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.MochilaEqpSlot)

                End If

                .Invent.Object(Slot).Equipped = 1
                .Invent.MochilaEqpObjIndex = ObjIndex
                .Invent.MochilaEqpSlot = Slot

        End Select

    End With
    
    'Actualiza
    Call UpdateUserInv(False, Userindex, Slot)
    
    Exit Sub
    
ErrHandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)

End Sub

Private Function CheckRazaUsaRopa(ByVal Userindex As Integer, _
                                  ItemIndex As Integer, _
                                  Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(Userindex)

        'Verifica si la raza puede usar la ropa
        If .raza = eRaza.Humano Or .raza = eRaza.Elfo Or .raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
        Else
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)

        End If
        
        'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
        If (.raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
            CheckRazaUsaRopa = False

        End If

    End With
    
    If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    
    Exit Function
    
ErrHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal Userindex As Integer, ByVal Slot As Byte)
    '*************************************************
    'Author: Unknown
    'Last modified: 03/02/2020
    'Handels the usage of items from inventory box.
    '24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legion.
    '24/01/2007 Pablo (ToxicWaste) - Utilizacion nueva de Barco en lvl 20 por clase Pirata y Pescador.
    '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
    '17/11/2009: ZaMa - Ahora se envia una orientacion de la posicion hacia donde esta el que uso el cuerno.
    '27/11/2009: Budi - Se envia indivualmente cuando se modifica a la Agilidad o la Fuerza del personaje.
    '08/12/2009: ZaMa - Agrego el uso de hacha de madera elfica.
    '10/12/2009: ZaMa - Arreglos y validaciones en todos las herramientas de trabajo.
    '03/02/2020: WyroX - Nivel minimo para poder usar y clase prohibida a pergaminos y barcas
    '*************************************************

    Dim obj      As ObjData

    Dim ObjIndex As Integer

    Dim TargObj  As ObjData

    Dim MiObj    As obj

    Dim sMotivo As String

    With UserList(Userindex)
    
        If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
        obj = ObjData(.Invent.Object(Slot).ObjIndex)
        
        If obj.Newbie = 1 And Not EsNewbie(Userindex) Then
            Call WriteConsoleMsg(Userindex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If obj.OBJType = eOBJType.otWeapon Then
            If obj.proyectil = 1 Then
                
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                If Not IntervaloPermiteUsar(Userindex, False) Then Exit Sub
            Else

                'dagas
                If Not IntervaloPermiteUsar(Userindex) Then Exit Sub

            End If

        Else

            If Not IntervaloPermiteUsar(Userindex) Then Exit Sub

        End If
        
        ObjIndex = .Invent.Object(Slot).ObjIndex
        .flags.TargetObjInvIndex = ObjIndex
        .flags.TargetObjInvSlot = Slot
        
        ' No se pueden usar muebles.
        If Not EsUsable(ObjIndex) Then Exit Sub

        ' Nivel minimo
        If .Stats.ELV < obj.MinLevel Then
            Call WriteConsoleMsg(Userindex, "Necesitas ser nivel " & obj.MinLevel & " para poder usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Select Case obj.OBJType

            Case eOBJType.otUseOnce

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If
        
                'Usa el item
                .Stats.MinHam = .Stats.MinHam + obj.MinHam

                If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
                .flags.Hambre = 0
                Call WriteUpdateHungerAndThirst(Userindex)
                'Sonido
                
                If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
                    Call ReproducirSonido(SendTarget.ToPCArea, Userindex, e_SoundIndex.MORFAR_MANZANA)
                Else
                    Call ReproducirSonido(SendTarget.ToPCArea, Userindex, e_SoundIndex.SOUND_COMIDA)

                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, Slot, 1)
                
                Call UpdateUserInv(False, Userindex, Slot)
        
            Case eOBJType.otOro

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                .Stats.Gld = .Stats.Gld + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                
                Call UpdateUserInv(False, Userindex, Slot)
                Call WriteUpdateGold(Userindex)
                
            Case eOBJType.otWeapon

                If .flags.Equitando = 1 Then
                    Call WriteConsoleMsg(Userindex, "No puedes usar una herramienta mientras estas en tu montura!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If Not .Stats.MinSta > 0 Then
                    Call WriteConsoleMsg(Userindex, "Estas muy cansad" & IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                If ObjData(ObjIndex).proyectil = 1 Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(Userindex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                ElseIf .flags.TargetObj = Lena Then

                    If .Invent.Object(Slot).ObjIndex = DAGA Then
                        If .Invent.Object(Slot).Equipped = 0 Then
                            Call WriteConsoleMsg(Userindex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                            
                        Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, Userindex)

                    End If

                Else
                    
                    Select Case ObjIndex
                    
                        Case CANA_PESCA, RED_PESCA, CANA_PESCA_NEWBIE
                            
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.pesca)  'Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case HACHA_LENADOR, HACHA_LENA_ELFICA, HACHA_LENADOR_NEWBIE
                            
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.Talar)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case PIQUETE_MINERO, PIQUETE_MINERO_NEWBIE
                        
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.Mineria)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case MARTILLO_HERRERO, MARTILLO_HERRERO_NEWBIE
                        
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.Herreria)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case SERRUCHO_CARPINTERO, SERRUCHO_CARPINTERO_NEWBIE
                            
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call EnivarObjConstruibles(Userindex)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case Else ' Las herramientas no se pueden fundir

                            If ObjData(ObjIndex).SkHerreria > 0 Then
                                ' Solo objetos que pueda hacer el herrero
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)

                            End If

                    End Select

                End If
            
            Case eOBJType.otPociones

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If Not IntervaloPermiteGolpeUsar(Userindex, False) Then
                    Call WriteConsoleMsg(Userindex, "Debes esperar unos momentos para tomar otra pocion!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                .flags.TomoPocion = True
                .flags.TipoPocion = obj.TipoPocion
                        
                Select Case .flags.TipoPocion
                
                    Case 1 'Modif la agilidad
                        .flags.DuracionEfecto = obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS

                        If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If

                        Call WriteUpdateDexterity(Userindex)
                        
                    Case 2 'Modif la fuerza
                        .flags.DuracionEfecto = obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS

                        If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If

                        Call WriteUpdateStrenght(Userindex)
                        
                    Case 3 'Pocion roja, restaura HP
                        'Usa el item
                        .Stats.MinHp = .Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)

                        If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If
                    
                    Case 4 'Pocion azul, restaura MANA
                        'Usa el item
                        'nuevo calculo para recargar mana
                        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV

                        If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If
                        
                    Case 5 ' Pocion violeta

                        If .flags.Envenenado = 1 Then
                            .flags.Envenenado = 0
                            Call WriteConsoleMsg(Userindex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)

                        End If

                        'Quitamos del inv el item
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If
                        
                    Case 6  ' Pocion Negra
                        If .flags.SlotReto > 0 Then Exit Sub
                        
                        If .flags.Privilegios And PlayerType.User Then
                            Call QuitarUserInvItem(Userindex, Slot, 1)
                            Call UserDie(Userindex)
                            Call WriteConsoleMsg(Userindex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)

                        End If

                End Select

                Call WriteUpdateUserStats(Userindex)
                Call UpdateUserInv(False, Userindex, Slot)
        
            Case eOBJType.otBebidas

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If

                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed

                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(Userindex)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                End If
                
                Call UpdateUserInv(False, Userindex, Slot)
            
            Case eOBJType.otLlaves

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If .flags.TargetObj = 0 Then Exit Sub
                TargObj = ObjData(.flags.TargetObj)

                'El objeto clickeado es una puerta?
                If TargObj.OBJType = eOBJType.otPuertas Then

                    'Esta cerrada?
                    If TargObj.Cerrada = 1 Then

                        'Cerrada con llave?
                        If TargObj.Llave > 0 Then
                            If TargObj.Clave = obj.Clave Then
                 
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Call WriteConsoleMsg(Userindex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            Else
                                Call WriteConsoleMsg(Userindex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                        Else

                            If TargObj.Clave = obj.Clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                Call WriteConsoleMsg(Userindex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Exit Sub
                            Else
                                Call WriteConsoleMsg(Userindex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                        End If

                    Else
                        Call WriteConsoleMsg(Userindex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If
            
            Case eOBJType.otBotellaVacia

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If

                If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
                    Call WriteConsoleMsg(Userindex, "No hay agua alli.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
                Call QuitarUserInvItem(Userindex, Slot, 1)

                If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If
                
                Call UpdateUserInv(False, Userindex, Slot)
            
            Case eOBJType.otBotellaLlena

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If

                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed

                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(Userindex)
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
                Call QuitarUserInvItem(Userindex, Slot, 1)

                If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If
                
                Call UpdateUserInv(False, Userindex, Slot)
            
            Case eOBJType.otPergaminos

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If .Stats.MaxMAN > 0 Then
                    If .flags.Hambre = 0 And .flags.Sed = 0 Then

                        If Not ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                            Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If

                        Call AgregarHechizo(Userindex, Slot)
                        Call UpdateUserInv(False, Userindex, Slot)
                    Else
                        Call WriteConsoleMsg(Userindex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)

                End If

            Case eOBJType.otMinerales

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If

                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
               
            Case eOBJType.otInstrumentos

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If obj.Real Then 'Es el Cuerno Real?
                    If FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                        If MapInfo(.Pos.Map).Pk = False Then
                            Call WriteConsoleMsg(Userindex, "No hay peligro aqui. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(Userindex)
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))

                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(Userindex, "Solo miembros del ejercito real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                ElseIf obj.Caos Then 'Es el Cuerno Legion?

                    If FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                        If MapInfo(.Pos.Map).Pk = False Then
                            Call WriteConsoleMsg(Userindex, "No hay peligro aqui. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(Userindex)
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))

                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(Userindex, "Solo miembros de la legion oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If

                'Si llega aca es porque es o Laud o Tambor o Flauta
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))

                End If
               
            Case eOBJType.otBarcos

                If Not ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) Or Not FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'Verifica si esta aproximado al agua antes de permitirle navegar
                If .Stats.ELV < 25 Then

                    ' Solo pirata y trabajador pueden navegar antes
                    If .Clase <> eClass.Worker And .Clase <> eClass.Pirat Then
                        Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    Else

                        ' Pero a partir de 20
                        If .Stats.ELV < 20 Then
                            
                            If .Clase = eClass.Worker And .Stats.UserSkills(eSkill.pesca) <> 100 Then
                                Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 20 y ademas tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                            Exit Sub
                        Else

                            ' Esta entre 20 y 25, si es trabajador necesita tener 100 en pesca
                            If .Clase = eClass.Worker Then
                                If .Stats.UserSkills(eSkill.pesca) <> 100 Then
                                    Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 20 o superior y ademas tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub

                                End If

                            End If

                        End If

                    End If

                End If
                
                If (LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False)) And .flags.Navegando = 0 Then
                    Call DoNavega(Userindex, obj, Slot)
                
                ElseIf (LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, False, True) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, False, True) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, False, True) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, False, True)) And .flags.Navegando = 1 Then
                    Call DoNavega(Userindex, obj, Slot)
                    
                Else
                    Call WriteConsoleMsg(Userindex, "Debes aproximarte al agua para usar un barco y a la tierra para desembarcar!", FontTypeNames.FONTTYPE_INFO)

                End If


            '<-------------> MONTURAS <----------->
            Case eOBJType.otMonturas
                If ClasePuedeUsarItem(Userindex, ObjIndex) Then
                    If .flags.invisible = 1 Then
                        Call WriteConsoleMsg(Userindex, "Estas invisible, no puedes montarte ni desmontarte en este estado!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "Estas muerto, no puedes montarte ni desmontarte en este estado!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If .flags.Navegando = 1 Then
                        Call WriteConsoleMsg(Userindex, "Estas navegando, no puedes montarte ni desmontarte en este estado!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call DoEquita(Userindex, obj, Slot)
                Else
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                End If
                    
            Case eOBJType.otManuales
            
                Select Case ObjIndex
                
                    Case 1127   ' Manual de Liderazgo
                        
                        If .Stats.UserSkills(eSkill.Liderazgo) < 100 Then
                            .Stats.UserSkills(eSkill.Liderazgo) = 100
                            Call QuitarUserInvItem(Userindex, Slot, 1)
                            Call UpdateUserInv(False, Userindex, Slot)
                        End If
                        
                    Case 1128   ' Manual de Supervivencia
                        
                        If .Stats.UserSkills(eSkill.Supervivencia) < 100 Then
                            .Stats.UserSkills(eSkill.Supervivencia) = 100
                            Call QuitarUserInvItem(Userindex, Slot, 1)
                            Call UpdateUserInv(False, Userindex, Slot)
                        End If
                        
                    Case 1129   ' Manual de Navegacion
                        
                        If .Stats.UserSkills(eSkill.Navegacion) < 100 Then
                            .Stats.UserSkills(eSkill.Navegacion) = 100
                            Call QuitarUserInvItem(Userindex, Slot, 1)
                            Call UpdateUserInv(False, Userindex, Slot)
                        End If
                        
                End Select
                    
            End Select
    
    End With

End Sub

Sub EnivarArmasConstruibles(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteBlacksmithWeapons(Userindex)

End Sub
 
Sub EnivarObjConstruibles(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteInitCarpenting(Userindex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteBlacksmithArmors(Userindex)

End Sub

Sub TirarTodo(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(Userindex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        Call TirarTodosLosItems(Userindex)
        
        Dim Cantidad As Long: Cantidad = .Stats.Gld - CLng(.Stats.ELV) * 10000
       
        ' Si estas en zona segura tampoco se tira el oro.
        If MapInfo(.Pos.Map).Pk Then
            
            If Cantidad > 0 Then
                Call TirarOro(Cantidad, Userindex)
            End If
            
        End If
        
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en TirarTodo. Error: " & Err.Number & " - " & Err.description)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With ObjData(index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And (.Caos <> 1 Or .NoSeCae = 0) And .OBJType <> eOBJType.otLlaves And .OBJType <> eOBJType.otBarcos And .NoSeCae = 0

    End With

End Function

Sub TirarTodosLosItems(ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    '12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
    '***************************************************
    On Error GoTo ErrHandler

    Dim i         As Byte

    Dim NuevaPos  As WorldPos

    Dim MiObj     As obj

    Dim ItemIndex As Integer

    Dim DropAgua  As Boolean
    
    With UserList(Userindex)

        For i = 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex

                    DropAgua = True

                    ' Es pirata?
                    If .Clase = eClass.Pirat Then

                        ' Si tiene galeon equipado
                        If .Invent.BarcoObjIndex = 476 Then

                            ' Limitacion por nivel, despues dropea normalmente
                            If .Stats.ELV = 20 Then
                                ' No dropea en agua
                                DropAgua = False

                            End If

                        End If

                    End If
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                    End If

                End If

            End If

        Next i

    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en TirarTodosLosItems. Error: " & Err.Number & " - " & Err.description)

End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 23/11/2009
    '07/11/09: Pato - Fix bug #2819911
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    Dim i         As Byte

    Dim NuevaPos  As WorldPos

    Dim MiObj     As obj

    Dim ItemIndex As Integer
    
    With UserList(Userindex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        For i = 1 To UserList(Userindex).CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
                    Tilelibre .Pos, NuevaPos, MiObj, True, True

                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                    End If

                End If

            End If

        Next i

    End With

End Sub

Sub TirarTodosLosItemsEnMochila(ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/09 (Budi)
    '***************************************************
    Dim i         As Byte

    Dim NuevaPos  As WorldPos

    Dim MiObj     As obj

    Dim ItemIndex As Integer
    
    With UserList(Userindex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        For i = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If Not ItemSeCae(ItemIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Acabas de tirar un objeto que no se cae normalmente ya que lo tenias en tu mochila u alforja y la desequipaste o tiraste", FontTypeNames.FONTTYPE_WARNING)
                End If

                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo MiObj
                MiObj.Amount = .Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                Tilelibre .Pos, NuevaPos, MiObj, True, True

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                End If

            End If

        Next i

    End With

End Sub

Public Function getObjType(ByVal ObjIndex As Integer) As eOBJType
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If ObjIndex > 0 Then
        getObjType = ObjData(ObjIndex).OBJType

    End If
    
End Function

Public Sub moveItem(ByVal Userindex As Integer, _
                    ByVal originalSlot As Integer, _
                    ByVal newSlot As Integer)

    Dim tmpObj      As UserObj

    Dim newObjIndex As Integer, originalObjIndex As Integer

    If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub

    With UserList(Userindex)

        If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
    
        tmpObj = .Invent.Object(originalSlot)
        .Invent.Object(originalSlot) = .Invent.Object(newSlot)
        .Invent.Object(newSlot) = tmpObj
    
        'Viva VB6 y sus putas deficiencias.
        If .Invent.AnilloEqpSlot = originalSlot Then
            .Invent.AnilloEqpSlot = newSlot
        ElseIf .Invent.AnilloEqpSlot = newSlot Then
            .Invent.AnilloEqpSlot = originalSlot

        End If
    
        If .Invent.ArmourEqpSlot = originalSlot Then
            .Invent.ArmourEqpSlot = newSlot
        ElseIf .Invent.ArmourEqpSlot = newSlot Then
            .Invent.ArmourEqpSlot = originalSlot

        End If
    
        If .Invent.BarcoSlot = originalSlot Then
            .Invent.BarcoSlot = newSlot
        ElseIf .Invent.BarcoSlot = newSlot Then
            .Invent.BarcoSlot = originalSlot

        End If
    
        If .Invent.CascoEqpSlot = originalSlot Then
            .Invent.CascoEqpSlot = newSlot
        ElseIf .Invent.CascoEqpSlot = newSlot Then
            .Invent.CascoEqpSlot = originalSlot

        End If
    
        If .Invent.EscudoEqpSlot = originalSlot Then
            .Invent.EscudoEqpSlot = newSlot
        ElseIf .Invent.EscudoEqpSlot = newSlot Then
            .Invent.EscudoEqpSlot = originalSlot

        End If
    
        If .Invent.MochilaEqpSlot = originalSlot Then
            .Invent.MochilaEqpSlot = newSlot
        ElseIf .Invent.MochilaEqpSlot = newSlot Then
            .Invent.MochilaEqpSlot = originalSlot

        End If
    
        If .Invent.MunicionEqpSlot = originalSlot Then
            .Invent.MunicionEqpSlot = newSlot
        ElseIf .Invent.MunicionEqpSlot = newSlot Then
            .Invent.MunicionEqpSlot = originalSlot

        End If
    
        If .Invent.WeaponEqpSlot = originalSlot Then
            .Invent.WeaponEqpSlot = newSlot
        ElseIf .Invent.WeaponEqpSlot = newSlot Then
            .Invent.WeaponEqpSlot = originalSlot

        End If

        Call UpdateUserInv(False, Userindex, originalSlot)
        Call UpdateUserInv(False, Userindex, newSlot)

    End With

End Sub
