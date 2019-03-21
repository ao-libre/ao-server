Attribute VB_Name = "Mod_Barcos"
' Basado en cliente DunkanAO

Option Explicit

'CONFIGURACION DE LOS BARCOS
Private Const Num_Barcos As Byte = 1


Private Type tBarcos
    Activo As Byte
    Char As Char
    Personajes As Byte ' Cantidad de personajes en el barco
    
    NumMapas As Byte
    Mapa() As WorldPos
    MapaTraslado() As WorldPos
    MapasRestantes As Byte
End Type

Public Barcos(1 To Num_Barcos) As tBarcos

Public Sub CargarBarcos()
    Dim i As Long
    Dim ii As Long
    Dim Dir As String
    
    Dir = App.Path & "\DAT\Barcos.dat"
    
    For i = 1 To Num_Barcos
        With Barcos(i)
            .Activo = 0
            .Char.body = val(GetVar(Dir, "VIAJE" & i, "Char"))
            .NumMapas = val(GetVar(Dir, "VIAJE" & i, "NumMapas"))
            
            'Mapas de viaje
            ReDim .Mapa(1 To .NumMapas) As WorldPos
            For ii = 1 To .NumMapas
                .Mapa(ii).Map = CInt(ReadField(1, GetVar(Dir, "VIAJE" & i, "Mapa" & ii), 45))
                .Mapa(ii).X = CInt(ReadField(2, GetVar(Dir, "VIAJE" & i, "Mapa" & ii), 45))
                .Mapa(ii).Y = CInt(ReadField(3, GetVar(Dir, "VIAJE" & i, "Mapa" & ii), 45))
            Next ii
            
            'Mapas traslados
            ReDim .MapaTraslado(1 To .NumMapas) As WorldPos
            For ii = 1 To .NumMapas
                .MapaTraslado(ii).Map = CInt(ReadField(1, GetVar(Dir, "VIAJE" & i, "Mapa" & ii & "_Traslado"), 45))
                .MapaTraslado(ii).X = CInt(ReadField(2, GetVar(Dir, "VIAJE" & i, "Mapa" & ii & "_Traslado"), 45))
                .MapaTraslado(ii).Y = CInt(ReadField(3, GetVar(Dir, "VIAJE" & i, "Mapa" & ii & "_Traslado"), 45))
            Next ii
        End With
    Next i
End Sub

Public Sub ViajeEnCurso(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
    
        Dim Diferencia As Integer
        
        ' La diferencia proviene de la cantidad de mapas , restando la cantidad faltante.
        ' EJ: Son 4 mapas, y los restantes son 3, significa que el uno ya lo transitamos.
        Diferencia = Barcos(1).NumMapas - Barcos(1).MapasRestantes
        
        ' Chequeo de llegada
        If Diferencia = 0 Then
            Call WriteConsoleMsg(UserIndex, "Llegada", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        ' Movimiento del barco
        Call GreedyWalkTo(NpcIndex, Barcos(1).Mapa(Diferencia).Map, Barcos(1).Mapa(Diferencia).X, Barcos(1).Mapa(Diferencia).Y)
        
    
    End With
End Sub
