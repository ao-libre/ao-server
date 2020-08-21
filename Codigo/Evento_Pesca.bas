Attribute VB_Name = "Evento_Pesca"

Option Explicit

Private Type tPescaEvent

    Activado As Byte
    Tiempo As Byte
    CantidadDeZonas As Byte

End Type

Private Type tZona

    Mapa As Integer
    Cantidad As Byte
    Peces() As Integer

End Type

Public PescaEvent As tPescaEvent

Public Zona()     As tZona

Public Sub LoadPeces()

    Dim Leer As clsIniManager

    Set Leer = New clsIniManager
 
    Call Leer.Initialize(App.Path & "\Dat\EventoPesca.dat")
 
    Dim i As Integer

    Dim j As Integer
 
    With PescaEvent
        ReDim Zona(1 To .CantidadDeZonas) As tZona
 
        For i = 1 To .CantidadDeZonas

            With Zona(i)
                .Mapa = Leer.GetValue("ZONA" & i, "Mapa")
                .Cantidad = Leer.GetValue("ZONA" & i, "Cantidad")
             
                ReDim Zona(i).Peces(1 To .Cantidad) As Integer
             
                For j = 1 To .Cantidad
                    .Peces(j) = Leer.GetValue("ZONA" & i, "Pez" & j)
                Next j

            End With

        Next i

    End With
 
    Set Leer = Nothing
End Sub

Public Function DamePez(ByVal ZonaUser As Byte) As Long
    DamePez = Zona(ZonaUser).Peces(RandomNumber(LBound(Zona(ZonaUser).Peces()), UBound(Zona(ZonaUser).Peces())))
End Function

Public Sub CheckEstadoDelMar(ByRef MinsEventoPesca As Long)

    With PescaEvent

        If MinsEventoPesca > .Tiempo Then
            If .Tiempo > 0 Then
                If .Activado = 0 Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El oceano se agita y comienza la temporada de pesca.", FontTypeNames.FONTTYPE_INFOBOLD))
                    
                    .Activado = 1

                    With PescaEvent
                        Dim i As Long

                        For i = 1 To .CantidadDeZonas
                            Dim EventoPescaMapasMensaje As String
                            With Zona(i)

                                If i = PescaEvent.CantidadDeZonas Then
                                    EventoPescaMapasMensaje = EventoPescaMapasMensaje & .Mapa
                                Else
                                    EventoPescaMapasMensaje = EventoPescaMapasMensaje & .Mapa & ", " 
                                End If

                            End With

                        Next i

                    End With

                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Mapas donde hay marea alta y es recomendable pescar ahora: " & EventoPescaMapasMensaje, FontTypeNames.FONTTYPE_INFOBOLD))
                    MinsEventoPesca = 0
                Else
                    .Activado = 0
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Las sombras aterrorizan a las especies y estas deciden volver al profundo oceano, la temporada de pesca termina.", FontTypeNames.FONTTYPE_INFOBOLD))
                    MinsEventoPesca = 0
                End If
            End If
        End If

    End With
End Sub
