Attribute VB_Name = "ModFacciones"
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

Public ArmaduraImperial1              As Integer

Public ArmaduraImperial2              As Integer

Public ArmaduraImperial3              As Integer

Public TunicaMagoImperial             As Integer

Public TunicaMagoImperialEnanos       As Integer

Public ArmaduraCaos1                  As Integer

Public ArmaduraCaos2                  As Integer

Public ArmaduraCaos3                  As Integer

Public TunicaMagoCaos                 As Integer

Public TunicaMagoCaosEnanos           As Integer

Public VestimentaImperialHumano       As Integer

Public VestimentaImperialEnano        As Integer

Public TunicaConspicuaHumano          As Integer

Public TunicaConspicuaEnano           As Integer

Public ArmaduraNobilisimaHumano       As Integer

Public ArmaduraNobilisimaEnano        As Integer

Public ArmaduraGranSacerdote          As Integer

Public VestimentaLegionHumano         As Integer

Public VestimentaLegionEnano          As Integer

Public TunicaLobregaHumano            As Integer

Public TunicaLobregaEnano             As Integer

Public TunicaEgregiaHumano            As Integer

Public TunicaEgregiaEnano             As Integer

Public SacerdoteDemoniaco             As Integer

Public Const NUM_RANGOS_FACCION       As Integer = 15

Private Const NUM_DEF_FACCION_ARMOURS As Byte = 3

Public Enum eTipoDefArmors

    ieBaja
    ieMedia
    ieAlta

End Enum

Public Type tFaccionArmaduras

    Armada(NUM_DEF_FACCION_ARMOURS - 1) As Integer
    Caos(NUM_DEF_FACCION_ARMOURS - 1) As Integer

End Type

' Matriz que contiene las armaduras faccionarias segun raza, clase, faccion y defensa de armadura
Public ArmadurasFaccion(1 To NUMCLASES, 1 To NUMRAZAS) As tFaccionArmaduras

' Contiene la cantidad de exp otorgada cada vez que aumenta el rango
Public RecompensaFacciones(NUM_RANGOS_FACCION)         As Long

Private Function GetArmourAmount(ByVal Rango As Integer, _
                                 ByVal TipoDef As eTipoDefArmors) As Integer
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 15/04/2010
    'Returns the amount of armours to give, depending on the specified rank
    '***************************************************

    Select Case TipoDef
        
        Case eTipoDefArmors.ieBaja
            GetArmourAmount = 20 / (Rango + 1)
            
        Case eTipoDefArmors.ieMedia
            GetArmourAmount = Rango * 2 / MaximoInt((Rango - 4), 1)
            
        Case eTipoDefArmors.ieAlta
            GetArmourAmount = Rango * 1.35
            
    End Select
    
End Function

Private Sub GiveFactionArmours(ByVal Userindex As Integer, ByVal IsCaos As Boolean)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 15/04/2010
    'Gives faction armours to user
    '***************************************************
    
    Dim ObjArmour As obj

    Dim Rango     As Integer
    
    With UserList(Userindex)
    
        Rango = val(IIf(IsCaos, .Faccion.RecompensasCaos, .Faccion.RecompensasReal)) + 1
    
        ' Entrego armaduras de defensa baja
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieBaja)
        
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieBaja)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieBaja)

        End If
        
        If Not MeterItemEnInventario(Userindex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)

        End If
        
        ' Entrego armaduras de defensa media
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieMedia)
        
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieMedia)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieMedia)

        End If
        
        If Not MeterItemEnInventario(Userindex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)

        End If
    
        ' Entrego armaduras de defensa alta
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieAlta)
        
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta)

        End If
        
        If Not MeterItemEnInventario(Userindex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)

        End If

    End With

End Sub

Public Sub GiveExpReward(ByVal Userindex As Integer, ByVal Rango As Long)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 15/04/2010
    'Gives reward exp to user
    '***************************************************
    
    Dim GivenExp As Long
    
    With UserList(Userindex)
        
        GivenExp = RecompensaFacciones(Rango)
        
        .Stats.Exp = .Stats.Exp + GivenExp
        
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        Call WriteConsoleMsg(Userindex, "Has sido recompensado con " & GivenExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)

        Call CheckUserLevel(Userindex)
        
    End With
    
End Sub

Public Sub EnlistarArmadaReal(ByVal Userindex As Integer)
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 15/04/2010
    'Handles the entrance of users to the "Armada Real"
    '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineacion neutral.
    '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
    '15/04/2010: ZaMa - Cambio en recompensas iniciales.
    '***************************************************

    With UserList(Userindex)

        If .Faccion.ArmadaReal = 1 Then
            Call WriteChatOverHead(Userindex, "Ya perteneces a las tropas reales!!! Ve a combatir criminales.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.FuerzasCaos = 1 Then
            Call WriteChatOverHead(Userindex, "Maldito insolente!!! Vete de aqui seguidor de las sombras.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If criminal(Userindex) Then
            Call WriteChatOverHead(Userindex, "No se permiten criminales en el ejercito real!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.CriminalesMatados < 30 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, solo has matado " & .Faccion.CriminalesMatados & ".", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Stats.ELV < 25 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
     
        If .Faccion.CiudadanosMatados > 0 Then
            Call WriteChatOverHead(Userindex, "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.Reenlistadas > 4 Then
            Call WriteChatOverHead(Userindex, "Has sido expulsado de las fuerzas reales demasiadas veces!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Reputacion.NobleRep < 1000000 Then
            Call WriteChatOverHead(Userindex, "Necesitas ser aun mas noble para integrar el ejercito real, solo tienes " & .Reputacion.NobleRep & "/1.000.000 puntos de nobleza", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .GuildIndex > 0 Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                Call WriteChatOverHead(Userindex, "Perteneces a un clan neutro, sal de el si quieres unirte a nuestras fuerzas!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

            End If

        End If
    
        .Faccion.ArmadaReal = 1
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
    
        Call WriteChatOverHead(Userindex, "Bienvenido al ejercito real!!! Aqui tienes tus vestimentas. Cumple bien tu labor exterminando criminales y me encargare de recompensarte.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
        ' TODO: Dejo esta variable por ahora, pero con chequear las reenlistadas deberia ser suficiente :S
        If .Faccion.RecibioArmaduraReal = 0 Then
        
            Call GiveFactionArmours(Userindex, False)
            Call GiveExpReward(Userindex, 0)
        
            .Faccion.RecibioArmaduraReal = 1
            .Faccion.NivelIngreso = .Stats.ELV
            .Faccion.FechaIngreso = DatePart("yyyy", Date) & "-" & DatePart("m", Date) & "-" & DatePart("d", Date)

            'Esto por ahora es inutil, siempre va a ser cero, pero bueno, despues va a servir.
            .Faccion.MatadosIngreso = .Faccion.CiudadanosMatados
        
            .Faccion.RecibioExpInicialReal = 1
            .Faccion.RecompensasReal = 0
            .Faccion.NextRecompensa = 70
        
        End If
    
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)
    
        Call LogEjercitoReal(.Name & " ingreso el " & Date & " cuando era nivel " & .Stats.ELV)

    End With

End Sub

Public Sub RecompensaArmadaReal(ByVal Userindex As Integer)

    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 15/04/2010
    'Handles the way of gaining new ranks in the "Armada Real"
    '15/04/2010: ZaMa - Agrego recompensas de oro y armaduras
    '***************************************************
    Dim Crimis    As Long

    Dim Lvl       As Byte

    Dim NextRecom As Long

    Dim Nobleza   As Long

    With UserList(Userindex)
        Lvl = .Stats.ELV
        Crimis = .Faccion.CriminalesMatados
        NextRecom = .Faccion.NextRecompensa
        Nobleza = .Reputacion.NobleRep
    
        If Crimis < NextRecom Then
            Call WriteChatOverHead(Userindex, "Mata " & NextRecom - Crimis & " criminales mas para recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        Select Case NextRecom

            Case 70:
                .Faccion.RecompensasReal = 1
                .Faccion.NextRecompensa = 130
        
            Case 130:
                .Faccion.RecompensasReal = 2
                .Faccion.NextRecompensa = 210
        
            Case 210:
                .Faccion.RecompensasReal = 3
                .Faccion.NextRecompensa = 320
        
            Case 320:
                .Faccion.RecompensasReal = 4
                .Faccion.NextRecompensa = 460
        
            Case 460:
                .Faccion.RecompensasReal = 5
                .Faccion.NextRecompensa = 640
        
            Case 640:

                If Lvl < 27 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 27 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasReal = 6
                .Faccion.NextRecompensa = 870
        
            Case 870:
                .Faccion.RecompensasReal = 7
                .Faccion.NextRecompensa = 1160
        
            Case 1160:
                .Faccion.RecompensasReal = 8
                .Faccion.NextRecompensa = 2000
        
            Case 2000:

                If Lvl < 30 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 30 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasReal = 9
                .Faccion.NextRecompensa = 2500
        
            Case 2500:

                If Nobleza < 2000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 2000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasReal = 10
                .Faccion.NextRecompensa = 3000
        
            Case 3000:

                If Nobleza < 3000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 3000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasReal = 11
                .Faccion.NextRecompensa = 3500
        
            Case 3500:

                If Lvl < 35 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 35 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                If Nobleza < 4000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 4000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasReal = 12
                .Faccion.NextRecompensa = 4000
        
            Case 4000:

                If Lvl < 36 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 36 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                If Nobleza < 5000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 5000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasReal = 13
                .Faccion.NextRecompensa = 5000
        
            Case 5000:

                If Lvl < 37 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 37 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                If Nobleza < 6000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 6000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasReal = 14
                .Faccion.NextRecompensa = 10000
        
            Case 10000:
                Call WriteChatOverHead(Userindex, "Eres uno de mis mejores soldados. Mataste " & Crimis & " criminales, sigue asi. Ya no tengo mas recompensa para darte que mi agradecimiento. Felicidades!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
        
            Case Else:
                Exit Sub

        End Select
    
        Call WriteChatOverHead(Userindex, "Aqui tienes tu recompensa " & TituloReal(Userindex) & "!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

        ' Recompensas de armaduras y exp
        Call GiveFactionArmours(Userindex, False)
        Call GiveExpReward(Userindex, .Faccion.RecompensasReal)

    End With

End Sub

Public Sub ExpulsarFaccionReal(ByVal Userindex As Integer, _
                               Optional Expulsado As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' 09/28/2010 C4b3z0n - Arreglado RT6 Overflow, el Desequipar() del escudo, ponia de parametro el ObjIndex del escudo en vez del EqpSlot.
    '***************************************************

    With UserList(Userindex)
        .Faccion.ArmadaReal = 0

        'Call PerderItemsFaccionarios(UserIndex)
        If Expulsado Then
            Call WriteConsoleMsg(Userindex, "Has sido expulsado del ejercito real!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has retirado del ejercito real!!!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        If .Invent.ArmourEqpObjIndex <> 0 Then

            'Desequipamos la armadura real si esta equipada
            If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(Userindex, .Invent.ArmourEqpSlot)

        End If
    
        If .Invent.EscudoEqpObjIndex <> 0 Then

            'Desequipamos el escudo de caos si esta equipado
            If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(Userindex, .Invent.EscudoEqpSlot)

        End If
    
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)

    End With

End Sub

Public Sub ExpulsarFaccionCaos(ByVal Userindex As Integer, _
                               Optional Expulsado As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' 09/28/2010 C4b3z0n - Arreglado RT6 Overflow, el Desequipar() del escudo, ponia de parametro el ObjIndex del escudo en vez del EqpSlot.
    '***************************************************

    With UserList(Userindex)
        .Faccion.FuerzasCaos = 0

        'Call PerderItemsFaccionarios(UserIndex)
        If Expulsado Then
            Call WriteConsoleMsg(Userindex, "Has sido expulsado de la Legion Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has retirado de la Legion Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        If .Invent.ArmourEqpObjIndex <> 0 Then

            'Desequipamos la armadura de caos si esta equipada
            If ObjData(.Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(Userindex, .Invent.ArmourEqpSlot)

        End If
    
        If .Invent.EscudoEqpObjIndex <> 0 Then

            'Desequipamos el escudo de caos si esta equipado
            If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(Userindex, .Invent.EscudoEqpSlot)

        End If
    
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)

    End With

End Sub

Public Function TituloReal(ByVal Userindex As Integer) As String
    '***************************************************
    'Autor: Unknown
    'Last Modification: 23/01/2007 Pablo (ToxicWaste)
    'Handles the titles of the members of the "Armada Real"
    '***************************************************

    Select Case UserList(Userindex).Faccion.RecompensasReal
            'Rango 1: Aprendiz (30 Criminales)
            'Rango 2: Escudero (70 Criminales)
            'Rango 3: Soldado (130 Criminales)
            'Rango 4: Sargento (210 Criminales)
            'Rango 5: Caballero (320 Criminales)
            'Rango 6: Comandante (460 Criminales)
            'Rango 7: Capitan (640 Criminales + > lvl 27)
            'Rango 8: Senescal (870 Criminales)
            'Rango 9: Mariscal (1160 Criminales)
            'Rango 10: Condestable (2000 Criminales + > lvl 30)
            'Rangos de Honor de la Armada Real: (Consejo de Bander)
            'Rango 11: Ejecutor Imperial (2500 Criminales + 2.000.000 Nobleza)
            'Rango 12: Protector del Reino (3000 Criminales + 3.000.000 Nobleza)
            'Rango 13: Avatar de la Justicia (3500 Criminales + 4.000.000 Nobleza + > lvl 35)
            'Rango 14: Guardian del Bien (4000 Criminales + 5.000.000 Nobleza + > lvl 36)
            'Rango 15: Campeon de la Luz (5000 Criminales + 6.000.000 Nobleza + > lvl 37)
    
        Case 0
            TituloReal = "Aprendiz"

        Case 1
            TituloReal = "Escudero"

        Case 2
            TituloReal = "Soldado"

        Case 3
            TituloReal = "Sargento"

        Case 4
            TituloReal = "Teniente"

        Case 5
            TituloReal = "Comandante"

        Case 6
            TituloReal = "Capitan"

        Case 7
            TituloReal = "Senescal"

        Case 8
            TituloReal = "Mariscal"

        Case 9
            TituloReal = "Condestable"

        Case 10
            TituloReal = "Ejecutor Imperial"

        Case 11
            TituloReal = "Protector del Reino"

        Case 12
            TituloReal = "Avatar de la Justicia"

        Case 13
            TituloReal = "Guardian del Bien"

        Case Else
            TituloReal = "Campeon de la Luz"

    End Select

End Function

Public Sub EnlistarCaos(ByVal Userindex As Integer)
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 27/11/2009
    '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineacion neutral.
    '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
    'Handles the entrance of users to the "Legion Oscura"
    '***************************************************

    With UserList(Userindex)

        If Not criminal(Userindex) Then
            Call WriteChatOverHead(Userindex, "Largate de aqui, bufon!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.FuerzasCaos = 1 Then
            Call WriteChatOverHead(Userindex, "Ya perteneces a la legion oscura!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.ArmadaReal = 1 Then
            Call WriteChatOverHead(Userindex, "Las sombras reinaran en Argentum. Fuera de aqui insecto real!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        '[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
        If .Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ahi: Recibio la experiencia para entrar?
            Call WriteChatOverHead(Userindex, "No permitire que ningun insecto real ingrese a mis tropas.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        '[/Barrin]
    
        If Not criminal(Userindex) Then
            Call WriteChatOverHead(Userindex, "Ja ja ja!! Tu no eres bienvenido aqui asqueroso ciudadano.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.CiudadanosMatados < 70 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes matar al menos 70 ciudadanos, solo has matado " & .Faccion.CiudadanosMatados & ".", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Stats.ELV < 25 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes ser al menos nivel 25!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .GuildIndex > 0 Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                Call WriteChatOverHead(Userindex, "Perteneces a un clan neutro, sal de el si quieres unirte a nuestras fuerzas!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

            End If

        End If
    
        If .Faccion.Reenlistadas > 4 Then
            If .Faccion.Reenlistadas = 200 Then
                Call WriteChatOverHead(Userindex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldia has atacado a mi ejercito. Vete de aqui!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Else
                Call WriteChatOverHead(Userindex, "Has sido expulsado de las fuerzas oscuras demasiadas veces!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

            End If

            Exit Sub

        End If
    
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.FuerzasCaos = 1
    
        Call WriteChatOverHead(Userindex, "Bienvenido al lado oscuro!!! Aqui tienes tus armaduras. Derrama sangre ciudadana y real, y seras recompensado, lo prometo.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
        If .Faccion.RecibioArmaduraCaos = 0 Then
                
            Call GiveFactionArmours(Userindex, True)
            Call GiveExpReward(Userindex, 0)
        
            .Faccion.RecibioArmaduraCaos = 1
            .Faccion.NivelIngreso = .Stats.ELV
            .Faccion.FechaIngreso = DatePart("yyyy", Date) & "-" & DatePart("m", Date) & "-" & DatePart("d", Date)
    
            .Faccion.RecibioExpInicialCaos = 1
            .Faccion.RecompensasCaos = 0
            .Faccion.NextRecompensa = 160

        End If
    
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)

        Call LogEjercitoCaos(.Name & " ingreso el " & Date & " cuando era nivel " & .Stats.ELV)

    End With

End Sub

Public Sub RecompensaCaos(ByVal Userindex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 15/04/2010
    'Handles the way of gaining new ranks in the "Legion Oscura"
    '15/04/2010: ZaMa - Agrego recompensas de oro y armaduras
    '***************************************************
    Dim Ciudas    As Long

    Dim Lvl       As Byte

    Dim NextRecom As Long

    With UserList(Userindex)
        Lvl = .Stats.ELV
        Ciudas = .Faccion.CiudadanosMatados
        NextRecom = .Faccion.NextRecompensa
    
        If Ciudas < NextRecom Then
            Call WriteChatOverHead(Userindex, "Mata " & NextRecom - Ciudas & " cuidadanos mas para recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        Select Case NextRecom

            Case 160:
                .Faccion.RecompensasCaos = 1
                .Faccion.NextRecompensa = 300
        
            Case 300:
                .Faccion.RecompensasCaos = 2
                .Faccion.NextRecompensa = 490
        
            Case 490:
                .Faccion.RecompensasCaos = 3
                .Faccion.NextRecompensa = 740
        
            Case 740:
                .Faccion.RecompensasCaos = 4
                .Faccion.NextRecompensa = 1100
        
            Case 1100:
                .Faccion.RecompensasCaos = 5
                .Faccion.NextRecompensa = 1500
        
            Case 1500:

                If Lvl < 27 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 27 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasCaos = 6
                .Faccion.NextRecompensa = 2010
        
            Case 2010:
                .Faccion.RecompensasCaos = 7
                .Faccion.NextRecompensa = 2700
        
            Case 2700:
                .Faccion.RecompensasCaos = 8
                .Faccion.NextRecompensa = 4600
        
            Case 4600:

                If Lvl < 30 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 30 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasCaos = 9
                .Faccion.NextRecompensa = 5800
        
            Case 5800:

                If Lvl < 31 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 31 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasCaos = 10
                .Faccion.NextRecompensa = 6990
        
            Case 6990:

                If Lvl < 33 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 33 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasCaos = 11
                .Faccion.NextRecompensa = 8100
        
            Case 8100:

                If Lvl < 35 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 35 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasCaos = 12
                .Faccion.NextRecompensa = 9300
        
            Case 9300:

                If Lvl < 36 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 36 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasCaos = 13
                .Faccion.NextRecompensa = 11500
        
            Case 11500:

                If Lvl < 37 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 37 - Lvl & " niveles para poder recibir la proxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

                .Faccion.RecompensasCaos = 14
                .Faccion.NextRecompensa = 23000
        
            Case 23000:
                Call WriteChatOverHead(Userindex, "Eres uno de mis mejores soldados. Mataste " & Ciudas & " ciudadanos . Tu unica recompensa sera la sangre derramada. Continua asi!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
        
            Case Else:
                Exit Sub
            
        End Select
    
        Call WriteChatOverHead(Userindex, "Bien hecho " & TituloCaos(Userindex) & ", aqui tienes tu recompensa!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
        ' Recompensas de armaduras y exp
        Call GiveFactionArmours(Userindex, True)
        Call GiveExpReward(Userindex, .Faccion.RecompensasCaos)
    
    End With

End Sub

Public Function TituloCaos(ByVal Userindex As Integer) As String
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 23/01/2007 Pablo (ToxicWaste)
    'Handles the titles of the members of the "Legion Oscura"
    '***************************************************
    'Rango 1: Acolito (70)
    'Rango 2: Alma Corrupta (160)
    'Rango 3: Paria (300)
    'Rango 4: Condenado (490)
    'Rango 5: Esbirro (740)
    'Rango 6: Sanguinario (1100)
    'Rango 7: Corruptor (1500 + lvl 27)
    'Rango 8: Heraldo Impio (2010)
    'Rango 9: Caballero de la Oscuridad (2700)
    'Rango 10: Senor del Miedo (4600 + lvl 30)
    'Rango 11: Ejecutor Infernal (5800 + lvl 31)
    'Rango 12: Protector del Averno (6990 + lvl 33)
    'Rango 13: Avatar de la Destruccion (8100 + lvl 35)
    'Rango 14: Guardian del Mal (9300 + lvl 36)
    'Rango 15: Campeon de la Oscuridad (11500 + lvl 37)

    Select Case UserList(Userindex).Faccion.RecompensasCaos

        Case 0
            TituloCaos = "Acolito"

        Case 1
            TituloCaos = "Alma Corrupta"

        Case 2
            TituloCaos = "Paria"

        Case 3
            TituloCaos = "Condenado"

        Case 4
            TituloCaos = "Esbirro"

        Case 5
            TituloCaos = "Sanguinario"

        Case 6
            TituloCaos = "Corruptor"

        Case 7
            TituloCaos = "Heraldo Impio"

        Case 8
            TituloCaos = "Caballero de la Oscuridad"

        Case 9
            TituloCaos = "Senor del Miedo"

        Case 10
            TituloCaos = "Ejecutor Infernal"

        Case 11
            TituloCaos = "Protector del Averno"

        Case 12
            TituloCaos = "Avatar de la Destruccion"

        Case 13
            TituloCaos = "Guardian del Mal"

        Case Else
            TituloCaos = "Campeon de la Oscuridad"

    End Select

End Function

