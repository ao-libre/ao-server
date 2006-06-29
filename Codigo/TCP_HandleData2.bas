Attribute VB_Name = "TCP_HandleData2"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Public Sub HandleData_2(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim T() As String
Dim i As Integer

Procesado = True 'ver al final del sub


    Select Case UCase$(rData)
        Case "/ONLINE"
            'No se envia más la lista completa de usuarios
            N = 0
            For LoopC = 1 To LastUser
                If UserList(LoopC).name <> "" And UserList(LoopC).flags.Privilegios <= PlayerType.Consejero Then
                    N = N + 1
                End If
            Next LoopC
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Número de usuarios: " & N & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIR"
            If UserList(userindex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub
            End If
            ''mato los comercios seguros
            If UserList(userindex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                        Call SendData(SendTarget.ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(userindex)
            End If
            Call Cerrar_Usuario(userindex)
            Exit Sub
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(userindex, UserList(userindex).name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            
            Exit Sub

            
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 3 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                      Exit Sub
            End If
            Select Case Npclist(UserList(userindex).flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                If FileExist(CharPath & UCase$(UserList(userindex).name) & ".chr", vbNormal) = False Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (userindex)
                      Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Case eNPCType.Timbero
                If UserList(userindex).flags.Privilegios > PlayerType.User Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    N = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & N & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(userindex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                          Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                          Exit Sub
             End If
             If Npclist(UserList(userindex).flags.TargetNPC).MaestroUser <> _
                userindex Then Exit Sub
             Npclist(UserList(userindex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(userindex).flags.TargetNPC, userindex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).MaestroUser <> _
              userindex Then Exit Sub
            Call FollowAmo(UserList(userindex).flags.TargetNPC)
            Call Expresar(UserList(userindex).flags.TargetNPC, userindex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
        Case "/DESCANSAR"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If HayOBJarea(UserList(userindex).Pos, FOGATA) Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "DOK")
                    If Not UserList(userindex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(userindex).flags.Descansar = Not UserList(userindex).flags.Descansar
            Else
                    If UserList(userindex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(userindex).flags.Descansar = False
                        Call SendData(SendTarget.ToIndex, userindex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/MEDITAR"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).flags.Privilegios > PlayerType.User Then
                UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call SendUserStatsBox(val(userindex))
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, userindex, 0, "MEDOK")
            If Not UserList(userindex).flags.Meditando Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Comenzas a meditar." & FONTTYPE_INFO)
            Else
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
            End If
           UserList(userindex).flags.Meditando = Not UserList(userindex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(userindex).flags.Meditando Then
                UserList(userindex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Te estás concentrando. En " & TIEMPO_INICIOMEDITAR & " segundos comenzarás a meditar." & FONTTYPE_INFO)
                
                UserList(userindex).Char.loops = LoopAdEternum
                If UserList(userindex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARCHICO
                ElseIf UserList(userindex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(userindex).Stats.ELV < 45 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARGRANDE
                Else
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARXGRANDE & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARXGRANDE
                End If
            Else
                UserList(userindex).Counters.bPuedeMeditar = False
                
                UserList(userindex).Char.FX = 0
                UserList(userindex).Char.loops = 0
                Call SendData(SendTarget.ToMap, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           Call RevivirUsuario(userindex)
           Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/CURAR"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
           Call SendUserStatsBox(userindex)
           Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Hás sido curado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/AYUDA"
           Call SendHelp(userindex)
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(userindex, userindex)
            Exit Sub
        
        Case "/SEG"
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "SEGOFF")
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "SEGON")
            End If
            UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            Exit Sub
    
    
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).flags.Comerciando Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(userindex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(userindex)
            '[Alejo]
            ElseIf UserList(userindex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(userindex).flags.TargetUser = userindex Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(userindex).flags.TargetUser).Pos, UserList(userindex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(userindex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(userindex).flags.TargetUser).ComUsu.DestUsu <> userindex Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(userindex).ComUsu.DestUsu = UserList(userindex).flags.TargetUser
                UserList(userindex).ComUsu.DestNick = UserList(UserList(userindex).flags.TargetUser).name
                UserList(userindex).ComUsu.cant = 0
                UserList(userindex).ComUsu.Objeto = 0
                UserList(userindex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(userindex, UserList(userindex).flags.TargetUser)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/Alejo]
        '[KEVIN]------------------------------------------
        Case "/BOVEDA"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(userindex)
                End If
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(userindex)
           Else
                  Call EnlistarCaos(userindex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(userindex)
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(userindex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(userindex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Próximo mantenimiento automático: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(userindex)
            Exit Sub
        
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(userindex) Then Exit Sub
            Call mdParty.CrearParty(userindex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(userindex)
            Exit Sub
        Case "/ENCUESTA"
            ConsultaPopular.SendInfoEncuesta (userindex)
    End Select

    If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)
        If UserList(userindex).GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "|+" & UserList(userindex).name & "> " & rData & FONTTYPE_GUILDMSG)
'TODO : Con la 0.11.7 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToClanArea, userindex, UserList(userindex).Pos.Map, "||" & vbYellow & "°< " & rData & " >°" & CStr(UserList(userindex).Char.CharIndex))
        End If
        
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then
            Call mdParty.BroadCastParty(userindex, mid$(rData, 7))
'TODO : Con la 0.11.7 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, userindex, UserList(userindex).Pos.Map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(userindex).Char.CharIndex))
        End If
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 11)) = "/CENTINELA " Then
        'Evitamos overflow y underflow
        If val(Right$(rData, Len(rData) - 11)) > &H7FFF Or val(Right$(rData, Len(rData) - 11)) < &H8000 Then Exit Sub
        
        tInt = val(Right$(rData, Len(rData) - 11))
        Call CentinelaCheckClave(userindex, tInt)
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINECLAN" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(userindex, UserList(userindex).GuildIndex)
        If UserList(userindex).GuildIndex <> 0 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Compañeros de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(userindex)
        Exit Sub
    End If
    
    '[yb]
    If UCase$(Left$(rData, 6)) = "/BMSG " Then
        rData = Right$(rData, Len(rData) - 6)
        If UserList(userindex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, userindex, 0, "|| (Consejero) " & UserList(userindex).name & "> " & rData & FONTTYPE_CONSEJO)
        End If
        If UserList(userindex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, userindex, 0, "|| (Consejero) " & UserList(userindex).name & "> " & rData & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.ToIndex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(userindex).name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 6)) = "/GMSG " And UserList(userindex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 6)
        Call LogGM(UserList(userindex).name, "Mensaje a Gms:" & rData, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & "> " & rData & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 3))
        Case "/GM"
            If Not Ayuda.Existe(UserList(userindex).name) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                Call Ayuda.Push(rData, UserList(userindex).name)
            Else
                Call Ayuda.Quitar(UserList(userindex).name)
                Call Ayuda.Push(rData, UserList(userindex).name)
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
            End If
            Exit Sub
    End Select
    
    
    
    Select Case UCase(Left(rData, 5))
        Case "/_BUG "
            N = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
            Print #N,
            Print #N,
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N, "Usuario:" & UserList(userindex).name & "  Fecha:" & Date & "    Hora:" & Time
            Print #N, "########################################################################"
            Print #N, "BUG:"
            Print #N, Right$(rData, Len(rData) - 5)
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N,
            Print #N,
            Close #N
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 6))
        Case "/DESC "
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes cambiar la descripción estando muerto." & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(userindex).Desc = Trim$(rData)
            Call SendData(SendTarget.ToIndex, userindex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(userindex, rData, tStr) Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        name = Right$(rData, Len(rData) - 7)
        If name = "" Then Exit Sub
        
        name = Replace(name, "\", "")
        name = Replace(name, "/", "")
        
        If FileExist(CharPath & name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tInt & "- " & GetVar(CharPath & name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Personaje """ & name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    
    
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)
            If Len(rData) < 6 Then
                 Call SendData(SendTarget.ToIndex, userindex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                 Call SendData(SendTarget.ToIndex, userindex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                 UserList(userindex).Password = rData
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))
            If tLong > 32000 Then tLong = 32000
            N = tLong
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            ElseIf UserList(userindex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
            ElseIf Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
            ElseIf Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            ElseIf N < 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            ElseIf N > 5000 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            ElseIf UserList(userindex).Stats.GLD < N Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + N
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - N
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call SendUserStatsBox(userindex)
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 10))
            'consultas populares muchacho'
        Case "/ENCUESTA "
            rData = Right(rData, Len(rData) - 10)
            If Len(rData) = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| Aca va la info de la encuesta" & FONTTYPE_GUILD)
                Exit Sub
            End If
            DummyInt = CLng(val(rData))
            Call SendData(SendTarget.ToIndex, userindex, 0, "|| " & ConsultaPopular.doVotar(userindex, DummyInt) & FONTTYPE_GUILD)
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
             End If
             
             If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(userindex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(userindex)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                        Debug.Print "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    End If
                ElseIf UserList(userindex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(userindex)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                End If
                Exit Sub
             
             End If
             
             If Len(rData) = 8 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rData = Right$(rData, Len(rData) - 9)
             If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(userindex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(userindex).name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.ToIndex, userindex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (userindex)
                  Exit Sub
             End If
             If val(rData) > 0 And val(rData) <= UserList(userindex).Stats.Banco Then
                  UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco - val(rData)
                  UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + val(rData)
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             End If
             Call SendUserStatsBox(val(userindex))
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).Pos, UserList(userindex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(userindex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(userindex).Pos, Npclist(UserList(userindex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(userindex).Stats.GLD Then
                  UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco + val(rData)
                  UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(rData)
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            End If
            Call SendUserStatsBox(val(userindex))
            Exit Sub
        Case "/DENUNCIAR "
            If UserList(userindex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "|| " & LCase$(UserList(userindex).name) & " DENUNCIA: " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToIndex, userindex, 0, "|| Denuncia enviada, espere.." & FONTTYPE_INFO)
            Exit Sub
        Case "/FUNDARCLAN"
        
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rData))
                    Case "ARMADA"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_ARMADA
                    Case "MAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_LEGION
                    Case "NEUTRO"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "GM"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "LEGAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "CRIMINAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "|| Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "SHOWFUN")
            Else
                UserList(userindex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 12))
        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(userindex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(userindex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 13))
        Case "/ACCEPTPARTY "
            rData = Right$(rData, Len(rData) - 13)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(userindex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rData, 14))
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            name = Replace(rData, "\", "")
            name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
    
    Procesado = False
End Sub
