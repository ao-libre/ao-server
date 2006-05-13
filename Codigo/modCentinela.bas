Attribute VB_Name = "modCentinela"
'*****************************************************************
'modCentinela.bas - ImperiumAO - v1.2
'
'Funciónes de control para usuarios que se encuentran trabajando
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

'*****************************************************************
'Augusto Rando(barrin@imperiumao.com.ar)
'   - First Relase
'
'Juan Martín Sotuyo Dodero (juansotuyo@gmail.com)
'   - Adapted to Alkon AO
'   - Small improvements and added logs to detect possible cheaters
'*****************************************************************

Option Explicit

Private Const NPC_CENTINELA_TIERRA As Integer = 16  'Índice del NPC en el .dat
Private Const NPC_CENTINELA_AGUA As Integer = 16     'Ídem anterior, pero en mapas de agua

Public CentinelaNPCIndex As Integer                'Índice del NPC en el servidor

Private Const TIEMPO_INICIAL As Byte = 2 'Tiempo inicial en minutos. No reducir sin antes revisar el timer que maneja estos datos.

Private Type tCentinela
    RevisandoUserIndex As Integer   '¿Qué índice revisamos?
    TiempoRestante As Integer       '¿Cuántos minutos le quedan al usuario?
    clave As Integer                'Clave que debe escribir
End Type

Public centinelaActivado As Boolean

Public Centinela As tCentinela

Private Sub GoToNextWorkingChar()
'############################################################
'Va al siguiente usuario que se encuentre trabajando
'############################################################
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).name <> "" And UserList(LoopC).Counters.Trabajando > 0 And UserList(LoopC).flags.Privilegios = PlayerType.User Then
            If Not UserList(LoopC).flags.CentinelaOK Then
                'Inicializamos
                Centinela.RevisandoUserIndex = LoopC
                Centinela.TiempoRestante = TIEMPO_INICIAL
                Centinela.clave = RandomNumber(1, 36000)
                
                'Ponemos al centinela en posición
                Call WarpCentinela(LoopC)
                
                If CentinelaNPCIndex Then
                    'Mandamos el mensaje (el centinela habla y aparece en consola para que no haya dudas
                    Call SendData(SendTarget.ToIndex, LoopC, 0, "||" & vbGreen & "°" & "Saludos " & UserList(LoopC).name & ", soy el Centinela de estas tierras. Me gustaría que escribas /CENTINELA " & Centinela.clave & " en no más de dos minutos." & "°" & CStr(Npclist(CentinelaNPCIndex).Char.CharIndex))
                    Call SendData(SendTarget.ToIndex, LoopC, 0, "||" & "Saludos " & UserList(LoopC).name & ", soy el Centinela de estas tierras. Me gustaría que escribas /CENTINELA " & Centinela.clave & " en no más de dos minutos." & FONTTYPE_CENTINELA)
                End If
                Exit Sub
            End If
        End If
    Next LoopC
    
    'No hay chars trabajando, eliminamos el NPC si todavía estaba en algún lado y esperamos otro minuto
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0
    End If
    
    'No estamos revisando a nadie
    Centinela.RevisandoUserIndex = 0
End Sub

Private Sub CentinelaFinalCheck()
'############################################################
'Al finalizar el tiempo, se retira y realiza la acción
'pertinente dependiendo del caso
'############################################################
On Error GoTo Error_Handler
    Dim name As String
    Dim numPenas As Integer
    
    If Not UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK Then
        'Logueamos el evento
        Call LogCentinela("Centinela baneo a " & UserList(Centinela.RevisandoUserIndex).name & " por uso de macro inasistido")
        
        'Ponemos el ban
        UserList(Centinela.RevisandoUserIndex).flags.Ban = 1
        
        name = UserList(Centinela.RevisandoUserIndex).name
        
        'Avisamos a los admins
        Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> El centinela ha baneado a " & name & FONTTYPE_SERVER)
        
        'ponemos el flag de ban a 1
        Call WriteVar(CharPath & name & ".chr", "FLAGS", "Ban", "1")
        'ponemos la pena
        numPenas = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & name & ".chr", "PENAS", "Cant", numPenas + 1)
        Call WriteVar(CharPath & name & ".chr", "PENAS", "P" & numPenas + 1, "CENTINELA : BAN POR MACRO INASISTIDO " & Date & " " & Time)
        
        'Evitamos loguear el logout
        Dim index As Integer
        index = Centinela.RevisandoUserIndex
        Centinela.RevisandoUserIndex = 0
        
        Call CloseSocket(index)
    End If
    
    Centinela.clave = 0
    Centinela.TiempoRestante = 0
    Centinela.RevisandoUserIndex = 0
    
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0
    End If
Exit Sub

Error_Handler:
    Centinela.clave = 0
    Centinela.TiempoRestante = 0
    Centinela.RevisandoUserIndex = 0
    
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0
    End If
    
    Call LogError("Error en el checkeo del centinela: " & Err.Description)
End Sub

Public Sub CentinelaCheckClave(ByVal UserIndex As Integer, ByVal clave As Integer)
'############################################################
'Corrobora la clave que le envia el usuario
'############################################################
    If clave = Centinela.clave And UserIndex = Centinela.RevisandoUserIndex Then
        UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK = True
        Centinela.RevisandoUserIndex = 0
        Call SendData(SendTarget.ToIndex, Centinela.RevisandoUserIndex, 0, "||" & vbWhite & "°" & "¡Muchas gracias " & UserList(Centinela.RevisandoUserIndex).name & "! Espero no haber sido una molestia" & "°" & CStr(Npclist(CentinelaNPCIndex).Char.CharIndex))
    Else
        Call CentinelaSendClave(UserIndex)
        
        If UserIndex <> Centinela.RevisandoUserIndex Then
            'Logueamos el evento
            Call LogCentinela("El usuario " & UserList(UserIndex).name & " respondió aunque no se le hablaba a él.")
        End If
    End If
End Sub

Public Sub ResetCentinelaInfo()
'############################################################
'Cada determinada cantidad de tiempo, volvemos a revisar
'############################################################
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If (UserList(LoopC).name <> "" And LoopC <> Centinela.RevisandoUserIndex) Then
            UserList(LoopC).flags.CentinelaOK = False
        End If
    Next LoopC
End Sub

Public Sub CentinelaSendClave(ByVal UserIndex As Integer)
'############################################################
'Enviamos al usuario la clave vía el personaje centinela
'############################################################
    If CentinelaNPCIndex = 0 Then Exit Sub
    
    If UserIndex = Centinela.RevisandoUserIndex Then
        If Not UserList(UserIndex).flags.CentinelaOK Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "¡La clave que te he dicho es " & "/CENTINELA " & Centinela.clave & " escríbelo rápido!" & "°" & CStr(Npclist(CentinelaNPCIndex).Char.CharIndex))
        Else
            'Logueamos el evento
            Call LogCentinela("El usuario " & UserList(Centinela.RevisandoUserIndex).name & " respondió más de una vez la contraseña correcta.")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Te agradezco, pero ya me has respondido. Me retiraré pronto." & "°" & CStr(Npclist(CentinelaNPCIndex).Char.CharIndex))
        End If
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No es a ti a quien estoy hablando, ¿no ves?" & "°" & CStr(Npclist(CentinelaNPCIndex).Char.CharIndex))
    End If
End Sub

Public Sub PasarMinutoCentinela()
'############################################################
'Control del timer. Llamado cada un minuto.
'############################################################
    If Not centinelaActivado Then Exit Sub
    
    If Centinela.RevisandoUserIndex = 0 Then
        Call GoToNextWorkingChar
    Else
        Centinela.TiempoRestante = Centinela.TiempoRestante - 1
        
        If Centinela.TiempoRestante = 0 Then
            Call CentinelaFinalCheck
            Call GoToNextWorkingChar
        Else
            'Recordamos al user que debe escribir
            If Matematicas.Distancia(Npclist(CentinelaNPCIndex).Pos, UserList(Centinela.RevisandoUserIndex).Pos) > 5 Then
                Call WarpCentinela(Centinela.RevisandoUserIndex)
            End If
            
            'El centinela habla y se manda a consola para que no quepan dudas
            Call SendData(SendTarget.ToIndex, Centinela.RevisandoUserIndex, 0, "||" & vbRed & "°¡" & UserList(Centinela.RevisandoUserIndex).name & ", tienes un minuto más para responder! Debes escribir /CENTINELA " & Centinela.clave & "." & "°" & CStr(Npclist(CentinelaNPCIndex).Char.CharIndex))
            Call SendData(SendTarget.ToIndex, Centinela.RevisandoUserIndex, 0, "||" & "¡" & UserList(Centinela.RevisandoUserIndex).name & ", tienes un minuto más para responder! Debes escribir /CENTINELA " & Centinela.clave & "." & FONTTYPE_CENTINELA)
        End If
    End If
End Sub

Private Sub WarpCentinela(ByVal UserIndex As Integer)
'############################################################
'Inciamos la revisión del usuario UserIndex
'############################################################
    'Evitamos conflictos de índices
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0
    End If
    
    If HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
        CentinelaNPCIndex = SpawnNpc(NPC_CENTINELA_AGUA, UserList(UserIndex).Pos, True, False)
    Else
        CentinelaNPCIndex = SpawnNpc(NPC_CENTINELA_TIERRA, UserList(UserIndex).Pos, True, False)
    End If
    
    'Si no pudimos crear el NPC, seguimos esperando a poder hacerlo
    If CentinelaNPCIndex = 0 Then _
        Centinela.RevisandoUserIndex = 0
End Sub

Public Sub CentinelaUserLogout()
'############################################################
'El usuario al que revisabamos se desconectó
'############################################################
    If Centinela.RevisandoUserIndex Then
        'Revisamos si no respondió ya
        If UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK Then Exit Sub
        
        'Logueamos el evento
        Call LogCentinela("El usuario " & UserList(Centinela.RevisandoUserIndex).name & " se desolgueó al pedirsele la contraseña")
        
        'Reseteamos y esperamos a otro PasarMinuto para ir al siguiente user
        Centinela.clave = 0
        Centinela.TiempoRestante = 0
        Centinela.RevisandoUserIndex = 0
        
        If CentinelaNPCIndex Then
            Call QuitarNPC(CentinelaNPCIndex)
            CentinelaNPCIndex = 0
        End If
    End If
End Sub

Private Sub LogCentinela(ByVal texto As String)
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 03/15/2006
'Loguea un evento del centinela
'*************************************************
On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\Centinela.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
Exit Sub

errhandler:
End Sub
