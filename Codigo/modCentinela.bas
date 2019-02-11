' programado por maTih.-

' TODO: CONFIGURAR ESTOS PARAMETROS 

Option Explicit
 
Public CentinelaEstado  As Boolean          'Está activado?
 
Const NUM_CENTINELAS    As Byte = 5         'Cantidad de centinelas.
Const NUM_NPC           As Integer = 16     'NpcNum del centinela.
 
Const MAPA_EXPLOTAR     As Integer = 15     'Numero de mapa en la qe se pinchan usuarios.
Const X_EXPLOTAR        As Byte = 50        'X
Const Y_EXPLOTAR        As Byte = 50        'Y
 
Const LIMITE_TIEMPO     As Long = 120000    'Tiempo límite (milisegundos), 2 minutos.
Const CARCEL_TIEMPO     As Byte = 5         'Minutos en la carcel
Const REVISION_TIEMPO   As Long = 1800000   'Tiempo de cada revisión (milisegundos) 1.800.000 = 30 minutos (60 segundos * 30 minutos) * 1000 milisegundos
 
Type Centinelas
     MiNpcIndex         As Integer          'NPCIndex del centinela.
     Invocado           As Boolean          'Si está invocado.
     RevisandoSlot      As Integer          'UI Del usuario.
     TiempoInicio       As Long             'Desde que empezó el chekeo al usuario.
     CodigoCheck        As String           'Codigo que debe ingresar el usuario.
End Type
 
Public Centinelas(1 To NUM_CENTINELAS)  As Centinelas
 
Sub CambiarEstado(ByVal gmIndex As Integer)
 
' @ Cambia el estado del centinela.
 
Dim TmpStr  As String
 
CentinelaEstado = Not CentinelaEstado
 
TmpStr = UserList(gmIndex).name & " Cambió el estado del Centinela:" & IIf(CentinelaEstado, " Ahora está activado.", " Ahora está desactivado.")
 
Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(TmpStr, FontTypeNames.FONTTYPE_CONSE))
 
Call LogGM(UserList(gmIndex).name, "Cambió el estado del centinela (Enabled:" & CentinelaEstado & ")")
 
End Sub
 
Sub EnviarAUsuario(ByVal userIndex As Integer, ByVal CIndex As Byte)
 
' @ Envia centinela a un usuario.
 
With Centinelas(CIndex)
        
     'Genera el código.
     .CodigoCheck = GenerarClave
     
     'Spawnea.
     .MiNpcIndex = SpawnNpc(NUM_NPC, DarPosicion(userIndex), True, False)
     
     'Setea el flag.
     .Invocado = (.MiNpcIndex <> 0)
     
     'No spawnea, error !
     If Not .Invocado Then .CodigoCheck = vbNullString: Exit Sub
     
     'Avisa al usuario sobre el char del centinela.
     Call AvisarUsuario(userIndex, CIndex)
     
     'Setea el tiempo.
     .TiempoInicio = GetTickCount()
     
     'Setea UI del usuario
     .RevisandoSlot = userIndex
     
End With
 
'Setea los datos del user.
With UserList(userIndex).CentinelaUsuario
     .CentinelaCheck = False                    'Por defecto, no ingreso la clave.
     .centinelaIndex = CIndex                   'Setea el index del mismo.
     .Codigo = Centinelas(CIndex).CodigoCheck   'Setea el código.
     .Revisando = True                          'Lo revisa un centinela.
End With
 
End Sub
 
Sub AvisarUsuarios()
 
' @ Envia la clave a los usuarios de los centinelas
 
Dim i   As Long
 
For i = 1 To NUM_CENTINELAS
    With Centinelas(i)
         'Si está invocado.
         If .Invocado Then
            'Avisa.
            Call AvisarUsuario(.RevisandoSlot, CByte(i))
         End If
    End With
Next i
 
End Sub
 
Sub AvisarUsuario(ByVal userSlot As Integer, ByVal centinelaIndex As Byte, Optional ByVal IngresoFallido As Boolean = False)
 
' @ Avisa al usuario la clave..
 
With Centinelas(centinelaIndex)
 
     Dim DataSend   As String
     
     'Para avisar.
     If Not IngresoFallido Then
        'Pasó la mitad de tiempo?
        If (GetTickCount() - .TiempoInicio) > (LIMITE_TIEMPO / 2) Then
            'Prepara el paquete a enviar.
            DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, Debes escribir /CENTINELA " & .CodigoCheck & " En menos de 2 minutos.", Npclist(.MiNpcIndex).Char.CharIndex, vbRed)
        Else
            DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, Tienes menos de un minuto para escribir /CENTINELA " & .CodigoCheck & ".", Npclist(.MiNpcIndex).Char.CharIndex, vbRed)
         End If
     Else
         DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, El código ingresado NO es correcto, debes escribir : /CENTINELA " & .CodigoCheck & ".", Npclist(.MiNpcIndex).Char.CharIndex, vbRed)
     End If
     
     'Envia.
     Call EnviarDatosASlot(userSlot, DataSend)
    
End With
 
End Sub
 
Sub ChekearUsuarios()
 
Dim LoopC   As Long
Dim CIndex  As Byte
 
For LoopC = 1 To LastUser
 
    With UserList(LoopC)
    
         'Lo revisa el centinela?
         If .CentinelaUsuario.Revisando Then
            Call TiempoUsuario(CInt(LoopC))
         Else
            'Está trabajando?
            If .Counters.Trabajando <> 0 Then
               'Si todavia no lo revisaron o si pasó más del tiempo sin revisar, vuelve a enviar.
               If Not .CentinelaUsuario.CentinelaCheck Or ((GetTickCount() - .CentinelaUsuario.UltimaRevision) > REVISION_TIEMPO) Then
                  'Busca un slot para centinela y se lo envia.
                  CIndex = ProximoCentinela
                  'Si hay slot
                  If CIndex <> 0 Then
                     'Envia
                     Call EnviarAUsuario(CInt(LoopC), CIndex)
                  End If
               End If
            End If
         End If
        
    End With
 
Next LoopC
 
End Sub
 
Sub IngresaClave(ByVal userIndex As Integer, ByRef Clave As String)
 
' @ Checkea la clave que ingresó el usuario.
 
Clave = UCase$(Clave)
 
Dim centinelaIndex  As Byte
 
centinelaIndex = UserList(userIndex).CentinelaUsuario.centinelaIndex
 
'No tiene centinela.
If Not centinelaIndex <> 0 Then Exit Sub
 
'No está revisandolo.
If Not UserList(userIndex).CentinelaUsuario.Revisando Then Exit Sub
 
'Checkea el código
If CheckCodigo(Clave, centinelaIndex) Then
   'Quita el centinela.
   Call AprobarUsuario(userIndex, centinelaIndex)
Else
   'Avisa.
   Call AvisarUsuario(userIndex, centinelaIndex, True)
End If
 
 
End Sub
 
Sub AprobarUsuario(ByVal userIndex As Integer, ByVal CIndex As Byte)
 
' @ Aprueba el control de un usuario.
 
'Dim ClearType  as centinelaUser
 
With UserList(userIndex)
     
     '.CentinelaUsuario = ClearType
     
     'Quita el char.
     Call LimpiarIndice(.CentinelaUsuario.centinelaIndex)
     
     With .CentinelaUsuario
          .CentinelaCheck = True
          .centinelaIndex = 0
          .Codigo = vbNullString
          .Revisando = False
          .UltimaRevision = 0
     End With
 
     Call Protocol.WriteConsoleMsg(userIndex, "El control ha finalizado.", FontTypeNames.FONTTYPE_DIOS)
     
End With
 
End Sub
 
Sub LimpiarIndice(ByVal centinelaIndex As Byte)
 
' @ Limpia un slot.
 
With Centinelas(centinelaIndex)
 
     .Invocado = False
     .CodigoCheck = vbNullString
     .RevisandoSlot = 0
     .TiempoInicio = 0
     
     'Estaba el char?
     If .MiNpcIndex <> 0 Then
        Call QuitarNPC(.MiNpcIndex)
     End If
 
End With
 
End Sub
 
Sub TiempoUsuario(ByVal userIndex As Integer)
 
' @ Checkea el tiempo para contestar de un usuario.
 
Dim centinelaIndex  As Byte
 
With UserList(userIndex).CentinelaUsuario
 
     centinelaIndex = .centinelaIndex
     
     'No hay indice ! WTF XD
     If Not centinelaIndex <> 0 Then Exit Sub
     
     'Acabó el tiempo y no ingresó la clave.
     If (GetTickCount - Centinelas(centinelaIndex).TiempoInicio) > LIMITE_TIEMPO Then
        Call UsuarioInActivo(userIndex)
     End If
 
End With
 
End Sub
 
Sub UsuarioInActivo(ByVal userIndex As Integer)
 
' @ No contestó el usuario, se lo pena.
 
'Telep al mapa.
Call WarpUserChar(userIndex, MAPA_EXPLOTAR, X_EXPLOTAR, Y_EXPLOTAR, True)
 
'Muere.
Call UserDie(userIndex)
 
'Tira los items.
Call TirarTodosLosItems(userIndex)
 
'Lo encarcela.
Call Encarcelar(userIndex, CARCEL_TIEMPO, "El centinela")
 
'Borra el centinela.
If UserList(userIndex).CentinelaUsuario.centinelaIndex <> 0 Then
    Call LimpiarIndice(UserList(userIndex).CentinelaUsuario.centinelaIndex)
End If
 
'Deja un mensaje.
Call Protocol.WriteConsoleMsg(userIndex, "El centinela te ha ejecutado y encarcelado por Macro Inasistido.", FontTypeNames.FONTTYPE_DIOS)
 
'Limpia el tipo del usuario.
Dim ClearType   As CentinelaUser
 
UserList(userIndex).CentinelaUsuario = ClearType
 
UserList(userIndex).CentinelaUsuario.CentinelaCheck = True
 
End Sub
 
Function GenerarClave() As String
 
' @ Arma la clave para un centinela.
 
Dim NumCharacters   As Byte     'Numero de carácteres de la clave.
Dim LoopC           As Long
 
NumCharacters = 7
 
For LoopC = 1 To NumCharacters
    'Una letra y un numero.
    If (LoopC Mod 2) <> 0 Then  '< Es número INPar.
       'Letra.
       GenerarClave = GenerarClave & Chr$(RandomNumber(97, 122))
    Else
       'número.
       GenerarClave = GenerarClave & RandomNumber(1, 9)
    End If
Next LoopC
 
'Pasa a mayúsculas.
GenerarClave = UCase$(GenerarClave)
 
End Function
 
Function DarPosicion(ByVal userIndex As Integer) As WorldPos
 
' @ Devuelve la posición para spawnear al centinela.
 
With UserList(userIndex)
     'Posición del usuario original.
     DarPosicion = .Pos
     
     'Mueve una posición a la derecha
     DarPosicion.X = .Pos.X + 1
     
End With
 
End Function
 
Function ProximoCentinela() As Byte
 
' @ Devuelve un slot para un centinela.
 
Dim i   As Long
 
For i = 1 To NUM_CENTINELAS
    'Si no está invocado.
    If Not Centinelas(i).Invocado Then
       'Devuelve el slot
       ProximoCentinela = CByte(i)
       Exit Function
    End If
Next i
 
ProximoCentinela = 0
 
End Function
 
Function CheckCodigo(ByRef Ingresada As String, ByVal CIndex As Byte) As Boolean
 
' @ Devuelve si el código es correcto.
 
CheckCodigo = (Not Ingresada <> Centinelas(CIndex).CodigoCheck)
 
End Function