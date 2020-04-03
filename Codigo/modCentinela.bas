Attribute VB_Name = "modCentinela"
' programado por maTih.-

' TODO: CONFIGURAR ESTOS PARAMETROS

Option Explicit
 
Public isCentinelaActivated As Boolean          'Esta activado?
 
Const NUM_CENTINELAS   As Byte = 5         'Cantidad de centinelas.

Const NUM_NPC          As Integer = 16     'NpcNum del centinela.
 
Const MAPA_EXPLOTAR    As Integer = 15     'Numero de mapa en la qe se pinchan usuarios.

Const X_EXPLOTAR       As Byte = 50        'X

Const Y_EXPLOTAR       As Byte = 50        'Y
 
Const LIMITE_TIEMPO    As Long = 120000    'Tiempo limite (milisegundos), 2 minutos.

Const CARCEL_TIEMPO    As Byte = 5         'Minutos en la carcel

Const REVISION_TIEMPO  As Long = 1800000   'Tiempo de cada revision (milisegundos) 1.800.000 = 30 minutos (60 segundos * 30 minutos) * 1000 milisegundos
 
Type Centinelas

    MiNpcIndex         As Integer          'NPCIndex del centinela.
    Invocado           As Boolean          'Si esta invocado.
    RevisandoSlot      As Integer          'UI Del usuario.
    TiempoInicio       As Long             'Desde que empezo el chekeo al usuario.
    CodigoCheck        As String           'Codigo que debe ingresar el usuario.

End Type
 
Public Centinelas(1 To NUM_CENTINELAS) As Centinelas
 
Sub CambiarEstado(ByVal gmIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/11/2019 (Recox)
'13/11/2019 Recox: La variable isCentinelaActivated tiene un nombre mas descriptivo
'***************************************************

    ' @ Cambia el estado del centinela.

    'Lo cambiamos en la memoria.
    isCentinelaActivated = Not isCentinelaActivated
    
    'Lo cambiamos en el Server.ini
    Call WriteVar(IniPath & "Server.ini", "INIT", "CentinelaAuditoriaTrabajoActivo", IIf(isCentinelaActivated, 1, 0))

    'Preparamos el aviso por consola.
    Dim message As String
    message = UserList(gmIndex).Name & " cambio el estado del Centinela a " & IIf(isCentinelaActivated, " ACTIVADO.", " DESACTIVADO.")
    
    'Mandamos el aviso por consola.
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_CENTINELA))
    
    'Lo registramos en los logs.
    Call LogGM(UserList(gmIndex).Name, message)
 
End Sub
 
Sub EnviarAUsuario(ByVal Userindex As Integer, ByVal CIndex As Byte)
 
    ' @ Envia centinela a un usuario.
 
    With Centinelas(CIndex)
        
        'Genera el codigo.
        .CodigoCheck = GenerarClave
     
        'Spawnea.
        .MiNpcIndex = SpawnNpc(NUM_NPC, DarPosicion(Userindex), True, False)
     
        'Setea el flag.
        .Invocado = (.MiNpcIndex <> 0)
     
        'No spawnea, error !
        If Not .Invocado Then
            .CodigoCheck = vbNullString
            Exit Sub
        End If

     
        'Avisa al usuario sobre el char del centinela.
        Call AvisarUsuario(Userindex, CIndex)
     
        'Setea el tiempo.
        .TiempoInicio = GetTickCount()
     
        'Setea UI del usuario
        .RevisandoSlot = Userindex
     
    End With
 
    'Setea los datos del user.
    With UserList(Userindex).CentinelaUsuario
        .CentinelaCheck = False                    'Por defecto, no ingreso la clave.
        .centinelaIndex = CIndex                   'Setea el index del mismo.
        .Codigo = Centinelas(CIndex).CodigoCheck   'Setea el codigo.
        .Revisando = True                          'Lo revisa un centinela.

    End With
 
End Sub
 
Sub AvisarUsuarios()
 
    ' @ Envia la clave a los usuarios de los centinelas
 
    Dim i As Long
 
    For i = 1 To NUM_CENTINELAS

        With Centinelas(i)

            'Si esta invocado.
            If .Invocado Then
                'Avisa.
                Call AvisarUsuario(.RevisandoSlot, CByte(i))

            End If

        End With

    Next i
 
End Sub
 
Sub AvisarUsuario(ByVal userSlot As Integer, _
                  ByVal centinelaIndex As Byte, _
                  Optional ByVal IngresoFallido As Boolean = False)
 
    ' @ Avisa al usuario la clave..
 
    With Centinelas(centinelaIndex)
 
        Dim DataSend As String
     
        'Para avisar.
        If Not IngresoFallido Then

            'Paso la mitad de tiempo?
            If (GetTickCount() - .TiempoInicio) > (LIMITE_TIEMPO / 2) Then
                'Prepara el paquete a enviar.
                DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, Debes escribir /CENTINELA " & .CodigoCheck & " En menos de 2 minutos.", Npclist(.MiNpcIndex).Char.CharIndex, vbYellow)
            Else
                DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, Tienes menos de un minuto para escribir /CENTINELA " & .CodigoCheck & ".", Npclist(.MiNpcIndex).Char.CharIndex, vbYellow)
            End If

        Else
            DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, El codigo ingresado NO es correcto, debes escribir : /CENTINELA " & .CodigoCheck & ".", Npclist(.MiNpcIndex).Char.CharIndex, vbYellow)

        End If
     
        'Envia.
        Call UserList(userSlot).outgoingData.WriteASCIIStringFixed(DataSend)
    
    End With
 
End Sub
 
Sub ChekearUsuarios()
 
    Dim LoopC  As Long

    Dim CIndex As Byte
 
    For LoopC = 1 To LastUser
 
        With UserList(LoopC)
    
            'Lo revisa el centinela?
            If .CentinelaUsuario.Revisando Then
                Call TiempoUsuario(CInt(LoopC))
            Else

                'Esta trabajando?
                If .Counters.Trabajando <> 0 Then

                    'Si todavia no lo revisaron o si paso mas del tiempo sin revisar, vuelve a enviar.
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
 
Sub IngresaClave(ByVal Userindex As Integer, ByRef Clave As String)
 
    ' @ Checkea la clave que ingreso el usuario.
 
    Clave = UCase$(Clave)
 
    Dim centinelaIndex As Byte
 
    centinelaIndex = UserList(Userindex).CentinelaUsuario.centinelaIndex
 
    'No tiene centinela.
    If Not centinelaIndex <> 0 Then Exit Sub
 
    'No esta revisandolo.
    If Not UserList(Userindex).CentinelaUsuario.Revisando Then Exit Sub
 
    'Checkea el codigo
    If CheckCodigo(Clave, centinelaIndex) Then
        'Quita el centinela.
        Call AprobarUsuario(Userindex, centinelaIndex)
    Else
        'Avisa.
        Call AvisarUsuario(Userindex, centinelaIndex, True)

    End If
 
End Sub
 
Sub AprobarUsuario(ByVal Userindex As Integer, ByVal CIndex As Byte)
 
    ' @ Aprueba el control de un usuario.
 
    'Dim ClearType  as centinelaUser
 
    With UserList(Userindex)
     
        '.CentinelaUsuario = ClearType
     
        'Quita el char.
        Call LimpiarIndice(.CentinelaUsuario.centinelaIndex)
     
        With .CentinelaUsuario
            .CentinelaCheck = True
            .centinelaIndex = 0
            .Codigo = vbNullString
            .Revisando = False
            .UltimaRevision = GetTickCount()
        End With
 
        Call Protocol.WriteConsoleMsg(Userindex, "El control ha finalizado.", FontTypeNames.FONTTYPE_DIOS)
     
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
 
Sub TiempoUsuario(ByVal Userindex As Integer)
 
    ' @ Checkea el tiempo para contestar de un usuario.
 
    Dim centinelaIndex As Byte
 
    With UserList(Userindex).CentinelaUsuario
 
        centinelaIndex = .centinelaIndex
     
        'No hay indice ! WTF XD
        If Not centinelaIndex <> 0 Then Exit Sub
     
        'Acabo el tiempo y no ingreso la clave.
        If (GetTickCount - Centinelas(centinelaIndex).TiempoInicio) > LIMITE_TIEMPO Then
            Call UsuarioInActivo(Userindex)
        End If
 
    End With
 
End Sub
 
Sub UsuarioInActivo(ByVal Userindex As Integer)
 
    ' @ No contesto el usuario, se lo pena.
 
    'Telep al mapa.
    Call WarpUserChar(Userindex, MAPA_EXPLOTAR, X_EXPLOTAR, Y_EXPLOTAR, True)
 

    'No creo que tirar los items sea justo, con encarcelarlo y matarlo es mas que suficiente. (Recox)
    'Aparte de que si muere desaparecen los items...
    'Muere.
    'Call UserDie(Userindex)
 
    'Tira los items.
    'Call TirarTodosLosItems(Userindex)
 
    'Lo encarcela.
    Call Encarcelar(Userindex, CARCEL_TIEMPO, "El centinela")
 
    'Borra el centinela.
    If UserList(Userindex).CentinelaUsuario.centinelaIndex <> 0 Then
        Call LimpiarIndice(UserList(Userindex).CentinelaUsuario.centinelaIndex)
    End If
 
    'Deja un mensaje.
    Call Protocol.WriteConsoleMsg(Userindex, "El centinela te ha ejecutado y encarcelado por Macro Inasistido.", FontTypeNames.FONTTYPE_DIOS)
 
    'Limpia el tipo del usuario.
    Dim ClearType As CentinelaUser
 
    UserList(Userindex).CentinelaUsuario = ClearType
 
    UserList(Userindex).CentinelaUsuario.CentinelaCheck = True
 
End Sub
 
Function GenerarClave() As String
 
    ' @ Arma la clave para un centinela.
 
    Dim NumCharacters As Byte     'Numero de caracteres de la clave.

    Dim LoopC         As Long
 
    NumCharacters = 4
 
    For LoopC = 1 To NumCharacters

        'Una letra y un numero.
        If (LoopC Mod 2) <> 0 Then  '< Es numero INPar.
            'Letra.
            GenerarClave = GenerarClave & Chr$(RandomNumber(97, 122))
        Else
            'numero.
            GenerarClave = GenerarClave & RandomNumber(1, 9)

        End If

    Next LoopC
 
    'Pasa a mayusculas.
    GenerarClave = UCase$(GenerarClave)
 
End Function
 
Function DarPosicion(ByVal Userindex As Integer) As WorldPos
 
    ' @ Devuelve la posicion para spawnear al centinela.
 
    With UserList(Userindex)
        'Posicion del usuario original.
        DarPosicion = .Pos
     
        'Mueve una posicion a la derecha
        DarPosicion.x = .Pos.x + 1
     
    End With
 
End Function
 
Function ProximoCentinela() As Byte
 
    ' @ Devuelve un slot para un centinela.
 
    Dim i As Long
 
    For i = 1 To NUM_CENTINELAS

        'Si no esta invocado.
        If Not Centinelas(i).Invocado Then
            'Devuelve el slot
            ProximoCentinela = CByte(i)
            Exit Function

        End If

    Next i
 
    ProximoCentinela = 0
 
End Function
 
Function CheckCodigo(ByRef Ingresada As String, ByVal CIndex As Byte) As Boolean
 
    ' @ Devuelve si el codigo es correcto.
 
    CheckCodigo = (Not Ingresada <> Centinelas(CIndex).CodigoCheck)
 
End Function
