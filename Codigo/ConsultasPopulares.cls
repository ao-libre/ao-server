VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultasPopulares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Modulo de consultas popluares
'En realidad el modulo inicial guardaba los datos de las votaciones
'en memoria, pero no le vi el punto, las votaciones son de vez en cuando nomas
'y el query a un .ini que registra todos los mails q ya votaron, es soportable
'si cuando uno vota y averiguo que el mail ya voto pero el PJ no, entonces seteo
'el flag de yavoto en el charfile ese tambien,
'por lo que la busqueda secuencial en el .dat que tiene todos los mails q ya votaron
'se hara .. 1 vez por PJ nomas.
'
'Hecha por el oso

Private Const ARCHIVOMAILS = "\logs\votaron.dat"
Private Const ARCHIVOCONFIG = "\dat\consultas.dat"

Private pEncuestaActualNum As Integer
Private pEncuestaActualTex As String
Private pNivelRequerido As Integer
Private pOpciones() As Integer


Public Property Let Numero(ByVal NumEncuesta As Integer)
    pEncuestaActualNum = NumEncuesta
End Property

Public Property Get Numero() As Integer
    Numero = pEncuestaActualNum
End Property

Public Property Let texto(ByVal Descripcion As String)
    pEncuestaActualTex = Descripcion
End Property

Public Property Get texto() As String
    texto = pEncuestaActualTex
End Property


Public Sub LoadData()
Dim CantOpciones As Integer
Dim i As Integer

    pEncuestaActualNum = val(GetVar(App.Path & ARCHIVOCONFIG, "INIT", "ConsultaActual"))
    pEncuestaActualTex = GetVar(App.Path & ARCHIVOCONFIG, "INIT", "ConsultaActualTexto")
    pNivelRequerido = GetVar(App.Path & ARCHIVOCONFIG, "INIT", "NivelRequerido")

    If pEncuestaActualNum > 0 Then
       'cargo todas las opciones
        CantOpciones = val(GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "CANTOPCIONES"))
        ReDim pOpciones(1 To CantOpciones) As Integer
        For i = 1 To CantOpciones
            pOpciones(i) = val(GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & i))
        Next i
    End If
End Sub

Public Function doVotar(ByVal UserIndex As Integer, opcion As Integer) As String
On Error GoTo errorh
Dim YaVoto As Boolean
Dim CharFile As String
Dim sufragio As Integer

    'revisar q no haya votado
    'grabar en el charfile el numero de encuesta
    'actualizar resultados encuesta
    If pEncuestaActualNum = 0 Then
        doVotar = "No hay consultas populares abiertas"
        Exit Function
    End If

    CharFile = CharPath & UserList(UserIndex).name & ".chr"
    
    
    
    If (UserList(UserIndex).Stats.ELV >= pNivelRequerido) Then
        If (OpcionValida(opcion)) Then
            YaVoto = val(GetVar(CharFile, "CONSULTAS", "Voto")) >= pEncuestaActualNum
            If Not YaVoto Then
                If Not MailYaVoto(UserList(UserIndex).email) Then
                    'pj apto para votar
                    sufragio = CLng(val(GetVar(App.Path & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion)))
                    sufragio = sufragio + 1
                    Call WriteVar(App.Path & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion, str(sufragio))
                    doVotar = "Tu voto ha sido computado. Opcion: " & opcion
                    Call MarcarPjComoQueYaVoto(UserIndex)
                    Call MarcarMailComoQueYaVoto(UserList(UserIndex).email)
                Else
                    Call MarcarPjComoQueYaVoto(UserIndex)
                    doVotar = "Este email ya voto en la consulta: " & pEncuestaActualTex
                End If
            Else
                doVotar = "Este personaje ya voto en la consulta: " & pEncuestaActualTex
            End If
        Else
            doVotar = "Esa no es una opcion para votar"
        End If
    Else
        doVotar = "Para votar en esta consulta debes ser nivel " & pNivelRequerido & " o superior"
    End If
   

Exit Function
errorh:
Call LogError("Error en ConsultasPopularse.doVotar: " & Err.Description)

End Function


Public Function SendInfoEncuesta(ByVal UserIndex As Integer) As String
Dim i As Integer
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||CONSULTA POPULAR NUMERO " & pEncuestaActualNum & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & pEncuestaActualTex & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Opciones de voto: " & FONTTYPE_GUILDMSG)
    For i = 1 To UBound(pOpciones)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||(Opcion " & i & "): " & GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & i) & FONTTYPE_GUILDMSG)
    Next i
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1. Tu voto no podra ser cambiado." & FONTTYPE_VENENO)
End Function


Private Sub MarcarPjComoQueYaVoto(ByVal UserIndex As Integer)
    Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "CONSULTAS", "Voto", str(pEncuestaActualNum))
End Sub


Private Function MailYaVoto(ByVal email As String) As Boolean
'abro el archivo, while not eof levnato 1 linea y comparo. Si da true, cierro
Dim ArchN As Integer
Dim Tmp As String

MailYaVoto = False

ArchN = FreeFile()

Open App.Path & ARCHIVOMAILS For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    If email = Tmp Then
        MailYaVoto = True
        Close #ArchN
        Exit Function
    End If
Loop

Close #ArchN
End Function


Private Sub MarcarMailComoQueYaVoto(ByVal email As String)
Dim ArchN As Integer

ArchN = FreeFile()

Open App.Path & ARCHIVOMAILS For Append As #ArchN
Print #ArchN, email

Close #ArchN

End Sub


Private Function OpcionValida(ByVal opcion As Integer) As Boolean
    OpcionValida = opcion > 0 And opcion <= UBound(pOpciones)
End Function

Private Function DameOpcion(ByVal opcion As Integer) As String
    If OpcionValida(opcion) Then
        DameOpcion = GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & opcion)
    Else
        DameOpcion = "ERROR"
    End If
End Function
