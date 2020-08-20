Attribute VB_Name = "ApiConnection"
'**************************************************************
'Autor: Lucas Recoaro (Recox)
'Este modulo es el encargado de hacer los requests a los diferentes endpoints de la API
'La API esta escrita en javascript (Node.js/Express) y podemos desde obtener datos a hacer backup de charfiles o cuentas
'https://github.com/ao-libre/ao-api-server
'**************************************************************

Option Explicit
Private XmlHttp As Object
Private Endpoint As String
Private Parameters As String

Public Sub ApiEndpointBackupCharfiles()
    'Este endpoint hace una copia de todos los charfiles a una base de datos mysql
    'No todos los parametros estan incluidos, es mas que nada para usar de rankings
    
    Endpoint = ApiUrlServer & "/api/v1/charfiles/backupcharfiles"
    
    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointBackupCuentas()
    'Este endpoint hace una copia de todos las cuentas a una base de datos mysql
    'Es mas que nada para poder hacer cosas con los usuarios
    'De forma mas facil en javascript
    
    Endpoint = ApiUrlServer & "/api/v1/accounts/backupaccountfiles"
    
    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointBackupLogs()
    'Este endpoint hace una copia de todos los logs a una base de datos mysql
    
    Endpoint = ApiUrlServer & "/api/v1/logs/backuplogs"

    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointSendWelcomeEmail(ByVal UserName As String, ByVal Password As String, ByVal Email As String)
    'Este endpoint envia un email de bienvenida al usuario, con su nombre de usuario y password para que no lo pierda :)
    
    Endpoint = ApiUrlServer & "/api/v1/emails/welcome"
    
    Parameters = "username=" & UserName & "&password=" & Password & "&emailTo=" & Email

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendLoginAccountEmail(ByVal Email As String, ByVal LastIpsUsed As String, ByVal CurrentIp As String)
    'Este endpoint envia un email de login cada ves que un usuario se conecta
    Endpoint = ApiUrlServer & "/api/v1/emails/loginAccount"
    
    Parameters = "emailTo=" & Email & "&lastIpsUsed=" & LastIpsUsed & "&currentIp=" & CurrentIp

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendResetPasswordAccountEmail(ByVal Email As String, ByVal NewPassword As String)
    'Este endpoint envia un email para cambiar password al usuario

    Endpoint = ApiUrlServer & "/api/v1/emails/resetAccountPassword"

    Parameters = "newPassword=" & NewPassword & "&emailTo=" & Email

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendUserConnectedMessageDiscord(ByVal UserName As String, ByVal desc As String, ByVal EsCriminal As Boolean, ByVal Clase As String)
    'Este endpoint envia un mensaje al chat avisando que alguien se conecto

    Endpoint = ApiUrlServer & "/api/v1/discord/sendConnectedMessage"
    Parameters = "userName=" & UserName & "&desc=" & desc & "&esCriminal=" & EsCriminal & "&clase=" & Clase

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendHappyHourStartedMessageDiscord(ByVal Message As String)
    'Este endpoint envia un mensaje al chat avisando que empezo el happy hour

    Endpoint = ApiUrlServer & "/api/v1/discord/sendHappyHourStartMessage"
    Parameters = "message=" & Message

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendHappyHourEndedMessageDiscord(ByVal Message As String)
    'Este endpoint envia un mensaje al chat avisando que termino el happy hour

    Endpoint = ApiUrlServer & "/api/v1/discord/sendHappyHourEndMessage"
    Parameters = "message=" & Message

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendHappyHourModifiedMessageDiscord(ByVal Message As String)
    'Este endpoint envia un mensaje al chat avisando que termino el happy hour

    Endpoint = ApiUrlServer & "/api/v1/discord/sendHappyHourModifiedMessage"
    Parameters = "message=" & Message

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendNewGuildCreatedMessageDiscord(ByVal Message As String, ByVal Desc As String, ByVal GuildName As String, ByVal Site As String)
    'Este endpoint envia un mensaje al chat avisando que se creo un clan

    Endpoint = ApiUrlServer & "/api/v1/discord/sendNewGuildCreated"
    Parameters = "message=" & Message & "&desc=" & Desc & "&guildname=" & Guildname & "&site=" & Site

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendCustomCharacterMessageDiscord(ByVal Chat As String, ByVal Name As String, ByVal desc As String)
    'Este endpoint envia un mensaje al discord desde dentro del juego por un usuario con el comando /discord
    Endpoint = ApiUrlServer & "/api/v1/discord/sendCustomCharacterMessageDiscord"
    Parameters = "userName=" & Name & "&desc=" & desc & "&chat=" & Chat

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub


Public Sub ApiEndpointSendWorldSaveMessageDiscord()
    'Este endpoint envia un mensaje al chat avisando que se creo un clan

    Endpoint = ApiUrlServer & "/api/v1/discord/sendWorldSaveMessage"
    Parameters = ""

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendCreateNewCharacterMessageDiscord(ByVal Name As String)
    'Este endpoint envia un mensaje al chat avisando que se creo un clan

    Endpoint = ApiUrlServer & "/api/v1/discord/sendCreatedNewCharacterMessage"
    Parameters = "name=" & Name

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendServerDataToApiToShowOnlineUsers()
    'Este endpoint envia estadisticas de cuantos online hay a la API para que se puedan ver en el bot de discord entre otros...
    'Este endpoint tiene hardcodeada la url de la api por que no es la interna, se usa para estadisticas y mostrar cuantos ons hay que cada server (Recox)
    Endpoint = "https://api.argentumonline.org/api/v1/servers/sendUsersOnline"
    Parameters = "serverName=" & NombreServidor & "&quantityUsers=" & LastUser & "&ip=" & IpPublicaServidor & "&port=" & Puerto 

    Call SendPOSTRequest(Endpoint, Parameters)
End Sub


Private Sub SendPOSTRequest(ByVal Endpoint As String, ByVal Parameters As String)

On Error GoTo ErrorHandler

    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "POST", Endpoint, True
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        
    'Por alguna razon tengo que castearlo a string, sino no funciona, la verdad no tengo idea por que ya que la variable es String
    XmlHttp.send CStr(Parameters)
    
    Set XmlHttp = Nothing

ErrorHandler:

    If Err.Number <> 0 Then
        Call LogError("Error POST endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.description)
    End If
    
End Sub

Private Sub SendGETRequest(ByVal Endpoint As String)
On Error GoTo ErrorHandler

    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "GET", Endpoint, True
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send

    Set XmlHttp = Nothing

ErrorHandler:

    If Err.Number <> 0 Then
        Call LogError("Error GET endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.description)
    End If
    
End Sub
