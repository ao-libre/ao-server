Attribute VB_Name = "ApiConnection"
'**************************************************************
'Autor: Lucas Recoaro (Recox)
'Este modulo es el encargado de hacer los requests a los diferentes endpoints de la API
'La API esta escrita en javascript (Node.js/Express) y podemos desde obtener datos a hacer backup de charfiles o cuentas
'https://github.com/ao-libre/ao-api-server
'**************************************************************

Option Explicit
Private XmlHttp As Object
Private UrlServer As String

Public Sub ApiEndpointBackupCharfiles()

    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint hace una copia de todos los charfiles a una base de datos mysql
    'No todos los parametros estan incluidos, es mas que nada para usar de rankings
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "GET", UrlServer & "/api/v1/charfiles/backupcharfiles", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send
End Sub

Public Sub ApiEndpointBackupCuentas()

    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint hace una copia de todos las cuentas a una base de datos mysql
    'Es mas que nada para poder hacer cosas con los usuarios
    'De forma mas facil en javascript
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "GET", UrlServer & "/api/v1/accounts/backupaccountfiles", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send
End Sub

Public Sub ApiEndpointBackupLogs()

    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint hace una copia de todos los logs a una base de datos mysql
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "GET", UrlServer & "/api/v1/logs/backuplogs", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send
End Sub

Public Sub ApiEndpointSendWelcomeEmail(ByVal UserName As String, ByVal Password As String, ByVal Email As String)

    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint envia un email de bienvenida al usuario, con su nombre de usuario y password para que no lo pierda :)
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "POST", UrlServer & "/api/v1/emails/welcome", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send "username=" & UserName & "&password=" & Password & "&emailTo=" & Email
End Sub

Public Sub ApiEndpointSendLoginAccountEmail(ByVal Email As String)

    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint envia un email de bienvenida al usuario, con su nombre de usuario y password para que no lo pierda :)
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "POST", UrlServer & "/api/v1/emails/loginAccount", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send "emailTo=" & Email
End Sub

Public Sub ApiEndpointSendResetPasswordAccountEmail(ByVal Email As String, ByVal NewPassword As String)

    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint envia un email de reset password al email de la cuenta
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "POST", UrlServer & "/api/v1/emails/resetAccountPassword", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send "newPassword=" & NewPassword & "&emailTo=" & Email
End Sub
