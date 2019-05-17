Attribute VB_Name = "ApiConnection"
'**************************************************************
'Autor: Lucas Recoaro (Recox)
'Este modulo es el encargado de hacer los requests a los diferentes endpoints de la API
'La API esta escrita en javascript (node.js/express) y podemos desde obtener datos a hacer backup de charfiles o cuentas
'https://github.com/ao-libre/ao-api-server
'**************************************************************

Option Explicit
Dim XmlHttp As Object

Public Sub ApiEndpointBackupCharfiles()
    Dim UrlServer As String
    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint hace una copia de todos los charfiles a una base de datos mysql
    'No todos los parametros estan incluidos, es mas que nada para usar de rankings
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    XmlHttp.Open "GET", UrlServer & "/api/v1/charfiles/backupcharfiles", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send
End Sub

Public Sub ApiEndpointBackupCuentas()
    Dim UrlServer As String
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
    Dim UrlServer As String
    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint hace una copia de todos los logs a una base de datos mysql
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    XmlHttp.Open "GET", UrlServer & "/api/v1/logs/backuplogs", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send
End Sub

Public Sub ApiEndpointSendWelcomeEmail(ByVal UserName As String, ByVal Password As String, ByVal Email As String)
    Dim UrlServer As String
    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint envia un email de bienvenida al usuario, con su nombre de usuario y password para que no lo pierda :)
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    XmlHttp.Open "POST", UrlServer & "/api/v1/emails/welcome", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send "username=" & UserName & "&password=" & Password & "&emailTo=" & Email
End Sub

Public Sub ApiEndpointSendLoginAccountEmail(ByVal Email As String)
    Dim UrlServer As String
    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint envia un email de bienvenida al usuario, con su nombre de usuario y password para que no lo pierda :)
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    XmlHttp.Open "POST", UrlServer & "/api/v1/emails/accountLogin", False
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send "emailTo=" & Email
End Sub
