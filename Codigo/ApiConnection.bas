Attribute VB_Name = "ApiConnection"
'**************************************************************
'Autor: Lucas Recoaro (Recox)
'Este modulo es el encargado de hacer los requests a los diferentes endpoints de la API
'La API esta escrita en javascript (node.js/express) y podemos desde obtener datos a hacer backup de charfiles o cuentas
'https://github.com/ao-libre/ao-api-server
'**************************************************************

Option Explicit

Private Sub MakeRequestToEndpoint(ByVal Endpoint As String)
    'Esto tendria que crearse en este modulo, trate pero no pude asi que hago esta solucion para poder avanzar
    'No se como crear el component Inet aqui, dejo lo que intente aunque no funciona
    'Si solucionan eso, hay que borrar el InetApi del frmMain, aunque asi funciona barbaro
    'Dim InetApi As InetCtlsObjects.Inet
    'Set Inet = New InetCtlsObjects.Inet
    frmMain.InetApi.OpenURL (Endpoint)
End Sub

Public Sub ApiEndpointBackupCharfiles()
    Dim UrlServer As String
    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint hace una copia de todos los charfiles a una base de datos mysql
    'No todos los parametros estan incluidos, es mas que nada para usar de rankings
    MakeRequestToEndpoint (UrlServer & "/api/v1/charfiles/backupcharfiles")
End Sub

Public Sub ApiEndpointBackupCuentas()
    Dim UrlServer As String
    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint hace una copia de todos las cuentas a una base de datos mysql
    'Es mas que nada para poder hacer cosas con los usuarios
    'De forma mas facil en javascript
    MakeRequestToEndpoint (UrlServer & "/api/v1/accounts/backupaccountfiles")
End Sub

Public Sub ApiEndpointBackupLogs()
    Dim UrlServer As String
    UrlServer = GetVar(IniPath & "Server.ini", "CONEXIONAPI", "UrlServer")
    
    'Este endpoint hace una copia de todos los logs a una base de datos mysql
    MakeRequestToEndpoint (UrlServer & "/api/v1/logs/backuplogs")
End Sub
