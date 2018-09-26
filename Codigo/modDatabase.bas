Attribute VB_Name = "modDatabase"
'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018

Option Explicit

Public Const Database_Enabled as Boolean = False
Public Database_Host as String
Public Database_Name as String
Public Database_Username as String
Public Database_Password as String
Public Database_Connection As ADODB.Connection
Public Database_ResultSet As ADODB.Recordset
 
Public Sub Database_Connect()
'***************************************************
'Author: Juan Andres Dalmasso
'Last Modification: 18/09/2018
'***************************************************
On Error GoTo ErrorHandler
 
Set Con = New ADODB.Connection
 
Con.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Database_Host & ";DATABASE=" & Database_Name & ";"UID=" & Database_Username & ";PWD=" & Database_Password & "; OPTION=3"
Con.CursorLocation = adUseClient
Con.Open

Exit Sub
ErrorHandler:
    Call LogCriticEvent("Unable to connect to Mysql Database: " & Err.number & " - " Err.Description)
End Sub

Public Sub Database_Close()
'***************************************************
'Author: Juan Andres Dalmasso
'Last Modification: 18/09/2018
'***************************************************
On Error GoTo ErrorHandler
     
Con.Close
Set Con = Nothing
     
Exit Sub
     
ErrorHandler:
    Call LogCriticEvent("Unable to close Mysql Database: " & Err.number & " - " Err.Description) 
End Sub