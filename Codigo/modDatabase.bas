Attribute VB_Name = "modDatabase"
'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018

Option Explicit

Public Const Database_Enabled As Boolean = False
Public Database_Host As String
Public Database_Name As String
Public Database_Username As String
Public Database_Password As String
Public Database_Connection As ADODB.Connection
Public Database_ResultSet As ADODB.Recordset
 
Public Sub Database_Connect()
'***************************************************
'Author: Juan Andres Dalmasso
'Last Modification: 18/09/2018
'***************************************************
On Error GoTo ErrorHandler
 
Set Database_Connection = New ADODB.Connection
 
Database_Connection.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Database_Host & ";DATABASE=" & Database_Name & ";UID=" & Database_Username & ";PWD=" & Database_Password & "; OPTION=3"
Database_Connection.CursorLocation = adUseClient
Database_Connection.Open

Exit Sub
ErrorHandler:
    Call LogCriticEvent("Unable to connect to Mysql Database: " & Err.Number & " - " & Err.description)
End Sub

Public Sub Database_Close()
'***************************************************
'Author: Juan Andres Dalmasso
'Last Modification: 18/09/2018
'***************************************************
On Error GoTo ErrorHandler
     
Database_Connection.Close
Set Database_Connection = Nothing
     
Exit Sub
     
ErrorHandler:
    Call LogCriticEvent("Unable to close Mysql Database: " & Err.Number & " - " & Err.description)
End Sub
