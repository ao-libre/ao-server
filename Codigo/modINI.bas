Attribute VB_Name = "ModIni"

Option Explicit

''Modulo para utilizar la DLL LeeInis.dll
''Creado por Alejandro Santos
'
''int DLLIMPORT INIDarError();
''unsigned int DLLIMPORT INICarga (const char* arch);
''unsigned int DLLIMPORT INIDescarga (unsigned int pa);
''int DLLIMPORT INIDarNumSecciones(unsigned int a);
''int DLLIMPORT INIDarNombreSeccion(unsigned int a, int n, char* buff, int tam);
''int DLLIMPORT INIBuscarSeccion(unsigned int a, const char* buff);
''long DLLIMPORT INIDarClave(unsigned int a, long n, const char* clave, const char* buff, long tam);
''long DLLIMPORT INIDarClaveInt(unsigned int a, long n, const char* clave);
''long DLLIMPORT INIDarNumClaves(unsigned int a, long n);
''long DLLIMPORT INIDarNombreClave(unsigned int a, long n, long clave, char* buff, long tam);
''long DLLIMPORT INIConf(unsigned int a, long defint, const char* defstr, long casesensitive);
'
'Public Declare Function INICarga Lib "LeeInis.dll" (ByVal Arch As String) As Long
'Public Declare Function INIDescarga Lib "LeeInis.dll" (ByVal A As Long) As Long
'Public Declare Function INIDarError Lib "LeeInis.dll" () As Long
'
'Public Declare Function INIDarNumSecciones Lib "LeeInis.dll" (ByVal A As Long) As Long
'Public Declare Function INIDarNombreSeccion Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long, ByVal Buff As String, ByVal Tam As Long) As Long
'Public Declare Function INIBuscarSeccion Lib "LeeInis.dll" (ByVal A As Long, ByVal Buff As String) As Long
'
'Public Declare Function INIDarClave Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long, ByVal Clave As String, ByVal Buff As String, ByVal Tam As Long) As Long
'Public Declare Function INIDarClaveInt Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long, ByVal Clave As String) As Long
'Public Declare Function INIDarNumClaves Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long) As Long
'Public Declare Function INIDarNombreClave Lib "LeeInis.dll" (ByVal A As Long, ByVal N As Long, ByVal Clave As Long, ByVal Buff As String, ByVal Tam As Long) As Long
'
'Public Declare Function INIConf Lib "LeeInis.dll" (ByVal A As Long, ByVal DefectoInt As Long, ByVal DefectoStr As String, ByVal CaseSensitive As Long) As Long
'
'
'Public Function INIDarClaveStr(A As Long, Seccion As Long, Clave As String) As String
'Dim Tmp As String
'Dim P As Long, R As Long
'
'Tmp = Space(3000)
'R = INIDarClave(A, Seccion, Clave, Tmp, 3000)
'P = InStr(1, Tmp, Chr(0))
'If P > 0 Then
'    Tmp = Left(Tmp, P - 1)
'    'Tmp = Replace(Tmp, Chr(34), "")
'    INIDarClaveStr = Tmp
'End If
'
'End Function
'
'Public Function INIDarNombreSeccionStr(A As Long, Seccion As Long) As String
'Dim Tmp As String
'Dim P As Long, R As Long
'
'Tmp = Space(3000)
'R = INIDarNombreSeccion(A, Seccion, Tmp, 3000)
'P = InStr(1, Tmp, Chr(0))
'If P > 0 Then
'    Tmp = Left(Tmp, P - 1)
'    INIDarNombreSeccionStr = Tmp
'End If
'
'End Function
'
'Public Function INIDarNombreClaveStr(A As Long, Seccion As Long, Clave As Long) As String
'Dim Tmp As String
'Dim P As Long, R As Long
'
'Tmp = Space(3000)
'R = INIDarNombreClave(A, Seccion, Clave, Tmp, 3000)
'P = InStr(1, Tmp, Chr(0))
'If P > 0 Then
'    Tmp = Left(Tmp, P - 1)
'    INIDarNombreClaveStr = Tmp
'End If
'
'End Function
'
