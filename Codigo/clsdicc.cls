VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "diccionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clase diccionario
'basico, plain sin queso ni papa fritas


'mi idea cuando hice esto, lo encontre en el rigido :p. Hecha por el oso

Option Explicit

Private Const MAX_ELEM = 100

Private Type diccElem
    clave As String
    def As Variant
End Type

Private p_elementos(1 To MAX_ELEM) As diccElem 'visual basic es una mierda para usar memoria dinamica, asi que uso esto
Private p_cant As Integer

Public Sub Class_Initialize()
'constructor
    p_cant = 0
End Sub

Public Sub Class_Terminate()
'destructor
    'destruir los variants?????
End Sub

Public Property Get CantElem() As Integer
    CantElem = p_cant
End Property

Public Function AtPut(ByVal clave As String, ByRef elem As Variant) As Boolean
Dim i As Integer

AtPut = False

If clave = "" Then Exit Function

clave = UCase$(clave)

If p_cant = MAX_ELEM Then
    AtPut = False
Else
    For i = 1 To p_cant
        If clave = p_elementos(i).clave Then
            p_elementos(i).def = elem
            AtPut = True
            Exit For    ' epa ;)
        End If
    Next i
    If Not AtPut Then
        p_cant = p_cant + 1
        p_elementos(p_cant).def = elem
        p_elementos(p_cant).clave = clave
        AtPut = True
    End If

End If
End Function

Public Function At(ByVal clave As String) As Variant
Dim i As Integer

clave = UCase$(clave)

    At = Null
    For i = 1 To p_cant
        If clave = p_elementos(i).clave Then
            At = p_elementos(i).def
            Exit Function
        End If
    Next i

End Function

Public Function AtIndex(ByVal i As Integer) As String
    AtIndex = p_elementos(i).clave
End Function


Public Function MayorValor(ByRef Cant As Integer) As String
'parchecito para el AO, me da la clave con mayor valor en valor
'y la cantidad de claves con ese valor (por si hay empate)
Dim i As Integer
Dim max As Integer
Dim clave As String
    max = -1
    Cant = 0
    clave = vbNullString
    For i = 1 To p_cant
        If max <= CInt(p_elementos(i).def) Then
            Cant = IIf(max = CInt(p_elementos(i).def), Cant + 1, 1)
            clave = IIf(max = CInt(p_elementos(i).def), clave & "," & p_elementos(i).clave, p_elementos(i).clave)
            max = CInt(p_elementos(i).def)
        End If
    Next i

    MayorValor = clave

End Function

Public Sub DumpAll()
Dim i As Integer

For i = 1 To MAX_ELEM
    p_elementos(i).clave = ""
    p_elementos(i).def = Null
Next i
p_cant = 0

End Sub
