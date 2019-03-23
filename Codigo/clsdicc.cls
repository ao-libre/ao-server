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
'**************************************************************
' diccionario.cls
'
' Designed and implemented by Mariono Barrou (El Oso)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

'clase diccionario
'basico, plain sin queso ni papa fritas

'mi idea cuando hice esto, lo encontre en el rigido :p. Hecha por el oso

Option Explicit

Private Const MAX_ELEM = 100

Private Type diccElem

    Clave As String
    def As Variant

End Type

Private p_elementos(1 To MAX_ELEM) As diccElem 'visual basic es una mierda para usar memoria dinamica, asi que uso esto

Private p_cant                     As Integer

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

Public Function AtPut(ByVal Clave As String, ByRef elem As Variant) As Boolean

    Dim i As Integer

    AtPut = False

    If LenB(Clave) = 0 Then Exit Function

    Clave = UCase$(Clave)

    If p_cant = MAX_ELEM Then
        AtPut = False
    Else

        For i = 1 To p_cant

            If Clave = p_elementos(i).Clave Then
                p_elementos(i).def = elem
                AtPut = True

                Exit For    ' epa ;)

            End If

        Next i

        If Not AtPut Then
            p_cant = p_cant + 1
            p_elementos(p_cant).def = elem
            p_elementos(p_cant).Clave = Clave
            AtPut = True

        End If

    End If

End Function

Public Function At(ByVal Clave As String) As Variant

    Dim i As Integer

    Clave = UCase$(Clave)

    At = Null

    For i = 1 To p_cant

        If Clave = p_elementos(i).Clave Then
            At = p_elementos(i).def
            Exit Function

        End If

    Next i

End Function

Public Function AtIndex(ByVal i As Integer) As String
    AtIndex = p_elementos(i).Clave

End Function

Public Function MayorValor(ByRef cant As Integer) As String

    'parchecito para el AO, me da la clave con mayor valor en valor
    'y la cantidad de claves con ese valor (por si hay empate)
    Dim i     As Integer

    Dim max   As Integer

    Dim Clave As String

    max = -1
    cant = 0
    Clave = vbNullString

    For i = 1 To p_cant

        If max <= CInt(p_elementos(i).def) Then
            cant = IIf(max = CInt(p_elementos(i).def), cant + 1, 1)
            Clave = IIf(max = CInt(p_elementos(i).def), Clave & "," & p_elementos(i).Clave, p_elementos(i).Clave)
            max = CInt(p_elementos(i).def)

        End If

    Next i

    MayorValor = Clave

End Function

Public Sub DumpAll()

    Dim i As Integer

    For i = 1 To MAX_ELEM
        p_elementos(i).Clave = vbNullString
        p_elementos(i).def = Null
    Next i

    p_cant = 0

End Sub
