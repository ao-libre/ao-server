VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
'                    Metodos publicos
'
' Public sub Push(byval i as variant) mete el elemento i
' al final de la cola.
'
' Public Function Pop As Variant: quita de la cola el primer elem
' y lo devuelve
'
' Public Function VerElemento(ByVal Index As Integer) As Variant
' muestra el elemento numero Index de la cola sin quitarlo
'
' Public Function PopByVal() As Variant: muestra el primer
' elemento de la cola sin quitarlo
'
' Public Property Get Longitud() As Integer: devuelve la
' cantidad de elementos que tiene la cola.

Private Const FRENTE = 1

Private MaxElem As Integer

Private Cola As Collection

Public Property Let MaxLenght(ByVal Max As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Sets the max queue lenght.
'***************************************************
    MaxElem = Max
End Property

Public Sub Reset()
On Error Resume Next

Dim i As Integer
For i = 1 To Me.Longitud
    Cola.Remove FRENTE
Next i

End Sub

Public Property Get Longitud() As Integer
Longitud = Cola.Count
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
IndexValido = i >= 1 And i <= Me.Longitud
End Function

Private Sub Class_Initialize()
Set Cola = New Collection
End Sub

Public Function VerElemento(ByVal index As Integer, Optional ByVal ConvertUpper As Boolean = True) As String
On Error Resume Next
    If IndexValido(index) Then
        'Pablo
        If ConvertUpper Then
            VerElemento = UCase$(Cola.Item(index))
        Else
            VerElemento = Cola.Item(index)
        End If
        
        '/Pablo
        'VerElemento = Cola(Index)
    Else
        VerElemento = 0
    End If
End Function


Public Sub Push(ByVal Detalle As String, Optional ByVal ConvertUpper As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: 14/11/2010
'14/11/2010: ZaMa - Ahora se valida si supera un maximo previamente establecido.
'***************************************************
On Error Resume Next
    'Mete elemento en la cola
    
    Dim aux As String
    aux = time$ & " "
    
    If ConvertUpper Then
        aux = aux & UCase$(Detalle)
    Else
        aux = aux & Detalle
    End If
    
    ' Validate that not exceed max lenght
    If MaxElem <> 0 Then
        If Cola.Count = MaxElem Then
            Call Cola.Remove(FRENTE) ' Quito el primer elemento
        End If
    End If
    
    Call Cola.Add(aux)
        
End Sub

Public Function Pop() As String
On Error Resume Next
'Quita elemento de la cola
If Cola.Count > 0 Then
    Pop = Cola(FRENTE)
    Call Cola.Remove(FRENTE)
Else
    Pop = 0
End If
End Function

Public Function PopByVal() As String
On Error Resume Next
'Call LogTarea("PopByVal SOS")

'Quita elemento de la cola
If Cola.Count > 0 Then
    PopByVal = Cola.Item(1)
Else
    PopByVal = 0
End If

End Function

Public Function Existe(ByVal Nombre As String) As Boolean
On Error Resume Next
  
 Dim V As String
 Dim i As Integer
 Dim NombreEnMayusculas As String
 NombreEnMayusculas = UCase$(Nombre)
 
 For i = 1 To Me.Longitud
        'Pablo
        V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
        '/Pablo
        'V = Me.VerElemento(i)
        If V = NombreEnMayusculas Then
            Existe = True
            Exit Function
        End If
 Next
 Existe = False

End Function

Public Sub Quitar(ByVal Nombre As String)
On Error Resume Next
Dim V As String
Dim i As Integer
Dim NombreEnMayusculas As String
 
NombreEnMayusculas = UCase$(Nombre)
 
For i = 1 To Me.Longitud
    'Pablo
    V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
    '/Pablo
    'V = Me.VerElemento(i)
    If V = NombreEnMayusculas Then
        Call Cola.Remove(i)
        Exit Sub
    End If
Next i
 
End Sub

Public Sub QuitarIndex(ByVal index As Integer)
On Error Resume Next
If IndexValido(index) Then Call Cola.Remove(index)
End Sub


Private Sub Class_Terminate()
'Destruimos el objeto Cola
Set Cola = Nothing
End Sub
