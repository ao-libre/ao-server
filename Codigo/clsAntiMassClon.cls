VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAntiMassClon"
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

Private Const MaximoPersonajesPorIP = 15
Private m_coleccion As New Collection

Public Function MaxPersonajes(sIp As String) As Boolean
    Dim i As Long

    For i = 1 To m_coleccion.Count
        If m_coleccion.Item(i).ip = sIp Then
            m_coleccion.Item(i).PersonajesCreados = m_coleccion.Item(i).PersonajesCreados + 1
            MaxPersonajes = (m_coleccion.Item(i).PersonajesCreados > MaximoPersonajesPorIP)
            If MaxPersonajes Then m_coleccion.Item(i).PersonajesCreados = 16
            Exit Function
        End If
    Next i

#If SeguridadAlkon Then
    Dim oIp As New UserIpAdress
    oIp.ip = sIp
    oIp.PersonajesCreados = 1
    m_coleccion.Add oIp
#End If

    MaxPersonajes = False
    Exit Function
End Function

Public Function VaciarColeccion()

On Error GoTo Errhandler

Dim i As Integer

For i = 1 To m_coleccion.Count
   Call m_coleccion.Remove(1)
Next


Exit Function
Errhandler:
    Call LogError("Error en RestarConexion " & Err.description)
End Function


