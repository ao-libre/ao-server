Attribute VB_Name = "modHexaStrings"
'Argentum Online 0.12.2
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

'Modulo realizado por Gonzalo Larralde(CDT) <gonzalolarralde@yahoo.com.ar>
'Para la conversion a caracteres de cadenas MD5 y de
'semi encriptaci�n de cadenas por ascii table offset

Option Explicit

Public Function hexMd52Asc(ByVal MD5 As String) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Long
    Dim L As String
    
    If Len(MD5) And &H1 Then MD5 = "0" & MD5
    
    For i = 1 To Len(MD5) \ 2
        L = mid$(MD5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr$(hexHex2Dec(L))
    Next i
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    hexHex2Dec = val("&H" & hex)
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Long
    Dim L As String
    
    For i = 1 To Len(Text)
        L = mid$(Text, i, 1)
        txtOffset = txtOffset & Chr$((Asc(L) + off) And &HFF)
    Next i
End Function
