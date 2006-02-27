Attribute VB_Name = "Matematicas"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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

Sub AddtoVar(ByRef Var As Variant, ByVal Addon As Variant, ByVal max As Variant)
'Le suma un valor a una variable respetando el maximo valor

If Var >= max Then
    Var = max
Else
    Var = Var + Addon
    If Var > max Then
        Var = max
    End If
End If

End Sub


Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
Porcentaje = (Total * Porc) / 100
End Function

Public Function SD(ByVal N As Integer) As Integer
'Call LogTarea("Function SD n:" & n)
'Suma digitos
Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer

auxint = N

Do
    
    digit = (auxint Mod 10)
    suma = suma + digit
    auxint = auxint \ 10
    
Loop While (auxint > 0)

SD = suma

End Function

Public Function SDM(ByVal N As Integer) As Integer
'Call LogTarea("Function SDM n:" & n)
'Suma digitos cada digito menos dos
Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer

auxint = N
'If auxint < 0 Then auxint = Abs(auxint)

Do
    
    digit = (auxint Mod 10)
    digit = digit - 1
    suma = suma + digit
    auxint = auxint \ 10
    
   
Loop While (auxint > 0)

SDM = suma

End Function

Public Function Complex(ByVal N As Integer) As Integer
'Call LogTarea("Complex")

If N Mod 2 <> 0 Then
    Complex = N * SD(N)
Else
    Complex = N * SDM(N)
End If

End Function

Function Distancia(wp1 As WorldPos, wp2 As WorldPos)

'Encuentra la distancia entre dos WorldPos

Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)

End Function

Function Distance(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Double

'Encuentra la distancia entre dos puntos

Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

'Lo puse en sub Main()
'Randomize Timer

'RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
'If RandomNumber > UpperBound Then RandomNumber = UpperBound

RandomNumber = Int(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function

