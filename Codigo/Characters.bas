Attribute VB_Name = "Characters"
'**************************************************************
' Characters.bas - library of functions to manipulate characters.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
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

Option Explicit

''
' Value representing invalid indexes.
Public Const INVALID_INDEX As Integer = 0

''
' Retrieves the UserList index of the user with the give char index.
'
' @param    CharIndex   The char index being used by the user to be retrieved.
' @return   The index of the user with the char placed in CharIndex or INVALID_INDEX if it's not a user or valid char index.
' @see      INVALID_INDEX

Public Function CharIndexToUserIndex(ByVal CharIndex As Integer) As Integer
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Takes a CharIndex and transforms it into a UserIndex. Returns INVALID_INDEX in case of error.
'***************************************************
    CharIndexToUserIndex = CharList(CharIndex)
    
    If CharIndexToUserIndex < 1 Or CharIndexToUserIndex > MaxUsers Then
        CharIndexToUserIndex = INVALID_INDEX
        Exit Function
    End If
    
    If UserList(CharIndexToUserIndex).Char.CharIndex <> CharIndex Then
        CharIndexToUserIndex = INVALID_INDEX
        Exit Function
    End If
End Function
