Attribute VB_Name = "PraetoriansCoopNPC"

''**************************************************************
'' PraetoriansCoopNPC.bas - Handles the Praeorians NPCs.
''
'' Implemented by Mariano Barrou (El Oso)
''**************************************************************
'
''**************************************************************************
''This program is free software; you can redistribute it and/or modify
''it under the terms of the Affero General Public License;
''either version 1 of the License, or any later version.
''
''This program is distributed in the hope that it will be useful,
''but WITHOUT ANY WARRANTY; without even the implied warranty of
''MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''Affero General Public License for more details.
''
''You should have received a copy of the Affero General Public License
''along with this program; if not, you can find it at http://www.affero.org/oagpl.html
''**************************************************************************
'
'Option Explicit
''''''''''''''''''''''''''''''''''''''''''
''' DECLARACIONES DEL MODULO PRETORIANO ''
''''''''''''''''''''''''''''''''''''''''''
''' Estas constantes definen que valores tienen
''' los NPCs pretorianos en el NPC-HOSTILES.DAT
''' Son FIJAS, pero se podria hacer una rutina que
''' las lea desde el npcshostiles.dat
'Public Const PRCLER_NPC As Integer = 900   ''"Sacerdote Pretoriano"
'Public Const PRGUER_NPC As Integer = 901   ''"Guerrero  Pretoriano"
'Public Const PRMAGO_NPC As Integer = 902   ''"Mago Pretoriano"
'Public Const PRCAZA_NPC As Integer = 903   ''"Cazador Pretoriano"
'Public Const PRKING_NPC As Integer = 904   ''"Rey Pretoriano"
'
'
'' 1 rey.
'' 3 guerres.
'' 1 caza.
'' 1 mago.
'' 2 clerigos.
'Public Const NRO_PRETORIANOS As Integer = 8
'
'' Contiene los index de los miembros del clan.
'Public Pretorianos(1 To NRO_PRETORIANOS) As Integer
'
'
''''''''''''''''''''''''''''''''''''''''''''''
''Esta constante identifica en que mapa esta
''la fortaleza pretoriana (no es lo mismo de
''donde estan los NPCs!).
''Se extrae el dato del server.ini en sub LoadSIni
Public MAPA_PRETORIANO          As Integer
Public PRETORIANO_X             As Byte
Public PRETORIANO_Y             As Byte

''''''''''''''''''''''''''''''''''''''''''''''
''Estos numeros son necesarios por cuestiones de
''sonido. Son los numeros de los wavs del cliente.
Public Const SONIDO_DRAGON_VIVO As Integer = 30

'''ALCOBAS REALES
'''OJO LOS BICHOS TAN HARDCODEADOS, NO CAMBIAR EL MAPA DONDE
'''ESTaN UBICADOS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'''MUCHO MENOS LA COORDENADA Y DE LAS ALCOBAS YA QUE DEBE SER LA MISMA!!!
'''(HAY FUNCIONES Q CUENTAN CON QUE ES LA MISMA!)
'Public Const ALCOBA1_X As Integer = 35
'Public Const ALCOBA1_Y As Integer = 25
'Public Const ALCOBA2_X As Integer = 67
'Public Const ALCOBA2_Y As Integer = 25

Public Enum ePretorianAI

    King = 1
    Healer
    SpellCaster
    SwordMaster
    Shooter
    Thief
    Last

End Enum

' Contains all the pretorian's combinations, and its the offsets
Public PretorianAIOffset(1 To 7) As Integer

Public PretorianDatNumbers()     As Integer
'
''Added by Nacho
''Cuantos pretorianos vivos quedan. Uno por cada alcoba
'Public pretorianosVivos As Integer
'

Public Sub LoadPretorianData()

    Dim PretorianDat As String

    PretorianDat = DatPath & "Pretorianos.dat"

    Dim NroCombinaciones As Integer

    NroCombinaciones = val(GetVar(PretorianDat, "MAIN", "Combinaciones"))

    ReDim PretorianDatNumbers(1 To NroCombinaciones)

    Dim TempInt        As Integer

    Dim Counter        As Long

    Dim PretorianIndex As Integer

    PretorianIndex = 1

    ' KINGS
    TempInt = val(GetVar(PretorianDat, "KING", "Cantidad"))
    PretorianAIOffset(ePretorianAI.King) = 1

    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "KING", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "KING", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' HEALERS
    TempInt = val(GetVar(PretorianDat, "HEALER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Healer) = PretorianIndex

    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "HEALER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "HEALER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' SPELLCASTER
    TempInt = val(GetVar(PretorianDat, "SPELLCASTER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.SpellCaster) = PretorianIndex

    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SPELLCASTER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SPELLCASTER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' SWORDSWINGER
    TempInt = val(GetVar(PretorianDat, "SWORDSWINGER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.SwordMaster) = PretorianIndex

    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SWORDSWINGER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SWORDSWINGER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' LONGRANGE
    TempInt = val(GetVar(PretorianDat, "LONGRANGE", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Shooter) = PretorianIndex

    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "LONGRANGE", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "LONGRANGE", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' THIEF
    TempInt = val(GetVar(PretorianDat, "THIEF", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Thief) = PretorianIndex

    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "THIEF", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "THIEF", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' Last
    PretorianAIOffset(ePretorianAI.Last) = PretorianIndex

    ' Inicializa los clanes pretorianos
    ReDim ClanPretoriano(ePretorianType.Default To ePretorianType.Custom) As clsClanPretoriano
    Set ClanPretoriano(ePretorianType.Default) = New clsClanPretoriano ' Clan default
    Set ClanPretoriano(ePretorianType.Custom) = New clsClanPretoriano ' Invocable por gms

End Sub

