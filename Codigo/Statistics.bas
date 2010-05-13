Attribute VB_Name = "Statistics"
'**************************************************************
' modStatistics.bas - Takes statistics on the game for later study.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus)
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

Private Type trainningData
    startTick As Long
    trainningTime As Long
End Type

Private Type fragLvlRace
    matrix(1 To 50, 1 To 5) As Long
End Type

Private Type fragLvlLvl
    matrix(1 To 50, 1 To 50) As Long
End Type

Private trainningInfo() As trainningData

Private fragLvlRaceData(1 To 7) As fragLvlRace
Private fragLvlLvlData(1 To 7) As fragLvlLvl
Private fragAlignmentLvlData(1 To 50, 1 To 4) As Long

'Currency just in case.... chats are way TOO often...
Private keyOcurrencies(255) As Currency

Public Sub Initialize()
    ReDim trainningInfo(1 To MaxUsers) As trainningData
End Sub

Public Sub UserConnected(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'A new user connected, load it's trainning time count
    trainningInfo(UserIndex).trainningTime = val(GetVar(CharPath & UCase$(UserList(UserIndex).name) & ".chr", "RESEARCH", "TrainningTime", 30))
    
    trainningInfo(UserIndex).startTick = (GetTickCount() And &H7FFFFFFF)
End Sub

Public Sub UserDisconnected(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With trainningInfo(UserIndex)
        'Update trainning time
        .trainningTime = .trainningTime + ((GetTickCount() And &H7FFFFFFF) - .startTick) / 1000
        
        .startTick = (GetTickCount() And &H7FFFFFFF)
        
        'Store info in char file
        Call WriteVar(CharPath & UCase$(UserList(UserIndex).name) & ".chr", "RESEARCH", "TrainningTime", CStr(.trainningTime))
    End With
End Sub

Public Sub UserLevelUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim handle As Integer
    handle = FreeFile()
    
    With trainningInfo(UserIndex)
        'Log the data
        Open App.Path & "\logs\statistics.log" For Append Shared As handle
        
        Print #handle, UCase$(UserList(UserIndex).name) & " completó el nivel " & CStr(UserList(UserIndex).Stats.ELV) & " en " & CStr(.trainningTime + ((GetTickCount() And &H7FFFFFFF) - .startTick) / 1000) & " segundos."
        
        Close handle
        
        'Reset data
        .trainningTime = 0
        .startTick = (GetTickCount() And &H7FFFFFFF)
    End With
End Sub

Public Sub StoreFrag(ByVal killer As Integer, ByVal victim As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim clase As Integer
    Dim raza As Integer
    Dim alignment As Integer
    
    If UserList(victim).Stats.ELV > 50 Or UserList(killer).Stats.ELV > 50 Then Exit Sub
    
    Select Case UserList(killer).clase
        Case eClass.Assasin
            clase = 1
        
        Case eClass.Bard
            clase = 2
        
        Case eClass.Mage
            clase = 3
        
        Case eClass.Paladin
            clase = 4
        
        Case eClass.Warrior
            clase = 5
        
        Case eClass.Cleric
            clase = 6
        
        Case eClass.Hunter
            clase = 7
        
        Case Else
            Exit Sub
    End Select
    
    Select Case UserList(killer).raza
        Case eRaza.Elfo
            raza = 1
        
        Case eRaza.Drow
            raza = 2
        
        Case eRaza.Enano
            raza = 3
        
        Case eRaza.Gnomo
            raza = 4
        
        Case eRaza.Humano
            raza = 5
        
        Case Else
            Exit Sub
    End Select
    
    If UserList(killer).Faccion.ArmadaReal Then
        alignment = 1
    ElseIf UserList(killer).Faccion.FuerzasCaos Then
        alignment = 2
    ElseIf criminal(killer) Then
        alignment = 3
    Else
        alignment = 4
    End If
    
    fragLvlRaceData(clase).matrix(UserList(killer).Stats.ELV, raza) = fragLvlRaceData(clase).matrix(UserList(killer).Stats.ELV, raza) + 1
    
    fragLvlLvlData(clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) = fragLvlLvlData(clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) + 1
    
    fragAlignmentLvlData(UserList(killer).Stats.ELV, alignment) = fragAlignmentLvlData(UserList(killer).Stats.ELV, alignment) + 1
End Sub

Public Sub DumpStatistics()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim handle As Integer
    handle = FreeFile()
    
    Dim line As String
    Dim i As Long
    Dim j As Long
    
    Open App.Path & "\logs\frags.txt" For Output As handle
    
    'Save lvl vs lvl frag matrix for each class - we use GNU Octave's ASCII file format
    
    Print #handle, "# name: fragLvlLvl_Ase"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(1).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlLvl_Bar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(2).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlLvl_Mag"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(3).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlLvl_Pal"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(4).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlLvl_Gue"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(5).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlLvl_Cle"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(6).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlLvl_Caz"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(7).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    
    
    
    
    'Save lvl vs race frag matrix for each class - we use GNU Octave's ASCII file format
    
    Print #handle, "# name: fragLvlRace_Ase"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(1).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlRace_Bar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(2).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlRace_Mag"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(3).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlRace_Pal"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(4).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlRace_Gue"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(5).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlRace_Cle"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(6).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlRace_Caz"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(7).matrix(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    
    
    
    
    
    'Save lvl vs class frag matrix for each race - we use GNU Octave's ASCII file format
    
    Print #handle, "# name: fragLvlClass_Elf"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 1))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlClass_Dar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 2))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlClass_Dwa"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 3))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlClass_Gno"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 4))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Print #handle, "# name: fragLvlClass_Hum"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 5))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    
    
    
    'Save lvl vs alignment frag matrix for each race - we use GNU Octave's ASCII file format
    
    Print #handle, "# name: fragAlignmentLvl"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 4"
    Print #handle, "# columns: 50"
    
    For j = 1 To 4
        For i = 1 To 50
            line = line & " " & CStr(fragAlignmentLvlData(i, j))
        Next i
        
        Print #handle, line
        line = vbNullString
    Next j
    
    Close handle
    
    
    
    'Dump Chat statistics
    handle = FreeFile()
    
    Open App.Path & "\logs\huffman.log" For Output As handle
    
    Dim Total As Currency
    
    'Compute total characters
    For i = 0 To 255
        Total = Total + keyOcurrencies(i)
    Next i
    
    'Show each character's ocurrencies
    If Total <> 0 Then
        For i = 0 To 255
            Print #handle, CStr(i) & "    " & CStr(Round(keyOcurrencies(i) / Total, 8))
        Next i
    End If
    
    Print #handle, "TOTAL =    " & CStr(Total)
    
    Close handle
End Sub

Public Sub ParseChat(ByRef S As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Long
    Dim key As Integer
    
    For i = 1 To Len(S)
        key = Asc(mid$(S, i, 1))
        
        keyOcurrencies(key) = keyOcurrencies(key) + 1
    Next i
    
    'Add a NULL-terminated to consider that possibility too....
    keyOcurrencies(0) = keyOcurrencies(0) + 1
End Sub
