Attribute VB_Name = "Statistics"
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

Public Sub Initialize()
    ReDim trainningInfo(1 To MaxUsers) As trainningData
End Sub

Public Sub UserConnected(ByVal userindex As Integer)
    'A new user connected, load it's trainning time count
    trainningInfo(userindex).trainningTime = CLng(GetVar(CharPath & UCase$(UserList(userindex).name) & ".chr", "RESEARCH", "TrainningTime", 10))
    
    trainningInfo(userindex).startTick = GetTickCount
    
'TODO : Get rid of this after reset!!!
    If trainningInfo(userindex).trainningTime = 0 Then _
        trainningInfo(userindex).trainningTime = -1
End Sub

Public Sub UserDisconnected(ByVal userindex As Integer)
'TODO : Get rid of this after reset!!!
    'Abort if char had already started trainning before this system was coded
    If trainningInfo(userindex).trainningTime = -1 Then Exit Sub
    
    With trainningInfo(userindex)
        'Update trainning time
        .trainningTime = .trainningTime + (GetTickCount() - .startTick) / 1000
    
        .startTick = GetTickCount
        
        'Store info in char file
        Call WriteVar(CharPath & UCase$(UserList(userindex).name) & ".chr", "RESEARCH", "TrainningTime", CStr(.trainningTime))
    End With
End Sub

Public Sub UserLevelUp(ByVal userindex As Integer)
    Dim handle As Integer
    handle = FreeFile()
    
    With trainningInfo(userindex)
    
'TODO : get rid of the If after reset!!!
        If .trainningTime <> -1 Then
            'Log the data
            Open App.Path & "\logs\statistics.log" For Append Shared As handle
            
            Print #handle, UCase$(UserList(userindex).name) & " completó el nivel " & CStr(UserList(userindex).Stats.ELV) & " en " & CStr(.trainningTime + (GetTickCount() - .startTick) / 1000) & " segundos."
            
            Close handle
        End If
        
        'Reset data
        .trainningTime = 0
        .startTick = GetTickCount()
    End With
End Sub

Public Sub StoreFrag(ByVal killer As Integer, ByVal victim As Integer)
    Dim clase As Integer
    Dim raza As Integer
    Dim alignment As Integer
    
    Select Case UCase$(UserList(killer).clase)
        Case "ASESINO"
            clase = 1
        
        Case "DRUIDA"
            clase = 2
        
        Case "MAGO"
            clase = 3
        
        Case "PALADIN"
            clase = 4
        
        Case "GUERRERO"
            clase = 5
        
        Case "CLERIGO"
            clase = 6
        
        Case "CAZADOR"
            clase = 7
        
        Case Else
            Exit Sub
    End Select
    
    Select Case UCase$(UserList(killer).raza)
        Case "ELFO"
            raza = 1
        
        Case "ELFO OSCURO"
            raza = 2
        
        Case "ENANO"
            raza = 3
        
        Case "GNOMO"
            raza = 4
        
        Case "HUMANO"
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
    Dim handle As Integer
    handle = FreeFile()
    
    Dim line As String
    Dim i As Long
    Dim j As Long
    
    Open App.Path & "\logs\frags" For Output As handle
    
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
        line = ""
    Next j
    
    Print #handle, "# name: fragLvlLvl_Dru"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(2).matrix(i, j))
        Next i
        
        Print #handle, line
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
    Next j
    
    Print #handle, "# name: fragLvlRace_Dru"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(2).matrix(i, j))
        Next i
        
        Print #handle, line
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
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
        line = ""
    Next j
    
    Close #handle
End Sub
