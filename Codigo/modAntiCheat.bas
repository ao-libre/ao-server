Attribute VB_Name = "modAntiCheat"
Option Explicit
 
Public Type TimeIntervalos
 
    UsarItem As Byte
    AtacaArco As Byte
    AtacaComun As Byte
    CastSpell As Byte
 
End Type
 
Public Sub ResetAllCount(ByVal Userindex As Integer)
 
    With UserList(Userindex)
        
        If (.Counters.Cheat.AtacaArco <> 0) Then
            .Counters.Cheat.AtacaArco = 0
        End If
 
        If (.Counters.Cheat.AtacaComun <> 0) Then
            .Counters.Cheat.AtacaComun = 0
        End If
 
        If (.Counters.Cheat.CastSpell <> 0) Then
            .Counters.Cheat.CastSpell = 0
        End If
 
        If (.Counters.Cheat.UsarItem <> 0) Then
            .Counters.Cheat.UsarItem = 0
        End If
                
    End With
 
End Sub
 
Public Sub RestaCount(ByVal Userindex As Integer, _
                      Optional ByVal Flecha As Byte = 0, _
                      Optional ByVal Golpe As Byte = 0, _
                      Optional ByVal Cast As Byte = 0, _
                      Optional ByVal Usar As Byte = 0)

    With UserList(Userindex)
 
        If (Flecha <> 0) Then
            .Counters.Cheat.AtacaArco = 0
        End If
 
        If (Golpe <> 0) Then
            .Counters.Cheat.AtacaComun = 0
        End If
 
        If (Cast <> 0) Then
            .Counters.Cheat.CastSpell = 0
        End If
 
        If (Usar <> 0) Then
            .Counters.Cheat.UsarItem = 0
        End If
 
    End With
 
End Sub
 
Public Sub AddCount(ByVal Userindex As Integer, _
                    Optional ByVal AddFlecha As Byte = 0, _
                    Optional ByVal AddGolpe As Byte = 0, _
                    Optional ByVal AddCast As Byte = 0, _
                    Optional ByVal AddUsar As Byte = 0)
 
    Dim Msj As String
 
    With UserList(Userindex)
 
        If (AddFlecha <> 0) Then
            .Counters.Cheat.AtacaArco = (.Counters.Cheat.AtacaArco + 1)
 
            If CheckInt(Userindex, Msj, 1) Then
                Call MsjCheat(Userindex, Msj)
            End If
                        
        End If
 
        If (AddGolpe <> 0) Then
            .Counters.Cheat.AtacaComun = (.Counters.Cheat.AtacaComun + 1)
 
            If CheckInt(Userindex, Msj, 2) Then
                Call MsjCheat(Userindex, Msj)
            End If
                        
        End If
       
        If (AddCast <> 0) Then
            .Counters.Cheat.CastSpell = (.Counters.Cheat.CastSpell + 1)
 
            If CheckInt(Userindex, Msj, 3) Then
                Call MsjCheat(Userindex, Msj)
            End If
                        
        End If
 
        If (AddUsar <> 0) Then
            .Counters.Cheat.UsarItem = (.Counters.Cheat.UsarItem + 1)
 
            If CheckInt(Userindex, Msj, 4) Then
                Call MsjCheat(Userindex, Msj)
            End If
                        
        End If
                
    End With
        
End Sub
 
Private Function CheckInt(ByVal Userindex As Integer, _
                          ByRef Msj As String, _
                          ByVal Intervalo As Byte) As Boolean
 
    Const MaxTol As Byte = 3
 
    With UserList(Userindex)
 
        Select Case Intervalo
        
            Case 1
   
                If (.Counters.Cheat.AtacaArco = MaxTol) Then
                    Msj = ". -" & "Sobrepaso el intervalo de Ataca Arco 3 veces seguidas." & vbNewLine & "Posible edicion de intervalos."
                    .Counters.Cheat.AtacaArco = 0
                    CheckInt = True
 
                    Exit Function
 
                End If
 
            Case 2
 
                If (.Counters.Cheat.AtacaComun = MaxTol) Then
                    Msj = ". -" & "Sobrepaso el intervalo de Ataca Comun 3 veces seguidas." & vbNewLine & "Posible edicion de intervalos."
                    .Counters.Cheat.AtacaComun = 0
                    CheckInt = True
  
                    Exit Function
 
                End If
 
            Case 3
 
                If (.Counters.Cheat.CastSpell = MaxTol) Then
                    Msj = ". -" & "Sobrepaso el intervalo de Cast Spell 3 veces seguidas." & vbNewLine & "Posible edicion de intervalos."
                    .Counters.Cheat.CastSpell = 0
                    CheckInt = True
 
                    Exit Function
 
                End If
 
            Case 4
 
               If (.Counters.Trabajando = 1) Then Exit Function

                If (.Counters.Cheat.UsarItem = MaxTol) Then
                    Msj = ". -" & "Sobrepaso el intervalo de Usar Items 3 veces seguidas." & vbNewLine & "Posible edicion de intervalos."
                    .Counters.Cheat.UsarItem = 0
                    CheckInt = True
 
                    Exit Function

                End If
     
        End Select
        
    End With
 
    CheckInt = False
        
End Function
 
Private Sub MsjCheat(ByVal Userindex As Integer, ByVal Msj As String)
 
    Dim sndData As String
 
    With UserList(Userindex)
        
        sndData = PrepareMessageConsoleMsg(.Name & Msj, FontTypeNames.FONTTYPE_SERVER)
                
        Call SendData(SendTarget.ToAdmins, 0, sndData)
            
        Call LogIntervalos(.Name, Msj)
                                
    End With
 
End Sub
 
Private Sub LogIntervalos(ByVal Nombre As String, ByVal Str As String)
 
    On Error GoTo errHandler
 
    Dim nfile As Integer
 
    nfile = FreeFile
        
    Open App.Path & "\AntiCheats\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date$ & " " & time$ & " " & Str
    Close #nfile
    
    Exit Sub
 
errHandler:
 
End Sub

