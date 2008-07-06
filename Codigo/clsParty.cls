VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsParty.cls
'
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

Private p_members(1 To PARTY_MAXMEMBERS) As tPartyMember
'miembros

Private p_expTotal As Long
'Estadistica :D

Private p_Fundador As Integer
'el creador

Private p_CantMiembros As Integer
'cantidad de miembros

Private p_SumaNivelesElevados As Single
'suma de todos los niveles elevados a la ExponenteNivelParty > Esta variable se usa para calcular la experiencia repartida en la Party.

'datos en los pjs: | indexParty(indice en p_members), partyLeader(userindex del lider) |

'Constructor de clase
Public Sub Class_Initialize()
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 p_SumaNiveles added (Tavo)
' - 07/04/08 p_SumaNiveles changed to p_SumaNivelesElevados (MarKoxX)
'***************************************************
    p_expTotal = 0
    p_CantMiembros = 0
    p_SumaNivelesElevados = 0
End Sub

'Destructor de clase
Public Sub Class_Terminate()
End Sub


Public Function MiExperiencia(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
'***************************************************
'Me dice cuanta experiencia tengo colectada ya en la party
Dim i As Integer
i = 0

While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex <> UserIndex
    i = i + 1
Wend

If i <= PARTY_MAXMEMBERS Then
    MiExperiencia = Fix(p_members(i).Experiencia)
Else        'esto no deberia pasar :p
    MiExperiencia = -1
End If

End Function

Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal mapa As Integer, X As Integer, Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 New formula for calculating the experience point of each user
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
' - 04/04/08 Ahora antes de calcular la experiencia a X usuario se fija si ese usuario existe (MarKoxX)
' - 07/04/08 New formula to calculate Experience for each user. (MarKoxX)
'***************************************************
'Se produjo un evento que da experiencia en la wp referenciada
Dim i As Integer
Dim UI As Integer
Dim expThisUser As Double

p_expTotal = p_expTotal + ExpGanada

For i = 1 To PARTY_MAXMEMBERS
    UI = p_members(i).UserIndex
    If UI > 0 Then
        ' Formula: Exp* (Nivel ^ ExponenteNivelParty) / sumadeNivelesElevados
        expThisUser = CDbl(ExpGanada * (UserList(p_members(i).UserIndex).Stats.ELV ^ ExponenteNivelParty) / p_SumaNivelesElevados)
        
        If mapa = UserList(UI).Pos.map And UserList(UI).flags.Muerto = 0 Then
            If Distance(UserList(UI).Pos.X, UserList(UI).Pos.Y, X, Y) <= PARTY_MAXDISTANCIA Then
                p_members(i).Experiencia = p_members(i).Experiencia + expThisUser
                If p_members(i).Experiencia < 0 Then
                    p_members(i).Experiencia = 0
                End If
                If PARTY_EXPERIENCIAPORGOLPE Then
                    UserList(UI).Stats.Exp = UserList(UI).Stats.Exp + Fix(expThisUser)
                    If UserList(UI).Stats.Exp > MAXEXP Then _
                        UserList(UI).Stats.Exp = MAXEXP
                    Call CheckUserLevel(UI)
                    Call WriteUpdateUserStats(UI)
                End If
            End If
        End If
    End If
Next i

End Sub

Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)
'feo feo, muy feo acceder a senddata desde aca, pero BUEEEEEEEEEEE...
Dim i As Integer

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).UserIndex > 0 Then
        Call WriteConsoleMsg(p_members(i).UserIndex, " [" & Sender & "] " & texto, FontTypeNames.FONTTYPE_PARTY)
    End If
Next i

End Sub

Public Function EsPartyLeader(ByVal UserIndex As Integer) As Boolean
    EsPartyLeader = (UserIndex = p_Fundador)
End Function

Public Function NuevoMiembro(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 There is no level prohibition (Tavo)
' - 07/04/08 Added const ExponenteNivelParty. (MarKoxX)
'***************************************************

Dim i As Integer
i = 1
While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex > 0
    i = i + 1
Wend

If i <= PARTY_MAXMEMBERS Then
    p_members(i).Experiencia = 0
    p_members(i).UserIndex = UserIndex
    NuevoMiembro = True
    p_CantMiembros = p_CantMiembros + 1
    p_SumaNivelesElevados = p_SumaNivelesElevados + (UserList(UserIndex).Stats.ELV ^ ExponenteNivelParty)
Else
    NuevoMiembro = False
End If

End Function

Public Function SaleMiembro(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number (Tavo)
' - 07/04/08 Added const ExponenteNivelParty. (MarKoxX)
'***************************************************
'el valor de retorno representa si se disuelve la party
Dim i As Integer
Dim j As Integer
i = 1
SaleMiembro = False
While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex <> UserIndex
    i = i + 1
Wend

If i = 1 Then
    'sale el founder, la party se disuelve
    SaleMiembro = True
    Call MandarMensajeAConsola("El lider disuelve la party.", "Servidor")
    For j = PARTY_MAXMEMBERS To 1 Step -1
        If p_members(j).UserIndex > 0 Then
            
            Call WriteConsoleMsg(p_members(j).UserIndex, " Abandonas la party liderada por " & UserList(p_members(1).UserIndex).name, FontTypeNames.FONTTYPE_PARTY)
    
            Call WriteConsoleMsg(p_members(j).UserIndex, " Durante la misma has conseguido " & CStr(Fix(p_members(j).Experiencia)) & " puntos de experiencia!", FontTypeNames.FONTTYPE_PARTY)
    
            If Not PARTY_EXPERIENCIAPORGOLPE Then
                UserList(p_members(j).UserIndex).Stats.Exp = UserList(p_members(j).UserIndex).Stats.Exp + Fix(p_members(j).Experiencia)
                If UserList(p_members(j).UserIndex).Stats.Exp > MAXEXP Then _
                    UserList(p_members(j).UserIndex).Stats.Exp = MAXEXP
                Call CheckUserLevel(p_members(j).UserIndex)
                Call WriteUpdateUserStats(p_members(j).UserIndex)
            End If
            Call MandarMensajeAConsola(UserList(p_members(j).UserIndex).name & " abandona la party.", "Servidor")
            UserList(p_members(j).UserIndex).PartyIndex = 0
            p_CantMiembros = p_CantMiembros - 1
            p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(UserIndex).Stats.ELV ^ ExponenteNivelParty)
            p_members(j).UserIndex = 0
            p_members(j).Experiencia = 0
        End If
    Next j
Else
    If i <= PARTY_MAXMEMBERS Then
        If Not PARTY_EXPERIENCIAPORGOLPE Then
            UserList(p_members(i).UserIndex).Stats.Exp = UserList(p_members(i).UserIndex).Stats.Exp + Fix(p_members(i).Experiencia)
            If UserList(p_members(i).UserIndex).Stats.Exp > MAXEXP Then _
                UserList(p_members(i).UserIndex).Stats.Exp = MAXEXP
            Call CheckUserLevel(p_members(i).UserIndex)
            Call WriteUpdateUserStats(p_members(i).UserIndex)
        End If
        Call MandarMensajeAConsola(UserList(p_members(i).UserIndex).name & " abandona la party.", "Servidor")
        'TODO: Revisar que esto este bien, y no este faltando/sobrando un mensaje, ahora solo los estoy corrigiendo
        Call WriteConsoleMsg(p_members(i).UserIndex, " Durante la misma has conseguido " & CStr(Fix(p_members(i).Experiencia)) & " puntos de experiencia!", FontTypeNames.FONTTYPE_PARTY)
        p_CantMiembros = p_CantMiembros - 1
        p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(UserIndex).Stats.ELV ^ ExponenteNivelParty)
        p_members(i).UserIndex = 0
        p_members(i).Experiencia = 0
        CompactMemberList
    End If
End If

End Function

Public Function HacerLeader(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 There is no level prohibition
'***************************************************
Dim i As Integer
Dim OldLeader As Integer
Dim oldExp As Double
Dim UserIndexIndex As Integer

UserIndexIndex = 0
HacerLeader = True

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).UserIndex > 0 Then
        If p_members(i).UserIndex = UserIndex Then
            UserIndexIndex = i
        End If
    End If
Next i

If Not HacerLeader Then Exit Function

If UserIndexIndex = 0 Then
    'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
    Call LogError("INCONSISTENCIA DE PARTIES")
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(" Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!", FontTypeNames.FONTTYPE_GUILD))
    HacerLeader = False
    Exit Function
End If


'aca esta todo bien y doy vuelta las collections
OldLeader = p_members(1).UserIndex
oldExp = p_members(1).Experiencia

p_members(1).UserIndex = p_members(UserIndexIndex).UserIndex     'que en realdiad es el userindex, pero no quiero inconsistencias moviendo experiencias
p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia

p_members(UserIndexIndex).UserIndex = OldLeader
p_members(UserIndexIndex).Experiencia = oldExp

p_Fundador = p_members(1).UserIndex

'no need to compact
End Function


Public Sub ObtenerMiembrosOnline(ByRef MemberList As String)
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
'***************************************************
'esta funcion viola un poco la correcta forma de hacer la clase
'pero bue...
Dim i As Integer
MemberList = "Nombre(Exp): "
For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).UserIndex > 0 Then
        MemberList = MemberList & " - " & UserList(p_members(i).UserIndex).name & " (" & CStr(Fix(p_members(i).Experiencia)) & ")"
    End If
Next i

MemberList = MemberList & ". Experiencia total: " & CStr(p_expTotal)

End Sub


Public Function PuedeEntrar(ByVal UserIndex As Integer, ByRef razon As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 There is no level prohibition
'***************************************************
'DEFINE LAS REGLAS DEL JUEGO PARA DEJAR ENTRAR A MIEMBROS
Dim esArmada As Boolean
Dim esCaos As Boolean
Dim MyLevel As Integer
Dim i As Integer
Dim rv As Boolean
Dim UI As Integer

rv = True
esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)
esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
MyLevel = UserList(UserIndex).Stats.ELV

rv = Distancia(UserList(p_members(1).UserIndex).Pos, UserList(UserIndex).Pos) <= MAXDISTANCIAINGRESOPARTY
If rv Then
    rv = (p_members(PARTY_MAXMEMBERS).UserIndex = 0)
    If rv Then
        For i = 1 To PARTY_MAXMEMBERS
            UI = p_members(i).UserIndex
            'pongo los casos que evitarian que pueda entrar
            'aspirante armada en party crimi
            If UI > 0 Then
                If esArmada And criminal(UI) Then
                    razon = "Los miembros de la armada no entran a una party con criminales."
                    rv = False
                End If
                'aspirante caos en party ciuda
                If esCaos And Not criminal(UI) Then
                    razon = "Los miembros de la legión oscura no entran a una party con ciudadanos."
                    rv = False
                End If
                'aspirante crimi en party armada
                If UserList(UI).Faccion.ArmadaReal = 1 And criminal(UserIndex) Then
                    razon = "Los criminales no entran a partys con miembros de la armada real."
                    rv = False
                End If
                'aspirante ciuda en party caos
                If UserList(UI).Faccion.FuerzasCaos = 1 And Not criminal(UserIndex) Then
                    razon = "Los miembros de la legión oscura no entran a una party con ciudadanos."
                    rv = False
                End If
            
                If Not rv Then Exit For 'violate una programacion estructurada
            End If
        Next i
    Else
        razon = "La mayor cantidad de miembros es " & PARTY_MAXMEMBERS
    End If
Else
    razon = "Te encuentras muy lejos del fundador."
End If

PuedeEntrar = rv

End Function


Public Sub FlushExperiencia()
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
'***************************************************
'esta funcion se invoca frente a cerradas del servidor. Flushea la experiencia
'acumulada a los usuarios.

Dim i As Integer
If Not PARTY_EXPERIENCIAPORGOLPE Then   'esto sirve SOLO cuando acumulamos la experiencia!
    For i = 1 To PARTY_MAXMEMBERS
        If p_members(i).UserIndex > 0 Then
            If p_members(i).Experiencia > 0 Then
                UserList(p_members(i).UserIndex).Stats.Exp = UserList(p_members(i).UserIndex).Stats.Exp + Fix(p_members(i).Experiencia)
                If UserList(p_members(i).UserIndex).Stats.Exp > MAXEXP Then _
                    UserList(p_members(i).UserIndex).Stats.Exp = MAXEXP
                Call CheckUserLevel(p_members(i).UserIndex)
            Else
                If Abs(UserList(p_members(i).UserIndex).Stats.Exp) > Abs(Fix(p_members(i).Experiencia)) Then
                    UserList(p_members(i).UserIndex).Stats.Exp = UserList(p_members(i).UserIndex).Stats.Exp + Fix(p_members(i).Experiencia)
                Else
                    UserList(p_members(i).UserIndex).Stats.Exp = 0
                End If
            End If
            p_members(i).Experiencia = 0
            Call WriteUpdateUserStats(p_members(i).UserIndex)
        End If
    Next i
End If

End Sub




Private Sub CompactMemberList()
Dim i As Integer
Dim freeIndex As Integer
i = 1
While i <= PARTY_MAXMEMBERS
    If p_members(i).UserIndex = 0 And freeIndex = 0 Then
        freeIndex = i
    ElseIf p_members(i).UserIndex > 0 And freeIndex > 0 Then
        p_members(freeIndex).Experiencia = p_members(i).Experiencia
        p_members(freeIndex).UserIndex = p_members(i).UserIndex
        p_members(i).UserIndex = 0
        p_members(i).Experiencia = 0
        'muevo el de la pos i a freeindex
        i = freeIndex
        freeIndex = 0
    End If
    i = i + 1
Wend

End Sub

Public Function CantMiembros() As Integer
CantMiembros = p_CantMiembros
End Function
