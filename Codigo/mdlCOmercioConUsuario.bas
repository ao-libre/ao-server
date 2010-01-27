Attribute VB_Name = "mdlCOmercioConUsuario"
'**************************************************************
' mdlComercioConUsuarios.bas - Allows players to commerce between themselves.
'
' Designed and implemented by Alejandro Santos (AlejoLP)
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

'[Alejo]
Option Explicit

Private Const MAX_ORO_LOGUEABLE As Long = 50000
Private Const MAX_OBJ_LOGUEABLE As Long = 1000

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    
    'El tipo de datos de Cant ahora es Long (antes Integer)
    'asi se puede comerciar con oro > 32k
    cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    Acepto As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
On Error GoTo Errhandler

'Si ambos pusieron /comerciar entonces
If UserList(Origen).ComUsu.DestUsu = Destino And _
   UserList(Destino).ComUsu.DestUsu = Origen Then
    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Origen, 0)
    'Decirle al origen que abra la ventanita.
    Call WriteUserCommerceInit(Origen)
    UserList(Origen).flags.Comerciando = True

    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Destino, 0)
    'Decirle al origen que abra la ventanita.
    Call WriteUserCommerceInit(Destino)
    UserList(Destino).flags.Comerciando = True

    'Call EnviarObjetoTransaccion(Origen)
Else
    'Es el primero que comercia ?
    Call WriteConsoleMsg(Destino, UserList(Origen).name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
    UserList(Destino).flags.TargetUser = Origen
    
End If

Call FlushBuffer(Destino)

Exit Sub
Errhandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.description)
End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer)
Dim ObjInd As Integer
Dim ObjCant As Long

'[Alejo]: En esta funcion se centralizaba el problema
'         de no poder comerciar con mas de 32k de oro.
'         Ahora si funciona!!!

ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.cant
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
    ObjInd = iORO
Else
    ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
End If

If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub

If ObjInd > 0 And ObjCant > 0 Then
    Call WriteChangeUserTradeSlot(AQuien, ObjInd, ObjCant)
    Call FlushBuffer(AQuien)
End If

End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .ComUsu.DestUsu > 0 Then
            Call WriteUserCommerceEnd(UserIndex)
        End If
        
        .ComUsu.Acepto = False
        .ComUsu.cant = 0
        .ComUsu.DestUsu = 0
        .ComUsu.Objeto = 0
        .ComUsu.DestNick = vbNullString
        .flags.Comerciando = False
    End With
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)

On Error GoTo Errhandler

    Dim Obj1 As Obj, Obj2 As Obj
    Dim OtroUserIndex As Integer
    Dim TerminarAhora As Boolean
    
    TerminarAhora = False
    
    If UserList(UserIndex).ComUsu.DestUsu <= 0 Or UserList(UserIndex).ComUsu.DestUsu > MaxUsers Then
        TerminarAhora = True
    End If
    
    OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
    
    If Not TerminarAhora Then
        If UserList(OtroUserIndex).flags.UserLogged = False Or UserList(UserIndex).flags.UserLogged = False Then
            TerminarAhora = True
        End If
    End If
    
    If Not TerminarAhora Then
        If UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex Then
            TerminarAhora = True
        End If
    End If
    
    If Not TerminarAhora Then
        If UserList(OtroUserIndex).name <> UserList(UserIndex).ComUsu.DestNick Then
            TerminarAhora = True
        End If
    End If
    
    If Not TerminarAhora Then
        If UserList(UserIndex).name <> UserList(OtroUserIndex).ComUsu.DestNick Then
            TerminarAhora = True
        End If
    End If
    
    If TerminarAhora = True Then
        Call FinComerciarUsu(UserIndex)
        
        If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
            Call FinComerciarUsu(OtroUserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
        
        Exit Sub
    End If
    
    UserList(UserIndex).ComUsu.Acepto = True
    TerminarAhora = False
    
    If UserList(OtroUserIndex).ComUsu.Acepto = False Then
        Call WriteConsoleMsg(UserIndex, "El otro usuario aun no ha aceptado tu oferta.", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
    
    If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
        Obj1.ObjIndex = iORO
        If UserList(UserIndex).ComUsu.cant > UserList(UserIndex).Stats.GLD Then
            Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
            TerminarAhora = True
        End If
    Else
        Obj1.amount = UserList(UserIndex).ComUsu.cant
        Obj1.ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).ObjIndex
        If Obj1.amount > UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).amount Then
            Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
            TerminarAhora = True
        End If
    End If
    
    If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
        Obj2.ObjIndex = iORO
        If UserList(OtroUserIndex).ComUsu.cant > UserList(OtroUserIndex).Stats.GLD Then
            Call WriteConsoleMsg(OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
            TerminarAhora = True
        End If
    Else
        Obj2.amount = UserList(OtroUserIndex).ComUsu.cant
        Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
        If Obj2.amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).amount Then
            Call WriteConsoleMsg(OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
            TerminarAhora = True
        End If
    End If
    
    'Por si las moscas...
    If TerminarAhora = True Then
        Call FinComerciarUsu(UserIndex)
        
        Call FinComerciarUsu(OtroUserIndex)
        Call FlushBuffer(OtroUserIndex)
        Exit Sub
    End If
    
    Call FlushBuffer(OtroUserIndex)
    
    'Desde acá corregí el bug que cuando se ofrecian mas de
    '10k de oro no le llegaban al destinatario.
    
    'pone el oro directamente en la billetera
    If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
        'quito la cantidad de oro ofrecida
        UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.cant
        If UserList(OtroUserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(OtroUserIndex).name & " solto oro en comercio seguro con " & UserList(UserIndex).name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.cant)
        Call WriteUpdateUserStats(OtroUserIndex)
        'y se la doy al otro
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.cant
        If UserList(OtroUserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(UserIndex).name & " recibio oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.cant)
        'Esta linea del log es al pedo. > Vuelvo a ponerla a pedido del CGMS
        Call WriteUpdateUserStats(UserIndex)
    Else
        'Quita el objeto y se lo da al otro
        If Not MeterItemEnInventario(UserIndex, Obj2) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj2)
        End If
    
        Call QuitarObjetos(Obj2.ObjIndex, Obj2.amount, OtroUserIndex)
        
        'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
        If ObjData(Obj2.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(OtroUserIndex).name & " le pasó en comercio seguro a " & UserList(UserIndex).name & " " & Obj2.amount & " " & ObjData(Obj2.ObjIndex).name)
        End If
    
        'Es mucha cantidad?
        If Obj2.amount > MAX_OBJ_LOGUEABLE Then
        'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Obj2.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(OtroUserIndex).name & " le pasó en comercio seguro a " & UserList(UserIndex).name & " " & Obj2.amount & " " & ObjData(Obj2.ObjIndex).name)
            End If
        End If
    End If
    
    'pone el oro directamente en la billetera
    If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
        'quito la cantidad de oro ofrecida
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.cant
        If UserList(UserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(UserIndex).name & " soltó oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & UserList(UserIndex).ComUsu.cant)
        Call WriteUpdateUserStats(UserIndex)
        'y se la doy al otro
        UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.cant
        'If UserList(UserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(OtroUserIndex).name & " recibio oro en comercio seguro con " & UserList(UserIndex).name & ". Cantidad: " & UserList(UserIndex).ComUsu.cant)
        'Esta linea del log es al pedo.
        Call WriteUpdateUserStats(OtroUserIndex)
    Else
        'Quita el objeto y se lo da al otro
        If Not MeterItemEnInventario(OtroUserIndex, Obj1) Then
            Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, Obj1)
        End If
    
        Call QuitarObjetos(Obj1.ObjIndex, Obj1.amount, UserIndex)
        
        'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
        If ObjData(Obj1.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).name & " le pasó en comercio seguro a " & UserList(OtroUserIndex).name & " " & Obj1.amount & " " & ObjData(Obj1.ObjIndex).name)
        End If
    
        'Es mucha cantidad?
        If Obj1.amount > MAX_OBJ_LOGUEABLE Then
        'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Obj1.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(OtroUserIndex).name & " le pasó en comercio seguro a " & UserList(UserIndex).name & " " & Obj1.amount & " " & ObjData(Obj1.ObjIndex).name)
            End If
        End If
    End If
    
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    
    Exit Sub

Errhandler:
    Dim UserNick As String
    Dim OtroUSerNick As String

    If UserIndex > 0 Then UserNick = UserList(UserIndex).name
    If OtroUserIndex > 0 Then OtroUSerNick = UserList(OtroUserIndex).name
    
    Call LogError("Error en AceptarComercioUsu. Error " & Err.Number & " : " & Err.description & " Userindex: " & UserIndex & _
       " Nick: " & UserNick & " OtroUserIndex: " & OtroUserIndex & " Nick: " & OtroUSerNick)
End Sub
'[/Alejo]
