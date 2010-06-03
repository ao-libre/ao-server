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

Public Const MAX_OFFER_SLOTS As Integer = 20
Public Const GOLD_OFFER_SLOT As Integer = MAX_OFFER_SLOTS + 1

Public Type tComercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto(1 To MAX_OFFER_SLOTS) As Integer 'Indice de los objetos que se desea dar
    GoldAmount As Long
    
    cant(1 To MAX_OFFER_SLOTS) As Long 'Cuantos objetos desea dar
    Acepto As Boolean
    Confirmo As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
'***************************************************
'Autor: Unkown
'Last Modification: 25/11/2009
'
'***************************************************
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
        Call WriteConsoleMsg(Destino, UserList(Origen).name & " desea comerciar. Si deseas aceptar, escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
        UserList(Destino).flags.TargetUser = Origen
        
    End If
    
    Call FlushBuffer(Destino)
    
    Exit Sub
Errhandler:
        Call LogError("Error en IniciarComercioConUsuario: " & Err.description)
End Sub

Public Sub EnviarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte)
'***************************************************
'Autor: Unkown
'Last Modification: 25/11/2009
'Sends the offer change to the other trading user
'25/11/2009: ZaMa - Implementado nuevo sistema de comercio con ofertas variables.
'***************************************************
    Dim ObjIndex As Integer
    Dim ObjAmount As Long
    
    With UserList(UserIndex)
        If OfferSlot = GOLD_OFFER_SLOT Then
            ObjIndex = iORO
            ObjAmount = UserList(.ComUsu.DestUsu).ComUsu.GoldAmount
        Else
            ObjIndex = UserList(.ComUsu.DestUsu).ComUsu.Objeto(OfferSlot)
            ObjAmount = UserList(.ComUsu.DestUsu).ComUsu.cant(OfferSlot)
        End If
    End With
   
    Call WriteChangeUserTradeSlot(UserIndex, OfferSlot, ObjIndex, ObjAmount)
    Call FlushBuffer(UserIndex)

End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unkown
'Last Modification: 25/11/2009
'25/11/2009: ZaMa - Limpio los arrays (por el nuevo sistema)
'***************************************************
    Dim i As Long
    
    With UserList(UserIndex)
        If .ComUsu.DestUsu > 0 Then
            Call WriteUserCommerceEnd(UserIndex)
        End If
        
        .ComUsu.Acepto = False
        .ComUsu.Confirmo = False
        .ComUsu.DestUsu = 0
        
        For i = 1 To MAX_OFFER_SLOTS
            .ComUsu.cant(i) = 0
            .ComUsu.Objeto(i) = 0
        Next i
        
        .ComUsu.GoldAmount = 0
        .ComUsu.DestNick = vbNullString
        .flags.Comerciando = False
    End With
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unkown
'Last Modification: 25/11/2009
'25/11/2009: ZaMa - Ahora se traspasan hasta 5 items + oro al comerciar
'***************************************************
    Dim TradingObj As Obj
    Dim OtroUserIndex As Integer
    Dim TerminarAhora As Boolean
    Dim OfferSlot As Integer
    Dim invBackUp() As UserOBJ
    Dim invBackUp2() As UserOBJ
    Dim gldBackUp As Long
    Dim gldBackUp2 As Long

    UserList(UserIndex).ComUsu.Acepto = True
    
    OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
    
    If UserList(OtroUserIndex).ComUsu.Acepto = False Then
        Exit Sub
    End If
    
    If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    End If
    
    invBackUp = UserList(UserIndex).Invent.Object
    invBackUp2 = UserList(OtroUserIndex).Invent.Object
    gldBackUp = UserList(UserIndex).Stats.GLD
    gldBackUp2 = UserList(OtroUserIndex).Stats.GLD
    
    ' Envio los items a quien corresponde
    For OfferSlot = 1 To MAX_OFFER_SLOTS + 1
        
        ' Items del 1er usuario
        With UserList(UserIndex)
            ' Le pasa el oro
            If OfferSlot = GOLD_OFFER_SLOT Then
                If .ComUsu.GoldAmount > .Stats.GLD Then
                    Call LogHackAttemp(.name & " IP:" & .ip & " intentó comerciar " & .ComUsu.GoldAmount & " y tenía " & .Stats.GLD)
                    
                    .Invent.Object = invBackUp
                    .Stats.GLD = gldBackUp
                    UserList(OtroUserIndex).Invent.Object = invBackUp2
                    UserList(OtroUserIndex).Stats.GLD = gldBackUp2
    
                    Call WriteConsoleMsg(UserIndex, "Comercio terminado.", FontTypeNames.FONTTYPE_TALK)
                    Call WriteConsoleMsg(OtroUserIndex, "Comercio terminado.", FontTypeNames.FONTTYPE_TALK)
                    
                    Call FinComerciarUsu(UserIndex)
                    
                    Call FinComerciarUsu(OtroUserIndex)
                    Call Protocol.FlushBuffer(OtroUserIndex)
                    
                    Exit Sub
                End If
                
                ' Quito la cantidad de oro ofrecida
                .Stats.GLD = .Stats.GLD - .ComUsu.GoldAmount
                ' Log
                If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(.name & " soltó oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & .ComUsu.GoldAmount)
                ' Update Usuario
                Call WriteUpdateUserStats(UserIndex)
                ' Se la doy al otro
                UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + .ComUsu.GoldAmount
                ' Update Otro Usuario
                Call WriteUpdateUserStats(OtroUserIndex)
                
            ' Le pasa lo ofertado de los slots con items
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.cant(OfferSlot)
                
                If Not TieneObjetos(TradingObj.ObjIndex, TradingObj.Amount, UserIndex) Then
                    Call LogHackAttemp(.name & " IP:" & .ip & " intentó comerciar una cantidad de objetos que no tenía.")
                    
                    .Invent.Object = invBackUp
                    .Stats.GLD = gldBackUp
                    UserList(OtroUserIndex).Invent.Object = invBackUp2
                    UserList(OtroUserIndex).Stats.GLD = gldBackUp2
                    
                    Call WriteConsoleMsg(UserIndex, "Comercio terminado.", FontTypeNames.FONTTYPE_TALK)
                    Call WriteConsoleMsg(OtroUserIndex, "Comercio terminado.", FontTypeNames.FONTTYPE_TALK)
                    
                    Call FinComerciarUsu(UserIndex)
                    
                    Call FinComerciarUsu(OtroUserIndex)
                    Call Protocol.FlushBuffer(OtroUserIndex)
                    
                    Exit Sub
                End If
                
                'Quita el objeto y se lo da al otro
                If Not MeterItemEnInventario(OtroUserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, TradingObj)
                End If
            
                Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, UserIndex)
                
                'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
                If ObjData(TradingObj.ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.name & " le pasó en comercio seguro a " & UserList(OtroUserIndex).name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).name)
                End If
            
                'Es mucha cantidad?
                If TradingObj.Amount > MAX_OBJ_LOGUEABLE Then
                'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(UserList(OtroUserIndex).name & " le pasó en comercio seguro a " & .name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).name)
                    End If
                End If
            End If
        End With
        
        ' Items del 2do usuario
        With UserList(OtroUserIndex)
            ' Le pasa el oro
            If OfferSlot = GOLD_OFFER_SLOT Then
                If .ComUsu.GoldAmount > .Stats.GLD Then
                    Call LogHackAttemp(.name & " IP:" & .ip & " intentó comerciar " & .ComUsu.GoldAmount & " y tenía " & .Stats.GLD)
                    
                    UserList(UserIndex).Invent.Object = invBackUp
                    UserList(UserIndex).Stats.GLD = gldBackUp
                    .Invent.Object = invBackUp2
                    .Stats.GLD = gldBackUp2
                    
                    Call WriteConsoleMsg(UserIndex, "Comercio terminado.", FontTypeNames.FONTTYPE_TALK)
                    Call WriteConsoleMsg(OtroUserIndex, "Comercio terminado.", FontTypeNames.FONTTYPE_TALK)
                    
                    Call FinComerciarUsu(OtroUserIndex)
                    Call Protocol.FlushBuffer(OtroUserIndex)
                    
                    Call FinComerciarUsu(UserIndex)
                    
                    Exit Sub
                End If
                
                ' Quito la cantidad de oro ofrecida
                .Stats.GLD = .Stats.GLD - .ComUsu.GoldAmount
                ' Log
                If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(.name & " soltó oro en comercio seguro con " & UserList(UserIndex).name & ". Cantidad: " & .ComUsu.GoldAmount)
                ' Update Usuario
                Call WriteUpdateUserStats(OtroUserIndex)
                'y se la doy al otro
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + .ComUsu.GoldAmount
                If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(UserIndex).name & " recibió oro en comercio seguro con " & .name & ". Cantidad: " & .ComUsu.GoldAmount)
                ' Update Otro Usuario
                Call WriteUpdateUserStats(UserIndex)
                
            ' Le pasa la oferta de los slots con items
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.cant(OfferSlot)
                
                If Not TieneObjetos(TradingObj.ObjIndex, TradingObj.Amount, UserIndex) Then
                    Call LogHackAttemp(.name & " IP:" & .ip & " intentó comerciar una cantidad de objetos que no tenía.")
                    
                    UserList(UserIndex).Invent.Object = invBackUp
                    UserList(UserIndex).Stats.GLD = gldBackUp
                    .Invent.Object = invBackUp2
                    .Stats.GLD = gldBackUp2
                    
                    Call WriteConsoleMsg(UserIndex, "Comercio terminado.", FontTypeNames.FONTTYPE_TALK)
                    Call WriteConsoleMsg(OtroUserIndex, "Comercio terminado.", FontTypeNames.FONTTYPE_TALK)
                    
                    Call FinComerciarUsu(OtroUserIndex)
                    Call Protocol.FlushBuffer(OtroUserIndex)
                    
                    Call FinComerciarUsu(UserIndex)
                    
                    Exit Sub
                End If
                
                'Quita el objeto y se lo da al otro
                If Not MeterItemEnInventario(UserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, TradingObj)
                End If
            
                Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, OtroUserIndex)
                
                'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
                If ObjData(TradingObj.ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.name & " le pasó en comercio seguro a " & UserList(UserIndex).name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).name)
                End If
            
                'Es mucha cantidad?
                If TradingObj.Amount > MAX_OBJ_LOGUEABLE Then
                'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.name & " le pasó en comercio seguro a " & UserList(UserIndex).name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).name)
                    End If
                End If
            End If
        End With
        
    Next OfferSlot

    ' End Trade
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Call Protocol.FlushBuffer(OtroUserIndex)
    
End Sub

Public Sub AgregarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long, ByVal IsGold As Boolean)
'***************************************************
'Autor: ZaMa
'Last Modification: 24/11/2009
'Adds gold or items to the user's offer
'***************************************************

    If PuedeSeguirComerciando(UserIndex) Then
        With UserList(UserIndex).ComUsu
            ' Si ya confirmo su oferta, no puede cambiarla!
            If Not .Confirmo Then
                If IsGold Then
                ' Agregamos (o quitamos) mas oro a la oferta
                    .GoldAmount = .GoldAmount + Amount
                    
                    ' Imposible que pase, pero por las dudas..
                    If .GoldAmount < 0 Then .GoldAmount = 0
                Else
                ' Agreamos (o quitamos) el item y su cantidad en el slot correspondiente
                    ' Si es 0 estoy modificando la cantidad, no agregando
                    If ObjIndex > 0 Then .Objeto(OfferSlot) = ObjIndex
                    .cant(OfferSlot) = .cant(OfferSlot) + Amount
                    
                    'Quitó todos los items de ese tipo
                    If .cant(OfferSlot) <= 0 Then
                        ' Removemos el objeto para evitar conflictos
                        .Objeto(OfferSlot) = 0
                        .cant(OfferSlot) = 0
                    End If
                End If
            End If
        End With
    End If

End Sub

Public Function PuedeSeguirComerciando(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 24/11/2009
'Validates wether the conditions for the commerce to keep going are satisfied
'***************************************************
Dim OtroUserIndex As Integer
Dim ComercioInvalido As Boolean

With UserList(UserIndex)
    ' Usuario valido?
    If .ComUsu.DestUsu <= 0 Or .ComUsu.DestUsu > MaxUsers Then
        ComercioInvalido = True
    End If
    
    OtroUserIndex = .ComUsu.DestUsu
    
    If Not ComercioInvalido Then
        ' Estan logueados?
        If UserList(OtroUserIndex).flags.UserLogged = False Or .flags.UserLogged = False Then
            ComercioInvalido = True
        End If
    End If
    
    If Not ComercioInvalido Then
        ' Se estan comerciando el uno al otro?
        If UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex Then
            ComercioInvalido = True
        End If
    End If
    
    If Not ComercioInvalido Then
        ' El nombre del otro es el mismo que al que le comercio?
        If UserList(OtroUserIndex).name <> .ComUsu.DestNick Then
            ComercioInvalido = True
        End If
    End If
    
    If Not ComercioInvalido Then
        ' Mi nombre  es el mismo que al que el le comercia?
        If .name <> UserList(OtroUserIndex).ComUsu.DestNick Then
            ComercioInvalido = True
        End If
    End If
    
    If Not ComercioInvalido Then
        ' Esta vivo?
        If UserList(OtroUserIndex).flags.Muerto = 1 Then
            ComercioInvalido = True
        End If
    End If
    
    ' Fin del comercio
    If ComercioInvalido = True Then
        Call FinComerciarUsu(UserIndex)
        
        If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
            Call FinComerciarUsu(OtroUserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
        
        Exit Function
    End If
End With

PuedeSeguirComerciando = True

End Function
