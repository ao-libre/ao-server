Attribute VB_Name = "SysTray"
'Argentum Online 0.12.2
'Copyright (C) 2002 Marquez Pablo Ignacio
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez
Option Explicit
'????????????????????????????
'????????????????????????????
'????????????????????????????
'                       SysTray
'????????????????????????????
'????????????????????????????
'????????????????????????????
'Para minimizar a la barra de tareas
'????????????????????????????
'????????????????????????????
'????????????????????????????

Type CWPSTRUCT

    lParam As Long
    wParam As Long
    message As Long
    hWnd As Long

End Type

Declare Function CallNextHookEx _
        Lib "user32" (ByVal hHook As Long, _
                      ByVal ncode As Long, _
                      ByVal wParam As Long, _
                      lParam As Any) As Long
Declare Sub CopyMemory _
        Lib "kernel32" _
        Alias "RtlMoveMemory" (hpvDest As Any, _
                               hpvSource As Any, _
                               ByVal cbCopy As Long)
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowsHookEx _
        Lib "user32" _
        Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                   ByVal lpfn As Long, _
                                   ByVal hmod As Long, _
                                   ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const WH_CALLWNDPROC = 4

Public Const WM_CREATE = &H1

Public hHook As Long

Public Function AppHook(ByVal idHook As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim CWP As CWPSTRUCT

    CopyMemory CWP, ByVal lParam, Len(CWP)

    Select Case CWP.message

        Case WM_CREATE
            SetForegroundWindow CWP.hWnd
            AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
            UnhookWindowsHookEx hHook
            hHook = 0
            Exit Function

    End Select

    AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)

End Function

