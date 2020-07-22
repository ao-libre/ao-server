VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWorldPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mX As Integer
Private mY As Integer
Private mMap As Integer
Public Property Get X() As Integer
    X = mX
End Property

Public Property Let X(posX As Integer)
    mX = posX
End Property

Public Property Get Y() As Integer
    Y = mY
End Property

Public Property Let Y(posY As Integer)
    mY = posY
End Property
Public Property Get Map() As Integer
    Map = mMap
End Property

Public Property Let Map(inMap As Integer)
    mMap = inMap
End Property
