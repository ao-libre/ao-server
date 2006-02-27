Attribute VB_Name = "mdlLeeMapas"

Option Explicit

'unsigned int DLLIMPORT MAPCargaMapa (const char *archmap, const char *archinf);
'unsigned int DLLIMPORT MAPCierraMapa(unsigned int dm);
'
'unsigned int DLLIMPORT MAPLeeMapa(unsigned int dm, BLOQUE *tile_map, BLOQUE_INF *tile_inf );
'

Public Type TileMap
    bloqueado As Byte
    grafs(1 To 4) As Integer
    trigger As Integer

    t1 As Integer 'espacio al pedo
End Type

Public Type TileInf
    dest_mapa As Integer
    dest_x As Integer
    dest_y As Integer
    
    npc As Integer
    
    obj_ind As Integer
    obj_cant As Integer
    
    t1 As Integer
    t2 As Integer
End Type

'Public Declare Function MAPCargaMapa Lib "LeeMapas.dll" (ByVal archmap As String, ByVal archinf As String) As Long
'Public Declare Function MAPCierraMapa Lib "LeeMapas.dll" (ByVal Dm As Long) As Long
'
'Public Declare Function MAPLeeMapa Lib "LeeMapas.dll" (ByVal Dm As Long, Tile_Map As TileMap, Tile_Inf As TileInf) As Long

