Attribute VB_Name = "AStar"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Type MapTile
    Character As Integer
    layer As Integer
    layer2 As Integer
End Type

Public Type TileVertex
    Brightness As Single
    Explored As Boolean
End Type

Public Type MAPFILEHDR                  'Map file header
    EngName As String * 16              'Map engine name.
    EngVersion As Integer               'Map engine version.
    MapName As String * 16              'Map short name.
    MapDesc As String * 198             'Map long description.
    TileSet As String * 16              'Name of tileset used for map.
    MapWidth As Long                    'Width of the map.
    MapHeight As Long                   'Height of the map.
    CharacterNumber As Integer
End Type

Public Type ChrStandart
    Health As Integer
    Armor As Integer
    SightRange As Integer
    Speed As Integer
    AttackRange As Integer
    AttackRate As Integer
    AttackDamage As Integer
    Mana As Integer
    Buttons(8) As Integer
    Movable() As Boolean
    MoveCost() As Integer
End Type

Public Type TCharacter
    Index As Integer
    x As Integer
    y As Integer
    HealthLeft As Integer
    Nplayer As Integer
    ScreenX As Double
    ScreenY As Double
    Agressive As Boolean
    TextNumber As Integer
End Type

Public Enum Actions
    Wait = 0
    Move = 1
    Attack = 2
End Enum

Public Enum ButtonList
    MoveBtn
    AttackBtn
    StopBtn
    PatrolBtn
    AgrStateBtn
End Enum

Public Const PlayerNumber = 1
Public Const SMapWidth = 230
Public Const SMapHeight = 115
Public Const PI As Double = 3.14159265
Public Const Alpha As Double = 45

Public TileSize As Integer
Public ScreenWidth As Integer
Public ScreenHeight As Integer
Public ScrollX As Integer
Public ScrollY As Integer
Public MainMap As New Map
Public tm As Long
Public fps As Integer
Public Selecting As Boolean
Public MouseX As Single, MouseY As Single
Public SelX As Single, SelY As Single
Public Hdr As MAPFILEHDR
Public ChrInfo() As ChrStandart

Public Function Distance(x1 As Integer, Y1 As Integer, x2 As Integer, Y2 As Integer) As Double
    Distance = ((x1 - x2) ^ 2 + (Y1 - Y2) ^ 2) ^ (1 / 2)
End Function

Public Function CX(xx, yy) As Double
Dim vec As D3DVECTOR

    mDX.VectorRotate vec, Vec3(xx, yy, 0), Vec3(0, 0, 1), (Alpha / 180) * PI
    mDX.VectorScale vec, vec, Distance(xx * TileSize, yy * TileSize, 0, 0)
    CX = vec.x
End Function

Public Function CY(xx, yy) As Double
Dim vec As D3DVECTOR

    mDX.VectorRotate vec, Vec3(xx, yy, 0), Vec3(0, 0, 1), (Alpha / 180) * PI
    mDX.VectorScale vec, vec, Distance(xx * TileSize, yy * TileSize, 0, 0)
    CY = vec.y / 2
End Function

Public Function XC(xx, yy) As Double
Dim vec As D3DVECTOR

    mDX.VectorRotate vec, Vec3(xx, yy * 2, 0), Vec3(0, 0, 1), -(Alpha / 180) * PI
    mDX.VectorScale vec, vec, Distance(CInt(xx), CInt(yy * 2), 0, 0)
    XC = vec.x / TileSize
End Function

Public Function YC(xx, yy) As Double
Dim vec As D3DVECTOR

    mDX.VectorRotate vec, Vec3(xx, yy * 2, 0), Vec3(0, 0, 1), -(Alpha / 180) * PI
    mDX.VectorScale vec, vec, Distance(CInt(xx), CInt(yy * 2), 0, 0)
    YC = vec.y / TileSize
End Function

Public Function Vec3(x, y, z) As D3DVECTOR
    Vec3.x = x
    Vec3.y = y
    Vec3.z = z
End Function
