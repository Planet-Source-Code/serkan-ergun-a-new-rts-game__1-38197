VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MapX As Integer
Public MapY As Integer
Public ScreenX As Integer
Public MaxSightRange As Integer
Public ScreenY As Integer

Private mtlTile(3) As D3DTLVERTEX
Private MapData() As MapTile
Private TileData() As TileVertex
Private Characters() As New Character
Private SelectedChrs() As Integer
Private RemovingList() As Integer
Private TileRect As RECT

Public Sub AddCharacter(Armor As Integer, AttackRange As Integer, Health As Integer, Nplayer As Integer, SightRange As Integer, x As Integer, y As Integer, Speed As Integer, AttackRate As Integer, AttackDamage As Integer)
Dim i As Integer

    i = nCharacter + 1
    ReDim Preserve Characters(i)
    With Characters(i)
        .Armor = Armor
        .AttackRange = AttackRange
        .Health = Health
        .HealthLeft = Health
        .Index = i
        .Nplayer = Nplayer
        .SightRange = SightRange
        .x = x
        .y = y
        .Speed = Speed
        .AttackRate = AttackRate
        .AttackDamage = AttackDamage
        .Agressive = True
        .TextNumber = Nplayer - 1
    End With
    MapData(x, y).Character = i
    CalcVertexBrightness i
    CheckCharacter i
    Characters(i).ScreenX = CX(x + 0.5, y + 0.5) - TileSize / 2
    Characters(i).ScreenY = CY(x + 0.5, y + 0.5) - TileSize
End Sub

Public Sub CheckCharacter(Index As Integer)
Dim MinX As Integer, Maxx As Integer
Dim MinY As Integer, MaxY As Integer
Dim ix As Integer, iy As Integer
Dim CenterX As Integer, CenterY As Integer
Dim dist As Double

    With Characters(Index)
        MinX = .x - MaxSightRange: MinY = .y - MaxSightRange
        Maxx = .x + MaxSightRange: MaxY = .y + MaxSightRange
        CenterX = .x: CenterY = .y
    End With
    If MinX < 0 Then MinX = 0
    If MinY < 0 Then MinY = 0
    If Maxx > MapX Then Maxx = MapX
    If MaxY > MapY Then MaxY = MapY
    For ix = MinX To Maxx
        For iy = MinY To MaxY
            If Not MapData(ix, iy).Character = 0 Then
                dist = Distance(ix, iy, CenterX, CenterY)
                If dist <= Characters(MapData(ix, iy).Character).SightRange Then
                    Characters(MapData(ix, iy).Character).AddToSeenList Index
                    Characters(MapData(CenterX, CenterY).Character).AddToSeenFromList MapData(ix, iy).Character
                End If
                If dist <= Characters(MapData(CenterX, CenterY).Character).SightRange Then
                    Characters(MapData(CenterX, CenterY).Character).AddToSeenList MapData(ix, iy).Character
                    Characters(MapData(ix, iy).Character).AddToSeenFromList Index
                End If
            End If
        Next iy
    Next ix
End Sub

Public Sub MoveCharacter(Index As Integer, x As Integer, y As Integer)
Dim i As Integer, Color As Long
    
    For i = 1 To Characters(Index).nSeenFromList
        With Characters(Characters(Index).GetSeenFromList(i))
            .RemoveFromSeenList .GetSeenListVariable(Index)
        End With
    Next i
    For i = 1 To Characters(Index).nSeenList
        With Characters(Characters(Index).GetSeenList(i))
            .RemoveFromSeenFromList .GetSeenFromListVariable(Index)
        End With
    Next i
    CalcVertexBrightness Index, False
    With Characters(Index)
        MapData(.x, .y).Character = 0
        .x = x: .y = y
        MapData(.x, .y).Character = Index
    End With
    CalcVertexBrightness Index
    Characters(Index).ClearSeenList
    Characters(Index).ClearSeenFromList
    CheckCharacter Index
End Sub

Public Sub RemoveCharacter(Index As Integer)
Dim i As Integer
    
    For i = 1 To Characters(Index).nSeenFromList
        With Characters(Characters(Index).GetSeenFromList(i))
            .RemoveFromSeenList .GetSeenListVariable(Index)
        End With
    Next i
    For i = 1 To Characters(Index).nSeenList
        With Characters(Characters(Index).GetSeenList(i))
            .RemoveFromSeenFromList .GetSeenFromListVariable(Index)
        End With
    Next i
    CalcVertexBrightness Index, False
    MapData(Characters(Index).x, Characters(Index).y).Character = 0
    If nCharacter = Index Then ReDim Preserve Characters(nCharacter - 1): Exit Sub
    For i = 1 To Characters(nCharacter).nSeenFromList
        With Characters(Characters(nCharacter).GetSeenFromList(i))
            .RemoveFromSeenList .GetSeenListVariable(nCharacter)
            If .DestAttack = nCharacter Then .DestAttack = Index
        End With
    Next i
    For i = 1 To Characters(nCharacter).nSeenList
        With Characters(Characters(nCharacter).GetSeenList(i))
            .RemoveFromSeenFromList .GetSeenFromListVariable(nCharacter)
        End With
    Next i
    Set Characters(Index) = Characters(nCharacter)
    ChangeSelIndex nCharacter, Index
    ChangeRemovingIndex nCharacter, Index
    Characters(Index).Index = Index
    MapData(Characters(Index).x, Characters(Index).y).Character = Index
    ReDim Preserve Characters(nCharacter - 1)
    CheckCharacter Index
End Sub

Public Function GetCharacter(Index As Integer) As Character
    Set GetCharacter = Characters(Index)
End Function

Public Function nCharacter() As Integer
    nCharacter = UBound(Characters)
End Function

Private Sub Class_Initialize()
    ReDim MapData(0, 0)
    ReDim Characters(0)
    ReDim SelectedChrs(0)
    ReDim RemovingList(0)
    TileRect.bottom = TileSize
    TileRect.right = TileSize
End Sub

Public Sub ClearCharacters()
    ReDim Characters(0)
    ReDim MapData(MapX, MapY)
End Sub

Public Function GetMapCharacter(x As Integer, y As Integer) As Integer
    GetMapCharacter = MapData(x, y).Character
End Function

Public Function GetMapLayer(x As Integer, y As Integer) As Integer
    GetMapLayer = MapData(x, y).layer
End Function

Public Sub DoAllActions()
Dim i As Integer
    
    For i = 1 To nCharacter
        DoEvents
        Characters(i).DoAction
    Next i
    For i = 1 To nRemovingList
        RemoveCharacter RemovingList(i)
    Next i
    ClearRemovingList
End Sub

Public Sub CalcVertexBrightness(CharIndex As Integer, Optional Adding As Boolean = True)
Dim lXDistance As Single
Dim lYDistance As Single
Dim lVertexDistance As Single
Dim lLightSteps As Single
Dim MinX As Integer, MinY As Integer
Dim Maxx As Integer, MaxY As Integer
Dim i As Integer, j As Integer
Dim tlx As Integer, tly As Integer
Dim t(1) As Single
Dim Dest As RECT
Dim cKey As DDCOLORKEY
Dim ddsd As DDSURFACEDESC2, pf As DDPIXELFORMAT
Dim Tile As DirectDrawSurface7

   With Characters(CharIndex)
        If .Nplayer = PlayerNumber Then
            MinX = .x - .SightRange: MinY = .y - .SightRange
            Maxx = .x + .SightRange: MaxY = .y + .SightRange
            If MinX < 0 Then MinX = 0
            If MinY < 0 Then MinY = 0
            If Maxx > MapX + 1 Then Maxx = MapX + 1
            If MaxY > MapY + 1 Then MaxY = MapY + 1
            For i = MinX To Maxx
                For j = MinY To MaxY
                    lXDistance = CSng(Abs(.x - i + 0.5))
                    lYDistance = CSng(Abs(.y - j + 0.5))
                    lVertexDistance = CSng(Sqr(lXDistance ^ 2 + lYDistance ^ 2))
                    If lVertexDistance <= .SightRange Then
                        lLightSteps = CSng(200 / .SightRange)
                        If Adding Then
                            If Not TileData(i, j).Explored Then
                                If i > MapX Then tlx = MapX Else tlx = i
                                If j > MapY Then tly = MapY Else tly = j
                                t(0) = Distance(CX(0, 0), CY(0, 0), CX((MainMap.MapX + 1), (MainMap.MapY + 1)), CY((MainMap.MapX + 1), (MainMap.MapY + 1)))
                                t(1) = Distance(CX(0, (MainMap.MapY + 1)), CY(0, (MainMap.MapY + 1)), CX((MainMap.MapX + 1), 0), CY((MainMap.MapX + 1), 0))
                                Dest = CreateRect(CX(tlx, tly + 1) * (SMapWidth / t(1)) + SMapWidth / 2, CY(tlx, tly) * (SMapHeight / t(0)), CX(tlx + 1, tly) * (SMapWidth / t(1)) + SMapWidth / 2, CY(tlx + 1, tly + 1) * (SMapHeight / t(0)))
                                If (Not Dest.left = Dest.right) And (Not Dest.top = Dest.bottom) Then
                                    With ddsd
                                        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
                                        .ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
                                        .lWidth = TileSize
                                        .lHeight = TileSize
                                    End With
                                    Set Tile = mDDraw.CreateSurface(ddsd)
                                    Tile.Blt TileRect, msTexture(MapData(tlx, tly).layer), TileRect, DDBLT_DONOTWAIT
                                    Tile.Blt TileRect, TileMask, TileRect, DDBLT_KEYSRC
                                    Tile.GetPixelFormat pf
                                    cKey.low = pf.lRBitMask + pf.lBBitMask: cKey.high = cKey.low
                                    Tile.SetColorKey DDCKEY_SRCBLT, cKey
                                    SmallMap.Blt Dest, Tile, TileRect, DDBLT_KEYSRC
                                End If
                                TileData(i, j).Brightness = TileData(i, j).Brightness + 55
                                TileData(i, j).Explored = True
                            End If
                            TileData(i, j).Brightness = TileData(i, j).Brightness + CSng(255 - (lVertexDistance * lLightSteps))
                        Else
                            TileData(i, j).Brightness = TileData(i, j).Brightness - CSng(255 - (lVertexDistance * lLightSteps))
                        End If
                    End If
                Next j
            Next i
        End If
    End With
End Sub

Public Sub Render()
Dim MinX As Integer, MinY As Integer
Dim Maxx As Integer, MaxY As Integer
Dim i As Integer, j As Integer, k As Integer
Dim cl As Single
Dim tlx As Double, tly As Double
Dim xx(3) As Double
Dim yy(3) As Double
Dim t(1) As Single

    msBack.SetForeColor RGB2DX(255, 255, 255)
    tlx = XC(ScreenX, ScreenY)
    tly = YC(ScreenX + ScreenWidth, ScreenY)
    MinX = tlx - 1
    MinY = tly - 1
    If MinX < 0 Then MinX = 0
    If MinY < 0 Then MinY = 0
    tlx = XC(ScreenX + ScreenWidth, ScreenY + ScreenHeight)
    tly = YC(ScreenX, ScreenY + ScreenHeight)
    Maxx = tlx + 1
    MaxY = tly + 1
    If Maxx > MapX Then Maxx = MapX
    If MaxY > MapY Then MaxY = MapY
    ClearDevice
    mD3DDevice.BeginScene
    'Rendering Main Screen
    For i = MinX To Maxx
        For j = MinY To MaxY
            cl = TileData(i, j).Brightness: If cl > 255 Then cl = 255
            xx(0) = CX(i, j): yy(0) = CY(i, j)
            xx(3) = CX(i + 1, j + 1): yy(3) = CY(i + 1, j + 1)
            If xx(0) > ScreenX - TileSize And yy(0) > ScreenY - TileSize And xx(3) < ScreenX + ScreenWidth + TileSize And yy(3) < ScreenY + ScreenHeight + TileSize Then
                xx(1) = CX(i, j + 1): yy(1) = CY(i, j + 1)
                xx(2) = CX(i + 1, j): yy(2) = CY(i + 1, j)
                cl = TileData(i, j).Brightness: If cl > 255 Then cl = 255
                mDX.CreateD3DTLVertex xx(0) - ScreenX, yy(0) - ScreenY, 0, 1, RGB2DX(cl, cl, cl), 0, 0, 0, mtlTile(0)
                cl = TileData(i + 1, j).Brightness: If cl > 255 Then cl = 255
                mDX.CreateD3DTLVertex xx(2) - ScreenX, yy(2) - ScreenY, 0, 1, RGB2DX(cl, cl, cl), 0, 1, 0, mtlTile(1)
                cl = TileData(i, j + 1).Brightness: If cl > 255 Then cl = 255
                mDX.CreateD3DTLVertex xx(1) - ScreenX, yy(1) - ScreenY, 0, 1, RGB2DX(cl, cl, cl), 0, 0, 1, mtlTile(2)
                cl = TileData(i + 1, j + 1).Brightness: If cl > 255 Then cl = 255
                mDX.CreateD3DTLVertex xx(3) - ScreenX, yy(3) - ScreenY, 0, 1, RGB2DX(cl, cl, cl), 0, 1, 1, mtlTile(3)
                mD3DDevice.SetTexture 0, msTexture(MapData(i, j).layer)
                mD3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlTile(0), 4, D3DDP_DEFAULT
                If Not MapData(i, j).layer2 = 0 Then
                    mD3DDevice.SetTexture 0, msTextureMask(MapData(i, j).layer2 - 1)
                    mD3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlTile(0), 4, D3DDP_DEFAULT
                End If
            End If
        Next j
    Next i
    'Rendering Characters
    For i = MinX To Maxx
        For j = MinY To MaxY
            k = MapData(i, j).Character
            If Not k = 0 Then
                With GetCharacter(k)
                    If .Nplayer = PlayerNumber Or .SeenFromPlayerNum(PlayerNumber) Then
                        If isSelected(k) Then
                            tlx = XC(.ScreenX + TileSize / 2, .ScreenY + TileSize) - 0.5
                            tly = YC(.ScreenX + TileSize / 2, .ScreenY + TileSize) - 0.5
                            mDX.CreateD3DTLVertex CX(tlx, tly) - ScreenX, CY(tlx, tly) - ScreenY, 0, 1, RGB2DX(255, 255, 255), 0, 0, 0, mtlTile(0)
                            mDX.CreateD3DTLVertex CX((tlx + 1), tly) - ScreenX, CY((tlx + 1), tly) - ScreenY, 0, 1, RGB2DX(255, 255, 255), 0, 1, 0, mtlTile(1)
                            mDX.CreateD3DTLVertex CX(tlx, tly + 1) - ScreenX, CY(tlx, tly + 1) - ScreenY, 0, 1, RGB2DX(255, 255, 255), 0, 0, 1, mtlTile(2)
                            mDX.CreateD3DTLVertex CX(tlx + 1, tly + 1) - ScreenX, CY(tlx + 1, tly + 1) - ScreenY, 0, 1, RGB2DX(255, 255, 255), 0, 1, 1, mtlTile(3)
                            mD3DDevice.SetTexture 0, Selection
                            mD3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlTile(0), 4, D3DDP_DEFAULT
                        End If
                        mDX.CreateD3DTLVertex CX(XC(.ScreenX, .ScreenY + TileSize), YC(.ScreenX, .ScreenY + TileSize) - 1) - ScreenX, CY(XC(.ScreenX, .ScreenY + TileSize), YC(.ScreenX, .ScreenY + TileSize) - 1) - ScreenY, 0, 1, 0, 0, .Direction * (1 / 4), .TexFrame * (1 / 5), mtlTile(0)
                        mDX.CreateD3DTLVertex CX(XC(.ScreenX + TileSize, .ScreenY + TileSize), YC(.ScreenX + TileSize, .ScreenY + TileSize) - 1) - ScreenX, CY(XC(.ScreenX + TileSize, .ScreenY + TileSize), YC(.ScreenX + TileSize, .ScreenY + TileSize) - 1) - ScreenY, 0, 1, 0, 0, (.Direction + 1) * (1 / 4), .TexFrame * (1 / 5), mtlTile(1)
                        mDX.CreateD3DTLVertex .ScreenX - ScreenX, .ScreenY + TileSize - ScreenY, 0, 1, 0, 0, .Direction * (1 / 4), (.TexFrame + 1) * (1 / 5), mtlTile(2)
                        mDX.CreateD3DTLVertex .ScreenX + TileSize - ScreenX, .ScreenY + TileSize - ScreenY, 0, 1, 0, 0, (.Direction + 1) * (1 / 4), (.TexFrame + 1) * (1 / 5), mtlTile(3)
                        mD3DDevice.SetTexture 0, msCharacter(.TextNumber)
                        mD3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlTile(0), 4, D3DDP_DEFAULT
                        mDX.CreateD3DTLVertex .ScreenX - ScreenX, .ScreenY - ScreenY, 0, 1, RGB2DX(255, 255, 255), 0, .Direction * (1 / 4), .TexFrame * (1 / 5), mtlTile(0)
                        mDX.CreateD3DTLVertex .ScreenX + TileSize - ScreenX, .ScreenY - ScreenY, 0, 1, RGB2DX(255, 255, 255), 0, (.Direction + 1) * (1 / 4), .TexFrame * (1 / 5), mtlTile(1)
                        mDX.CreateD3DTLVertex .ScreenX - ScreenX, .ScreenY + TileSize - ScreenY, 0, 1, RGB2DX(255, 255, 255), 0, .Direction * (1 / 4), (.TexFrame + 1) * (1 / 5), mtlTile(2)
                        mDX.CreateD3DTLVertex .ScreenX + TileSize - ScreenX, .ScreenY + TileSize - ScreenY, 0, 1, RGB2DX(255, 255, 255), 0, (.Direction + 1) * (1 / 4), (.TexFrame + 1) * (1 / 5), mtlTile(3)
                        mD3DDevice.SetTexture 0, msCharacter(.TextNumber)
                        mD3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlTile(0), 4, D3DDP_DEFAULT
                    End If
                End With
            End If
        Next j
    Next i
    mD3DDevice.EndScene
    msBack.Blt CreateRect(10, ScreenHeight - SMapHeight - 10, 10 + SMapWidth, ScreenHeight - 10), SmallMap, CreateRect(0, 0, SMapWidth, SMapHeight), DDBLT_KEYSRC
    t(0) = Distance(CX(0, 0), CY(0, 0), CX((MainMap.MapX + 1), (MainMap.MapY + 1)), CY((MainMap.MapX + 1), (MainMap.MapY + 1)))
    t(1) = Distance(CX(0, (MainMap.MapY + 1)), CY(0, (MainMap.MapY + 1)), CX((MainMap.MapX + 1), 0), CY((MainMap.MapX + 1), 0))
   If Selecting Then msBack.DrawBox MouseX, MouseY, SelX, SelY
    'rendering Characters to Small Map
    For i = 1 To nCharacter
        With Characters(i)
            If .Nplayer = PlayerNumber Or .SeenFromPlayerNum(PlayerNumber) Then
                msBack.BltColorFill CreateRect(10 + .ScreenX * (SMapWidth / t(1)) + SMapWidth / 2, ScreenHeight - 10 - SMapHeight + .ScreenY * (SMapHeight / t(0)), 10 + (.ScreenX + TileSize) * (SMapWidth / t(1)) + SMapWidth / 2, ScreenHeight - 10 - SMapHeight + (.ScreenY + TileSize) * SMapHeight / (t(0))), vbWhite
            End If
        End With
    Next i
    msBack.DrawBox 10 + ScreenX * (SMapWidth / t(1)) + SMapWidth / 2, ScreenHeight - 10 - SMapHeight + ScreenY * (SMapHeight / t(0)), 10 + (ScreenX + ScreenWidth) * (SMapWidth / t(1)) + SMapWidth / 2, ScreenHeight - 10 - SMapHeight + (ScreenY + ScreenHeight) * (SMapHeight / (t(0)))
    msFront.Flip Nothing, 0
End Sub

Public Sub SelectCharacters(rct As RECT)
Dim MinX As Integer, Maxx As Integer, MinY As Integer, MaxY As Integer
Dim i As Integer, j As Integer

    MinX = XC(rct.left + ScreenX - TileSize / 2, rct.top + ScreenY - TileSize / 2)
    Maxx = XC(rct.right + ScreenX - TileSize / 2, rct.bottom + ScreenY - TileSize / 2)
    MinY = YC(rct.right + ScreenX - TileSize / 2, rct.top + ScreenY - TileSize / 2)
    MaxY = YC(rct.left + ScreenX - TileSize / 2, rct.bottom + ScreenY - TileSize / 2)
    If Maxx > MapX Then Maxx = MapX
    If MaxY > MapY Then MaxY = MapY
    If MinX < 0 Then MinX = 0
    If MinY < 0 Then MinY = 0
    rct.left = rct.left + ScreenX
    rct.right = rct.right + ScreenX
    rct.top = rct.top + ScreenY
    rct.bottom = rct.bottom + ScreenY
    For i = MinX To Maxx
        For j = MinY To MaxY
            If Not GetMapCharacter(i, j) = 0 Then
                With Characters(GetMapCharacter(i, j))
                        If .Nplayer = PlayerNumber Then AddToSelectedList GetMapCharacter(i, j)
                End With
            End If
        Next j
    Next i
End Sub

Public Sub AddToSelectedList(ChrIndex As Integer)
Dim i As Integer
    
    For i = 1 To nSelectedList
        If ChrIndex = SelectedChrs(i) Then Exit Sub
    Next i
    ReDim Preserve SelectedChrs(nSelectedList + 1)
    SelectedChrs(nSelectedList) = ChrIndex
End Sub

Public Function nSelectedList() As Integer
    nSelectedList = UBound(SelectedChrs)
End Function

Public Sub ClearSelectedList()
    ReDim SelectedChrs(0)
End Sub

Public Function GetSelectedChr(ListIndex As Integer) As Integer
    GetSelectedChr = SelectedChrs(ListIndex)
End Function

Public Function isSelected(ind As Integer) As Boolean
Dim i As Integer
    
    isSelected = False
    For i = 1 To nSelectedList
        If SelectedChrs(i) = ind Then isSelected = True: Exit Function
    Next i
End Function

Public Sub ReadMap(FileName As String)
Dim Hdr As MAPFILEHDR, TCharacters() As TCharacter, i As Integer

    Open FileName For Binary As #1
    Get #1, , Hdr
    If Hdr.EngName = "Serkanýnki      " And Hdr.EngVersion = App.Major Then
        ReDim MapData(Hdr.MapWidth, Hdr.MapHeight)
        ReDim TileData(Hdr.MapWidth, Hdr.MapHeight)
        Get #1, , MapData
        MapX = Hdr.MapWidth - 1: MapY = Hdr.MapHeight - 1
        ReDim TCharacters(Hdr.CharacterNumber)
        ReDim Characters(Hdr.CharacterNumber + 1)
        Get #1, , TCharacters
    End If
    Close #1
    OpenBMPPack Hdr.TileSet
    For i = 0 To Hdr.CharacterNumber
        With Characters(i + 1)
            .Agressive = TCharacters(i).Agressive
            .Armor = ChrInfo(TCharacters(i).TextNumber).Armor
            .AttackDamage = ChrInfo(TCharacters(i).TextNumber).AttackDamage
            .AttackRange = ChrInfo(TCharacters(i).TextNumber).AttackRange
            .AttackRate = ChrInfo(TCharacters(i).TextNumber).AttackRate
            .Health = ChrInfo(TCharacters(i).TextNumber).Health
            .HealthLeft = TCharacters(i).HealthLeft
            .Index = TCharacters(i).Index
            .Nplayer = TCharacters(i).Nplayer
            .ScreenX = TCharacters(i).ScreenX
            .ScreenY = TCharacters(i).ScreenY
            .SightRange = ChrInfo(TCharacters(i).TextNumber).SightRange
            .Speed = ChrInfo(TCharacters(i).TextNumber).Speed
            .TextNumber = TCharacters(i).TextNumber
            .x = TCharacters(i).x
            .y = TCharacters(i).y
            MapData(.x, .y).Character = .Index
        End With
        CheckCharacter i + 1
        CalcVertexBrightness i + 1
    Next i
End Sub

Public Sub ChangeSelIndex(Changing As Integer, ChangingTo As Integer)
Dim i As Integer

    For i = 1 To nSelectedList
        If SelectedChrs(i) = Changing Then
            SelectedChrs(i) = ChangingTo
            Exit Sub
        End If
    Next i
End Sub

Public Sub ChangeRemovingIndex(Changing As Integer, ChangingTo As Integer)
Dim i As Integer

    For i = 1 To nRemovingList
        If RemovingList(i) = Changing Then
            RemovingList(i) = ChangingTo
            Exit Sub
        End If
    Next i
End Sub

Public Sub AddToRemovingList(CIndex As Integer)
Dim i As Integer

    For i = 1 To nRemovingList
        If RemovingList(i) = CIndex Then Exit Sub
    Next i
    ReDim Preserve RemovingList(nRemovingList + 1)
    RemovingList(nRemovingList) = CIndex
End Sub

Public Function nRemovingList()
    nRemovingList = UBound(RemovingList)
End Function

Public Sub ClearRemovingList()
    ReDim RemovingList(0)
End Sub
