Attribute VB_Name = "Dx7"
Option Explicit

Public Enum COLORKEYOPTIONS
    None
    Black
    White
    Magenta
End Enum

Public Const PI = 3.14159265
Public Const Rad = PI / 180
Public Const LR_LOADFROMFILE = &H10
Public Const LR_CREATEDIBSECTION = &H2000
Public Const SRCCOPY = &HCC0020

Public mDX As New DirectX7
Public mDDraw As DirectDraw7
Public mD3D As Direct3D7
Public mD3DDevice As Direct3DDevice7
Public msFront As DirectDrawSurface7
Public msBack As DirectDrawSurface7
Public msTexture() As DirectDrawSurface7
Public msTextureMask() As DirectDrawSurface7
Public msCharacter() As DirectDrawSurface7
Public TileMask As DirectDrawSurface7
Public SmallMap As DirectDrawSurface7
Public AltBar As DirectDrawSurface7
Public Selection As DirectDrawSurface7
Public SMapDC As Long
Public SelectedChr As Integer

Private Function ExclusiveMode() As Boolean
Dim lngTestExMode As Long
    
    lngTestExMode = mDDraw.TestCooperativeLevel
    If (lngTestExMode = DD_OK) Then
        ExclusiveMode = True
    Else
        ExclusiveMode = False
    End If
End Function

Public Function LostSurfaces() As Boolean
    LostSurfaces = False
    Do Until ExclusiveMode
        DoEvents
        LostSurfaces = True
    Loop
    DoEvents
    If LostSurfaces Then
        mDDraw.RestoreAllSurfaces
    End If
End Function

Public Function CalcCoordX(Offset As Long, Length As Long, Angle As Integer)
    CalcCoordX = Offset + Sin(Angle * Rad) * Length
End Function

Public Function CalcCoordY(Offset As Long, Length As Long, Angle As Integer)
    CalcCoordY = Offset + Cos(Angle * Rad) * Length
End Function

Public Function CreateSurface(File As String, Width As Long, Height As Long, Optional ColKey As COLORKEYOPTIONS = None)
Dim msSurface As DirectDrawSurface7
Dim ddsd As DDSURFACEDESC2
Dim cKey As DDCOLORKEY
Dim ddpf As DDPIXELFORMAT
    
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd.lHeight = Width
    ddsd.lWidth = Height
    If File = "" Then
        Set msSurface = mDDraw.CreateSurface(ddsd)
    Else
        Set msSurface = mDDraw.CreateSurfaceFromFile(File, ddsd)
    End If
    Select Case ColKey
        Case COLORKEYOPTIONS.Black
            cKey.low = 0
            cKey.high = 0
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
        Case COLORKEYOPTIONS.White
            msSurface.GetPixelFormat ddpf
            cKey.low = ddpf.lRBitMask + ddpf.lGBitMask + ddpf.lBBitMask
            cKey.high = cKey.low
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
        Case COLORKEYOPTIONS.Magenta
            msSurface.GetPixelFormat ddpf
            cKey.low = ddpf.lRBitMask + ddpf.lBBitMask
            cKey.high = cKey.low
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
    End Select
    Set CreateSurface = msSurface
End Function

Public Function RGB2DX(R As Single, G As Single, B As Single) As Long
    RGB2DX = mDX.CreateColorRGBA(CSng((1 / 255) * R), CSng((1 / 255) * G), CSng((1 / 255) * B), 1)
End Function

Public Function CreateTexture(File As String, Width As Long, Height As Long, Optional ColKey As COLORKEYOPTIONS = 0) As DirectDrawSurface7
Dim enumTex As Direct3DEnumPixelFormats
Dim msSurface As DirectDrawSurface7
Dim ddsd As DDSURFACEDESC2
Dim bOK As Boolean
Dim lK As Long
Dim cKey As DDCOLORKEY
Dim ddpf As DDPIXELFORMAT

    ddsd.lFlags = DDSD_CAPS Or DDSD_TEXTURESTAGE Or DDSD_PIXELFORMAT
    If Height <> 0 And Width <> 0 Then
        ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        ddsd.lHeight = Height
        ddsd.lWidth = Width
    End If
    Set enumTex = mD3DDevice.GetTextureFormatsEnum()
    For lK = 1 To enumTex.GetCount()
        bOK = True
        Call enumTex.GetItem(lK, ddsd.ddpfPixelFormat)
        With ddsd.ddpfPixelFormat
            If .lRGBBitCount <> 16 Then bOK = False
        End With
        If bOK = True Then Exit For
    Next
    If bOK = False Then
        Err.Raise 8001, , "No support for 16 bit textures found!"
        Exit Function
    End If
    If mD3DDevice.GetDeviceGuid() = "IID_IDirect3DHALDevice" Then
        ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
        ddsd.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        ddsd.lTextureStage = 0
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
        ddsd.ddsCaps.lCaps2 = 0
        ddsd.lTextureStage = 0
    End If
    If File = "" Then
        Set msSurface = mDDraw.CreateSurface(ddsd)
    Else
        Set msSurface = mDDraw.CreateSurfaceFromFile(File, ddsd)
    End If
    Select Case ColKey
        Case COLORKEYOPTIONS.Black
            cKey.low = 0
            cKey.high = 0
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
        Case COLORKEYOPTIONS.White
            msSurface.GetPixelFormat ddpf
            cKey.low = ddpf.lRBitMask + ddpf.lGBitMask + ddpf.lBBitMask
            cKey.high = cKey.low
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
        Case COLORKEYOPTIONS.Magenta
            msSurface.GetPixelFormat ddpf
            cKey.low = ddpf.lRBitMask + ddpf.lBBitMask
            cKey.high = cKey.low
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
    End Select
    Set CreateTexture = msSurface
End Function

Public Sub InitDX(Width As Long, Height As Long, Depth As Long, DeviceGUID As String)
Dim ddsd As DDSURFACEDESC2
Dim caps As DDSCAPS2
Dim cKey As DDCOLORKEY, pf As DDPIXELFORMAT

    Set mDDraw = mDX.DirectDrawCreate("")
    mDDraw.SetCooperativeLevel Form1.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT
    mDDraw.SetDisplayMode Width, Height, Depth, 0, DDSDM_DEFAULT
    ddsd.lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_3DDEVICE Or DDSCAPS_PRIMARYSURFACE
    ddsd.lBackBufferCount = 1
    Set msFront = mDDraw.CreateSurface(ddsd)
    caps.lCaps = DDSCAPS_BACKBUFFER Or DDSCAPS_3DDEVICE
    Set msBack = msFront.GetAttachedSurface(caps)
    Set mD3D = mDDraw.GetDirect3D
    Set mD3DDevice = mD3D.CreateDevice(DeviceGUID, msBack)
    mD3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, 1
    With ddsd
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
        .lWidth = SMapWidth
        .lHeight = SMapHeight
    End With
    Set SmallMap = mDDraw.CreateSurfaceFromFile(App.Path & "\mapmask.bmp", ddsd)
     SmallMap.GetPixelFormat pf
    cKey.low = pf.lBBitMask + pf.lRBitMask: cKey.high = cKey.low
    SmallMap.SetColorKey DDCKEY_SRCBLT, cKey
    ddsd.lWidth = 800: ddsd.lHeight = 150
End Sub

Public Sub ClearDevice()
Dim rClear(0) As D3DRECT

    rClear(0).x2 = ScreenWidth
    rClear(0).Y2 = ScreenHeight
    mD3DDevice.Clear 1, rClear, D3DCLEAR_TARGET, RGB2DX(0, 0, 0), 0, 0
End Sub

Public Sub CreateTextures(Ms() As DirectDrawSurface7, Optional ColKey As COLORKEYOPTIONS = COLORKEYOPTIONS.None)
Dim TEMPDXD As DDSURFACEDESC2, dcDXS As Long, i As Integer, i2
Dim cKey As DDCOLORKEY, ddpf As DDPIXELFORMAT
Dim msSurface As DirectDrawSurface7

    For i = 0 To UBound(MyBMP)
        ReDim Preserve Ms(i)
        With TEMPDXD
            .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            .ddsCaps.lCaps = DDSCAPS_TEXTURE
            .ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
            .lWidth = MyBMP(i).Wid
            .lHeight = MyBMP(i).Hei
            .lTextureStage = 0
        End With
        Set Ms(i) = mDDraw.CreateSurface(TEMPDXD)
        Ms(i).restore
        dcDXS = Ms(i).GetDC()
        BitBlt dcDXS, 0, 0, MyBMP(i).Wid, MyBMP(i).Hei, MyBMP(i).GFX, 0, 0, vbSrcCopy
        Ms(i).ReleaseDC dcDXS
        Select Case ColKey
            Case COLORKEYOPTIONS.Black
                Ms(i).GetPixelFormat ddpf
                cKey.low = 0
                cKey.high = ddpf.lRBitMask + ddpf.lGBitMask + ddpf.lBBitMask
            Case COLORKEYOPTIONS.White
                Ms(i).GetPixelFormat ddpf
                cKey.low = ddpf.lRBitMask + ddpf.lGBitMask + ddpf.lBBitMask
                cKey.high = cKey.low
            Case COLORKEYOPTIONS.Magenta
                Ms(i).GetPixelFormat ddpf
                cKey.low = ddpf.lRBitMask + ddpf.lBBitMask
                cKey.high = cKey.low
        End Select
        Ms(i).SetColorKey DDCKEY_SRCBLT, cKey
    Next i
End Sub

Public Function CreateRect(left As Long, top As Long, right As Long, bottom As Long) As RECT
Dim s As RECT

    s.left = left
    s.bottom = bottom
    s.right = right
    s.top = top
    CreateRect = s
End Function

Public Function Minn(a1 As Single, a2 As Single) As Single
    If a1 < a2 Then Minn = a1 Else Minn = a2
End Function

Public Function Maxx(a1 As Single, a2 As Single) As Single
    If a1 > a2 Then Maxx = a1 Else Maxx = a2
End Function
