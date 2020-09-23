Attribute VB_Name = "BmpPacker"
Option Explicit

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type BITMAPINFOHEADER
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type
   
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type BITMAPFILEHEADER
    bfType            As String * 2
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOhFileBits         As Long
End Type

Public Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Public Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Public Type Color
    B As Byte
    G As Byte
    R As Byte
End Type

Private Type Pack_Header
    Magic As String * 6
    Tnumber As Integer
    Tname() As String
    Cnumber As Integer
    Cname() As String
    Mnumber As Integer
    Mname() As String
    MaxSightRange As Integer
    TlSize As Integer
End Type

Public Type Display_Info
    GFX As Long
    OBJ As Long
    BPP As Byte
    FSize As Long
    Wid As Integer
    Hei As Integer
End Type

Dim FileHeader As BITMAPFILEHEADER
Dim InfoHeader As BITMAPINFOHEADER
Dim Color As Color
Dim OldSeek As Long
Dim x As Integer
Dim y As Integer
Dim BmPalette As LOGPALETTE
Dim File As String
Dim PicBits() As Byte
Public MyBMP() As Display_Info

Private Sub Get_Single_BMP(Index As Integer)
    Get #1, , FileHeader
    Get #1, , InfoHeader
    If FileHeader.bfType <> "BM" Then
        MsgBox "Invalid Bitmap Header!->" & Index & vbCr & FileHeader.bfType
        Exit Sub
    End If
    If InfoHeader.biCompression <> 0 Then
        MsgBox "This Bitmap has bitmap compression which is not supported by this program!"
        Exit Sub
    End If
    If InfoHeader.biBitCount <= 4 Then
        MsgBox "1-4 bit images are not supported", 16
        Exit Sub
    End If
    MyBMP(Index).BPP = InfoHeader.biBitCount
    MyBMP(Index).FSize = FileHeader.bfSize
    MyBMP(Index).OBJ = CreateCompatibleBitmap(Form1.hdc, InfoHeader.biWidth, InfoHeader.biHeight)
    MyBMP(Index).GFX = CreateCompatibleDC(Form1.hdc)
    SelectObject MyBMP(Index).GFX, MyBMP(Index).OBJ
    MyBMP(Index).Wid = InfoHeader.biWidth
    MyBMP(Index).Hei = InfoHeader.biHeight
    If InfoHeader.biClrUsed <= 256 And InfoHeader.biBitCount < 24 Then
        CreatePalette
    End If
    If InfoHeader.biBitCount = 8 Then
        Dim ClL As Byte
        Dim ColX As Long
        For y = InfoHeader.biHeight - 1 To -1 Step -1
            For x = 0 To InfoHeader.biWidth - 1
                Get #1, , ClL
                ColX = RGB(BmPalette.palPalEntry(ClL).peRed, BmPalette.palPalEntry(ClL).peGreen, BmPalette.palPalEntry(ClL).peBlue)
                SetPixelV MyBMP(Index).GFX, x, y, ColX
            Next x
        Next y
    Else
        For y = InfoHeader.biHeight - 1 To -1 Step -1
            For x = 0 To InfoHeader.biWidth - 1
                Get #1, , Color
                SetPixelV MyBMP(Index).GFX, x, y, RGB(Color.R, Color.G, Color.B)
            Next x
        Next y
    End If
    Seek #1, OldSeek + FileHeader.bfSize
    OldSeek = Seek(1)
End Sub

Public Sub OpenBMPPack(File As String)
Dim i As Integer
Dim HHL As Pack_Header

    If File = "" Or Dir(File) = "" Then Exit Sub
    Open File For Binary As #1
    Get #1, , HHL
    If HHL.Magic <> "TTLSET" Then
        MsgBox "Invalid TTLSET file!", 16
        Exit Sub
    End If
    MainMap.MaxSightRange = HHL.MaxSightRange
    TileSize = HHL.TlSize
    OldSeek = Seek(1)
    ReDim MyBMP(HHL.Tnumber - 1)
    For i = 1 To HHL.Tnumber
        Get_Single_BMP i - 1
        DoEvents
    Next
    CreateTextures msTexture, Black
    ReDim ChrInfo(HHL.Cnumber - 1)
    ReDim MyBMP(HHL.Cnumber - 1)
    Get #1, , ChrInfo
    OldSeek = Seek(1)
    For i = 1 To HHL.Cnumber
        Get_Single_BMP i - 1
        DoEvents
    Next
    CreateTextures msCharacter, Black
    ReDim MyBMP(HHL.Mnumber - 1)
    For i = 1 To HHL.Mnumber
        Get_Single_BMP i - 1
        DoEvents
    Next
    CreateTextures msTextureMask, Black
Close #1
End Sub

Private Sub CreatePalette()
Dim i As Long
Dim BlueByte As Byte
Dim RedByte As Byte
Dim GreenByte As Byte
Dim AByte As Byte
Dim ClrU As Long
    
    If InfoHeader.biBitCount = 24 Then Exit Sub
    BmPalette.palVersion = &H300
    BmPalette.palNumEntries = InfoHeader.biClrUsed
    If InfoHeader.biClrUsed <> 0 Then ClrU = InfoHeader.biClrUsed - 1
    For i = 0 To ClrU
        Get #1, , BlueByte
        Get #1, , GreenByte
        Get #1, , RedByte
        Get #1, , AByte
        BmPalette.palPalEntry(i).peBlue = BlueByte
        BmPalette.palPalEntry(i).peGreen = GreenByte
        BmPalette.palPalEntry(i).peRed = RedByte
        BmPalette.palPalEntry(i).peFlags = AByte
    Next
End Sub

