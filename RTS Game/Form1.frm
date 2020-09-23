VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   723
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Map Files (*.mpl)|*.mpl"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type info
    Width As Integer
    Height As Integer
    Bit As Integer
End Type
Public SmallMapClicked As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyLeft: ScrollX = -1
        Case Is = vbKeyRight: ScrollX = 1
        Case Is = vbKeyUp: ScrollY = -1
        Case Is = vbKeyDown: ScrollY = 1
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyEscape: Unload Me
        Case Is = vbKeyLeft: ScrollX = 0
        Case Is = vbKeyRight: ScrollX = 0
        Case Is = vbKeyUp: ScrollY = 0
        Case Is = vbKeyDown: ScrollY = 0
    End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim FileName As String, inf As info

    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen
    If Err.Number <> 0 Then End
    FileName = CommonDialog1.FileName
    Me.Show
    Open App.Path & "\settings.dat" For Random As #1 Len = Len(inf)
    Get #1, , inf
    Close #1
    ScreenWidth = inf.Width: ScreenHeight = inf.Height
    InitDX CLng(ScreenWidth), CLng(ScreenHeight), CLng(inf.Bit), "IID_IDirect3DHALDevice"
    Set TileMask = CreateTexture(App.Path & "\mapmask.bmp", 64, 64, COLORKEYOPTIONS.Black)
    Set Selection = CreateTexture(App.Path & "\selection.bmp", 64, 64, COLORKEYOPTIONS.Black)
    MainMap.ReadMap FileName
    begin
End Sub
Private Sub begin()
Dim t(1) As Single
    Do
        DoEvents
        tm = timeGetTime
        MainMap.DoAllActions
        MainMap.Render
        fps = Int(1000 / (timeGetTime - tm))
        MainMap.ScreenX = MainMap.ScreenX + ScrollX * 10 * (100 / fps)
        MainMap.ScreenY = MainMap.ScreenY + ScrollY * 10 * (100 / fps)
        t(0) = Distance(CX(0, 0), CY(0, 0), CX((MainMap.MapX + 1), (MainMap.MapY + 1)), CY((MainMap.MapX + 1), (MainMap.MapY + 1)))
        t(1) = Distance(CX(0, (MainMap.MapY + 1)), CY(0, (MainMap.MapY + 1)), CX((MainMap.MapX + 1), 0), CY((MainMap.MapX + 1), 0))
        If SmallMapClicked Then
            MainMap.ScreenX = (MouseX - 10 - SMapWidth / 2) / (SMapWidth / t(1)) - ScreenWidth / 2
            MainMap.ScreenY = (MouseY - ScreenHeight + 10 + SMapHeight) / (SMapHeight / t(0)) - ScreenHeight / 2
        End If
    Loop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then SmallMapClicked = 10 < x And x < 10 + SMapWidth And ScreenHeight - SMapHeight - 10 < y And y < ScreenHeight - 10
    Selecting = Button = 1 And Not SmallMapClicked
    If Selecting Then SelX = MouseX: SelY = MouseY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
    Select Case x
        Case Is >= ScreenWidth - 1: ScrollX = 1
        Case Is <= 1: ScrollX = -1
        Case Else: ScrollX = 0
    End Select
    Select Case y
        Case Is >= ScreenHeight - 1: ScrollY = 1
        Case Is <= 1: ScrollY = -1
        Case Else: ScrollY = 0
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim arg(3) As Integer, Xa As Integer, Ya As Integer, i As Integer

    Xa = Int(XC(MainMap.ScreenX + x, MainMap.ScreenY + y))
    Ya = Int(YC(MainMap.ScreenX + x, MainMap.ScreenY + y))
    If Xa < 0 Then Xa = 0
    If Ya < 0 Then Ya = 0
    If Xa > MainMap.MapX Then Xa = MainMap.MapX
    If Ya > MainMap.MapY Then Ya = MainMap.MapY
    Select Case Button
        Case Is = 1
            If Not SmallMapClicked Then
                MainMap.ClearSelectedList
                MainMap.SelectCharacters CreateRect(Minn(SelX, MouseX + 1), Minn(SelY, MouseY + 1), Maxx(SelX, MouseX + 1), Maxx(SelY, MouseY + 1))
                Selecting = False
            End If
            SmallMapClicked = False
        Case Is = 2
            If MainMap.nSelectedList = 0 Then Exit Sub
            arg(0) = Xa: arg(1) = Ya
            For i = 1 To MainMap.nSelectedList
                With MainMap.GetCharacter(MainMap.GetSelectedChr(i))
                    If .DestX = arg(0) And .DestY = arg(1) Then Exit Sub
                    If .Nplayer = PlayerNumber Then
                        .AddCommand Actions.Move, arg, Shift = 1
                    End If
                End With
            Next i
    End Select
End Sub

