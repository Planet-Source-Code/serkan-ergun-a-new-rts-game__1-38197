VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Character"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MarkD
    Cost As Single
    Set As Long
    Direction As Long
End Type

Private Type Command
    Action As Actions
    Arguement() As Integer
End Type

Private Type Point
    x As Double
    y As Double
End Type

Public index As Integer
Public x As Integer
Public y As Integer
Public Health As Integer
Public HealthLeft As Integer
Public Armor As Integer
Public Nplayer As Integer
Public SightRange As Integer
Public Speed As Integer
Public AttackRange As Integer
Public DestX As Integer, DestY As Integer
Public CommandFinished As Boolean
Public ScreenX As Double
Public ScreenY As Double
Public AttackRate As Integer
Public AttackDamage As Integer
Public Agressive As Boolean
Public DestAttack As Integer
Public TextNumber As Integer
Public Direction As Integer
Public TexFrame As Integer
Private FindingTime As Long
Private CheckingTileMove As Integer
Private SeenList() As Integer 'includes indexes of the characters that this character sees
Private SeenFromList() As Integer ' includes indexes of the characters that sees this character
Private MarkData() As MarkD
Private ChrCommands() As Command
Private ControlPoints() As Point
Private MovingTime As Double
Private AttackingTime As Integer
Private Buttons(8) As ButtonList

Public Sub AddToSeenList(SeenIndex As Integer)
Dim i As Integer

    If SeenIndex = index Then Exit Sub
    For i = 1 To nSeenList
        If SeenList(i) = SeenIndex Then Exit Sub
    Next i
    i = nSeenList + 1
    ReDim Preserve SeenList(i)
    SeenList(i) = SeenIndex
End Sub

Public Function nSeenList() As Integer
    nSeenList = UBound(SeenList)
End Function

Public Function GetSeenList(ListIndex) As Integer
    GetSeenList = SeenList(ListIndex)
End Function

Public Sub ClearSeenList()
    ReDim SeenList(0)
End Sub

Public Function GetSeenListVariable(ListVariable As Integer) As Integer
Dim i As Integer

    For i = 1 To nSeenList
        If SeenList(i) = ListVariable Then GetSeenListVariable = i: Exit Function
    Next i
End Function

Public Sub RemoveFromSeenList(ListIndex As Integer)
    SeenList(ListIndex) = SeenList(nSeenList)
    ReDim Preserve SeenList(nSeenList - 1)
End Sub

Private Sub Class_Initialize()
    ReDim SeenList(0)
    ReDim SeenFromList(0)
    ReDim ChrCommands(0)
    ReDim ControlPoints(0)
    CommandFinished = True
End Sub

Public Sub AddToSeenFromList(SeenFromIndex As Integer)
Dim i As Integer
    
    If SeenFromIndex = index Then Exit Sub
    For i = 1 To nSeenFromList
        If SeenFromList(i) = SeenFromIndex Then Exit Sub
    Next i
    i = nSeenFromList + 1
    ReDim Preserve SeenFromList(i)
    SeenFromList(i) = SeenFromIndex
End Sub

Public Function nSeenFromList() As Integer
    nSeenFromList = UBound(SeenFromList)
End Function

Public Function GetSeenFromList(ListIndex As Integer) As Integer
    GetSeenFromList = SeenFromList(ListIndex)
End Function

Public Function GetSeenFromListVariable(ListVariable As Integer) As Integer
Dim i As Integer

    For i = 1 To nSeenFromList
        If SeenFromList(i) = ListVariable Then GetSeenFromListVariable = i: Exit Function
    Next i
End Function

Public Sub ClearSeenFromList()
    ReDim SeenFromList(0)
End Sub

Public Sub RemoveFromSeenFromList(ListIndex As Integer)
    SeenFromList(ListIndex) = SeenFromList(nSeenFromList)
    ReDim Preserve SeenFromList(nSeenFromList - 1)
End Sub

Private Function Cost(pStartX As Integer, pStartY As Integer, pDestinationX As Integer, pDestinationY As Integer) As Single
    Cost = Abs(pStartX - pDestinationX) + Abs(pStartY - pDestinationY)
End Function

Private Function CreateMarkData(pMap As Map, pDestinationX As Integer, pDestinationY As Integer, pStartX As Integer, pStartY As Integer, MinX As Integer, MinY As Integer, Maxx As Integer, MaxY As Integer) As Boolean
Dim OPENSet As New Heap
Dim AddNode As New AStarNode
Dim BestNode As AStarNode
Dim CurrentNode As HeapNode
Dim i As Long, j As Long, k As Long
Dim InsertDirection As Long

    If Not MainMap.GetMapCharacter(pDestinationX, pDestinationY) = 0 Then CreateMarkData = False: Exit Function
    If Not ChrInfo(TextNumber).Movable(MainMap.GetMapLayer(pDestinationX, pDestinationY)) Then CreateMarkData = False: Exit Function
    InsertDirection = 1
    ReDim MarkData(Maxx - MinX, MaxY - MinY)
    AddNode.x = pDestinationX
    AddNode.y = pDestinationY
    AddNode.gval = 0
    AddNode.hval = Cost(pDestinationX, pDestinationY, pStartX, pStartY)
    OPENSet.Add AddNode, AddNode.gval + AddNode.hval, "R" & pDestinationX & "C" & pDestinationY
    MarkData(AddNode.x, AddNode.y).Cost = AddNode.gval + AddNode.hval
    MarkData(AddNode.x, AddNode.y).Set = 1
    Set AddNode = Nothing
    FindingTime = timeGetTime
    Do While OPENSet.Count > 0
        Set CurrentNode = OPENSet.GetLeftMostElement
        Set BestNode = CurrentNode.Item
        If BestNode.x = pStartX And BestNode.y = pStartY Then
            Exit Do
        End If
        If timeGetTime - FindingTime > 100 Then CreateMarkData = False: Exit Function
        If InsertDirection = 1 Then
            i = 0
            j = 3
        Else
            i = 3
            j = 0
        End If
        For i = i To j Step InsertDirection
            Set AddNode = New AStarNode
            Select Case i
                Case 0
                    AddNode.y = BestNode.y - 1
                    AddNode.x = BestNode.x
                Case 1
                    AddNode.x = BestNode.x - 1
                    AddNode.y = BestNode.y
                Case 2
                    AddNode.x = BestNode.x + 1
                    AddNode.y = BestNode.y
                Case 3
                    AddNode.y = BestNode.y + 1
                    AddNode.x = BestNode.x
            End Select
            If AddNode.x < MinX Or AddNode.x > Maxx Or AddNode.y < MinY Or AddNode.y > MaxY Then
            Else
                
                If ChrInfo(TextNumber).Movable(pMap.GetMapLayer(AddNode.x, AddNode.y)) And (pMap.GetMapCharacter(AddNode.x, AddNode.y) = 0 Or pMap.GetMapCharacter(AddNode.x, AddNode.y) = index) Then
                    AddNode.gval = BestNode.gval + ChrInfo(TextNumber).MoveCost(pMap.GetMapLayer(BestNode.x, BestNode.y))
                    AddNode.hval = Cost(AddNode.x, AddNode.y, pStartX, pStartY)
                    If MarkData(AddNode.x, AddNode.y).Set = 0 Then
                        MarkData(AddNode.x, AddNode.y).Cost = AddNode.gval + AddNode.hval
                        MarkData(AddNode.x, AddNode.y).Set = 1
                        MarkData(AddNode.x, AddNode.y).Direction = ReverseDirection(i)
                        OPENSet.Add AddNode, AddNode.gval + AddNode.hval, "R" & AddNode.x & "C" & AddNode.y
                    Else
                        If MarkData(AddNode.x, AddNode.y).Cost > AddNode.gval + AddNode.hval And MarkData(AddNode.x, AddNode.y).Set = 1 Then
                            OPENSet.Delete "R" & AddNode.x & "C" & AddNode.y, MarkData(AddNode.x, AddNode.y).Cost
                            MarkData(AddNode.x, AddNode.y).Cost = AddNode.gval + AddNode.hval
                            MarkData(AddNode.x, AddNode.y).Direction = ReverseDirection(i)
                            OPENSet.Add AddNode, AddNode.gval + AddNode.hval, "R" & AddNode.x & "C" & AddNode.y
                        Else
                        End If
                    End If
                End If
            End If
            Set AddNode = Nothing
        Next
        InsertDirection = InsertDirection * -1
        If OPENSet.Delete(CurrentNode.ItemKey, CurrentNode.Value) = False Then
            Debug.Print "Deletion of best node failed"
        End If
        MarkData(BestNode.x, BestNode.y).Set = 2
        Set AddNode = Nothing
        Set BestNode = Nothing
        Set CurrentNode = Nothing
    Loop
    CreateMarkData = OPENSet.Count > 0
    Set OPENSet = Nothing
End Function

Private Function ReverseDirection(pDir As Long) As Long
    Select Case pDir
        Case 0
            ReverseDirection = 3
        Case 1
            ReverseDirection = 2
        Case 2
            ReverseDirection = 1
        Case 3
            ReverseDirection = 0
    End Select
End Function

Public Sub DoAction()
Dim CheckIf As Boolean, arg(3) As Integer, i As Integer, damage As Integer

    If nCommands > 0 Then
        If CommandFinished Then
            Select Case ChrCommands(1).Action
                Case Is = Actions.Move
                    If Not ChrCommands(1).Arguement(2) = 0 Then
                        ChrCommands(1).Arguement(2) = ChrCommands(1).Arguement(2) - 1
                        ChrCommands(1).Arguement(3) = ChrCommands(1).Arguement(3) - 1
                        Exit Sub
                    End If
                    If Not ChrCommands(1).Arguement(3) = 0 Then ChrCommands(1).Arguement(3) = ChrCommands(1).Arguement(3) - 1: Exit Sub
                    DestX = ChrCommands(1).Arguement(0)
                    DestY = ChrCommands(1).Arguement(1)
                    CheckIf = CreateMarkData(MainMap, DestX, DestY, Me.x, Me.y, 0, 0, MainMap.MapX, MainMap.MapY)
                    If CheckIf Then
                        CreateControlPoints Me.x, Me.y, DestX, DestY
                        arg(0) = DestX: arg(1) = DestY
                    Else
                        DestX = x: DestY = y: RemoveCommand: Exit Sub
                    End If
                    CommandFinished = False
                    CheckingTileMove = 1
                    MovingTime = 1
            End Select
        Else
            Select Case ChrCommands(1).Action
                Case Is = Actions.Move
                    'Tile Moving
                    If CInt(MovingTime) > CheckingTileMove And Not CInt(MovingTime) >= UBound(ControlPoints) Then
                        For i = CheckingTileMove To CInt(MovingTime) - 1
                            Select Case MarkData(x, y).Direction
                                Case 0
                                    If Not MainMap.GetMapCharacter(x, y - 1) = 0 Then Me.ScreenX = CX(x + 0.5, y + 0.5) - TileSize / 2: Me.ScreenY = CY(x + 0.5, y + 0.5) - TileSize: RemoveCommand: arg(0) = DestX: arg(1) = DestY: AddCommand Actions.Move, arg: Exit Sub
                                    MainMap.MoveCharacter index, x, y - 1
                                    Direction = 0
                                Case 1
                                    If Not MainMap.GetMapCharacter(x - 1, y) = 0 Then Me.ScreenX = CX(x + 0.5, y + 0.5) - TileSize / 2: Me.ScreenY = CY(x + 0.5, y + 0.5) - TileSize: RemoveCommand: arg(0) = DestX: arg(1) = DestY: AddCommand Actions.Move, arg: Exit Sub
                                    MainMap.MoveCharacter index, x - 1, y
                                    Direction = 3
                                Case 2
                                    If Not MainMap.GetMapCharacter(x + 1, y) = 0 Then Me.ScreenX = CX(x + 0.5, y + 0.5) - TileSize / 2: Me.ScreenY = CY(x + 0.5, y + 0.5) - TileSize: RemoveCommand: arg(0) = DestX: arg(1) = DestY: AddCommand Actions.Move, arg: Exit Sub
                                    MainMap.MoveCharacter index, x + 1, y
                                    Direction = 2
                                Case 3
                                    If Not MainMap.GetMapCharacter(x, y + 1) = 0 Then Me.ScreenX = CX(x + 0.5, y + 0.5) - TileSize / 2: Me.ScreenY = CY(x + 0.5, y + 0.5) - TileSize: RemoveCommand: arg(0) = DestX: arg(1) = DestY: AddCommand Actions.Move, arg: Exit Sub
                                    MainMap.MoveCharacter index, x, y + 1
                                    Direction = 1
                            End Select
                        Next i
                        CheckingTileMove = CInt(MovingTime)
                    End If
                    'Screen Moving
                    If MovingTime > UBound(ControlPoints) - 1 Then
                        MovingTime = UBound(ControlPoints) - 1
                        FindScreenCoo
                        RemoveCommand
                        ReDim ControlPoints(0)
                        Exit Sub
                    Else
                        FindScreenCoo
                    End If
                    MovingTime = MovingTime + (1 / fps) * (Speed - ChrInfo(TextNumber).MoveCost(MainMap.GetMapLayer(x, y)))
                    TexFrame = (MovingTime - Int(MovingTime)) * 5
                    If CheckSeenList And Agressive Then
                        If Distance(MainMap.GetCharacter(DestAttack).x, MainMap.GetCharacter(DestAttack).y, x, y) <= Me.AttackRange Then
                            ChrCommands(1).Action = Attack
                            Exit Sub
                        End If
                    End If
                 Case Is = Actions.Attack
                    If DestAttack = 0 Then
                        AttackingTime = 0
                        ChrCommands(1).Action = Move
                        Exit Sub
                    End If
                    With MainMap.GetCharacter(DestAttack)
                        If .HealthLeft <= 0 Then
                            AttackingTime = 0
                            MainMap.AddToRemovingList DestAttack
                            DestAttack = 0
                            ChrCommands(1).Action = Move
                            Exit Sub
                        End If
                        If AttackingTime >= Me.AttackRate Then
                            damage = Me.AttackDamage - .Armor
                            If damage < 0 Then damage = 0
                            .HealthLeft = .HealthLeft - damage
                        Else
                            AttackingTime = AttackingTime + timeGetTime - tm
                        End If
                    End With
            End Select
        End If
    End If
End Sub

Public Sub AddCommand(CommandType As Actions, arguements() As Integer, Optional ToList As Boolean = False)
Dim i As Integer
    
    If Not ToList Then ClearCommands: CommandFinished = True
    i = nCommands + 1
    ReDim Preserve ChrCommands(i)
    ChrCommands(i).Action = CommandType
    Select Case ChrCommands(i).Action
        Case Is = Move
            ReDim ChrCommands(i).Arguement(3)
            With ChrCommands(i)
                .Arguement(0) = arguements(0)
                .Arguement(1) = arguements(1)
                .Arguement(2) = arguements(2)
                .Arguement(3) = arguements(3)
            End With
        Case Is = Attack
            ReDim ChrCommands(i).Arguement(1)
            With ChrCommands(i)
                .Arguement(0) = arguements(0)
                .Arguement(1) = arguements(1)
            End With
    End Select
End Sub

Private Sub ClearCommands()
    ReDim ChrCommands(0)
    ReDim ControlPoints(0)
End Sub

Private Sub RemoveCommand()
Dim i As Integer
    
    For i = 0 To nCommands - 1
        ChrCommands(i) = ChrCommands(i + 1)
    Next i
    ReDim Preserve ChrCommands(nCommands - 1)
    CommandFinished = True
    ReDim ControlPoints(0)
End Sub

Private Function nCommands() As Integer
    nCommands = UBound(ChrCommands)
End Function

Private Function FindScreenCoo()
Dim i As Double
Dim TotalBasis As Double, CurrentBasis As Double, p As Point
Dim Mini As Integer, Maxi As Integer
Dim tlx As Double, tly As Double
    
    Mini = Int(MovingTime - 2)
    Maxi = Int(MovingTime + 2)
    If Mini < 0 Then Mini = 0
    If Maxi > UBound(ControlPoints) Then Maxi = UBound(ControlPoints)
    For i = Mini To Maxi
        DoEvents
        CurrentBasis = FindBasis(MovingTime - i)
        p.x = p.x + ControlPoints(i).x * CurrentBasis
        p.y = p.y + ControlPoints(i).y * CurrentBasis
        TotalBasis = TotalBasis + CurrentBasis
    Next i
    p.x = p.x / TotalBasis
    p.y = p.y / TotalBasis
    tlx = CX(p.x / TileSize + 0.5, p.y / TileSize + 0.5) - TileSize / 2
    tly = CY(p.x / TileSize + 0.5, p.y / TileSize + 0.5) - TileSize
    ScreenX = tlx: ScreenY = tly
End Function

Private Sub CreateControlPoints(StrX As Integer, StrY As Integer, DesX As Integer, DesY As Integer)
Dim i As Integer, CurX As Double, CurY As Double

    MovingTime = 1
    CurX = StrX: CurY = StrY
    i = UBound(ControlPoints) + 1
    ReDim Preserve ControlPoints(i)
    With ControlPoints(i)
        .x = CurX * TileSize
        .y = CurY * TileSize
    End With
    Do
        i = UBound(ControlPoints) + 1
        If CurX < 0 Or CurY < 0 Then RemoveCommand:  Exit Sub
        ReDim Preserve ControlPoints(i)
        Select Case MarkData(CurX, CurY).Direction
            Case 0
                CurX = CurX: CurY = (CurY - 1)
            Case 1
                CurX = (CurX - 1): CurY = CurY
            Case 2
                CurX = (CurX + 1): CurY = CurY
            Case 3
                CurX = CurX: CurY = (CurY + 1)
        End Select
        With ControlPoints(i)
            .x = CurX * TileSize
            .y = CurY * TileSize
        End With
        If CurX = DesX And CurY = DesY Then
            ReDim Preserve ControlPoints(i + 1)
            With ControlPoints(0)
                .x = ControlPoints(1).x * 2 - ControlPoints(2).x
                .y = ControlPoints(1).y * 2 - ControlPoints(2).y
            End With
            With ControlPoints(i + 1)
                .x = ControlPoints(i).x * 2 - ControlPoints(i - 1).x
                .y = ControlPoints(i).y * 2 - ControlPoints(i - 1).y
            End With
            Exit Do
        End If
    Loop
End Sub
    
Public Function FindBasis(t As Double) As Double
    Select Case t
        Case Is < -2: FindBasis = 0
        Case Is <= -1: FindBasis = ((2 + t) ^ 3) / 6
        Case Is <= 0: FindBasis = (4 - 6 * t ^ 2 - 3 * t ^ 3) / 6
        Case Is <= 1: FindBasis = (4 - 6 * t ^ 2 + 3 * t ^ 3) / 6
        Case Is <= 2: FindBasis = ((2 - t) ^ 3) / 6
        Case Else: FindBasis = 0
    End Select
End Function

Public Function SeenFromPlayerNum(numb As Integer) As Boolean
Dim i As Integer
    
    SeenFromPlayerNum = False
    For i = 1 To nSeenFromList
        If MainMap.GetCharacter(SeenFromList(i)).Nplayer = numb Then SeenFromPlayerNum = True: Exit Function
    Next i
End Function

Private Function CheckSeenList() As Boolean
Dim i As Integer
    
    CheckSeenList = False
    For i = 1 To nSeenList
        If Not MainMap.GetCharacter(SeenList(i)).Nplayer = Me.Nplayer Then
            DestAttack = SeenList(i)
            CheckSeenList = True
            Exit Function
        End If
    Next i
End Function

Public Function GetButton(index As Integer) As ButtonList
    GetButton = Buttons(index)
End Function
