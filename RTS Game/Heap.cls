VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Heap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarTopNode As HeapNode
Private mvarCount As Long
Private mvarTreeHeight As Long
Private mvarLeftOrRightDeletion As Long

Private Property Get LeftOrRightDeletion() As Long
    LeftOrRightDeletion = mvarLeftOrRightDeletion
    mvarLeftOrRightDeletion = mvarLeftOrRightDeletion Xor 1
End Property

Private Property Let LeftOrRightDeletion(ByVal vData As Long)
    mvarLeftOrRightDeletion = vData
End Property

Private Property Let TreeHeight(ByVal vData As Long)
    mvarTreeHeight = vData
End Property

Public Property Get TreeHeight() As Long
    TreeHeight = mvarTreeHeight
End Property

Private Property Let Count(ByVal vData As Long)
    mvarCount = vData
End Property

Public Property Get Count() As Long
    Count = mvarCount
End Property

Public Property Set TopNode(ByVal vData As HeapNode)
    Set mvarTopNode = vData
End Property

Public Property Get TopNode() As HeapNode
    Set TopNode = mvarTopNode
End Property

Private Sub Class_Initialize()
    Set TopNode = Nothing
    TreeHeight = 0
    Count = 0
    LeftOrRightDeletion = 0
End Sub

Private Sub Class_Terminate()
    DestroyHeap
End Sub

Public Function Delete(ByVal ItemKey As String, ByVal Value As Single) As Boolean
Dim NodeToDelete As HeapNode
Dim ParentNode As HeapNode
Dim FoundAtDepth As Long
Dim ChildNode As HeapNode
Dim SeperateTree As HeapNode
Dim ReinsertNode As HeapNode
Dim LeftOrRightChild As Long
Dim DepthFound As Long

    Set NodeToDelete = FindInHeap(Value, ItemKey)
    If NodeToDelete Is Nothing Then
        Delete = False
        Exit Function
    End If
    Delete = True
    If NodeToDelete.LeftChild Is Nothing Then
        If NodeToDelete.RightChild Is Nothing Then
            If NodeToDelete.ParentNode Is Nothing Then
                Set NodeToDelete = Nothing
                Set TopNode = Nothing
            Else
                Set ParentNode = NodeToDelete.ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then
                    Set ParentNode.LeftChild = Nothing
                Else
                    Set ParentNode.RightChild = Nothing
                End If
                Set NodeToDelete = Nothing
            End If
        Else
            If NodeToDelete.ParentNode Is Nothing Then
                Set TopNode = NodeToDelete.RightChild
                Set NodeToDelete.RightChild.ParentNode = Nothing
                Set NodeToDelete = Nothing
            Else
                Set ParentNode = NodeToDelete.ParentNode
                Set NodeToDelete.RightChild.ParentNode = ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then
                    Set ParentNode.LeftChild = NodeToDelete.RightChild
                Else
                    Set ParentNode.RightChild = NodeToDelete.RightChild
                End If
                Set NodeToDelete = Nothing
            End If
        End If
    Else
        If NodeToDelete.RightChild Is Nothing Then
            If NodeToDelete.ParentNode Is Nothing Then
                Set TopNode = NodeToDelete.LeftChild
                Set NodeToDelete.LeftChild.ParentNode = Nothing
                Set NodeToDelete = Nothing
            Else
                Set ParentNode = NodeToDelete.ParentNode
                Set NodeToDelete.LeftChild.ParentNode = ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then
                    Set ParentNode.LeftChild = NodeToDelete.LeftChild
                Else
                    Set ParentNode.RightChild = NodeToDelete.LeftChild
                End If
                Set NodeToDelete.LeftChild.ParentNode = ParentNode
                Set NodeToDelete = Nothing
            End If
        Else
            If LeftOrRightDeletion = 0 Then
                Set ChildNode = NodeToDelete.LeftChild
                Set SeperateTree = ChildNode.RightChild
                If Not (SeperateTree Is Nothing) Then
                    Set SeperateTree.ParentNode = Nothing
                End If
                Set ChildNode.RightChild = NodeToDelete.RightChild
                Set ChildNode.RightChild.ParentNode = ChildNode
                If NodeToDelete.ParentNode Is Nothing Then
                    Set ChildNode.ParentNode = Nothing
                    Set TopNode = ChildNode
                Else
                    Set ChildNode.ParentNode = NodeToDelete.ParentNode
                    If NodeToDelete.Value <= NodeToDelete.ParentNode.Value Then
                        Set NodeToDelete.ParentNode.LeftChild = ChildNode
                    Else
                        Set NodeToDelete.ParentNode.RightChild = ChildNode
                    End If
                End If
                Set NodeToDelete = Nothing
                If Not (SeperateTree Is Nothing) Then
                    Set ReinsertNode = FindFreeSpace(SeperateTree.Value, LeftOrRightChild, DepthFound)
                    Set SeperateTree.ParentNode = ReinsertNode
                    If LeftOrRightChild = 0 Then
                        Set ReinsertNode.LeftChild = SeperateTree
                    Else
                        Set ReinsertNode.RightChild = SeperateTree
                    End If
                End If
            Else
                Set ChildNode = NodeToDelete.RightChild
                Set SeperateTree = ChildNode.LeftChild
                If Not (SeperateTree Is Nothing) Then
                    Set SeperateTree.ParentNode = Nothing
                End If
                Set ChildNode.LeftChild = NodeToDelete.LeftChild
                Set ChildNode.LeftChild.ParentNode = ChildNode
                If NodeToDelete.ParentNode Is Nothing Then
                    Set ChildNode.ParentNode = Nothing
                    Set TopNode = ChildNode
                Else
                    Set ChildNode.ParentNode = NodeToDelete.ParentNode
                    If NodeToDelete.Value <= NodeToDelete.ParentNode.Value Then
                        Set NodeToDelete.ParentNode.LeftChild = ChildNode
                    Else
                        Set NodeToDelete.ParentNode.RightChild = ChildNode
                    End If
                End If
                Set NodeToDelete = Nothing
                If Not (SeperateTree Is Nothing) Then
                    Set ReinsertNode = FindFreeSpace(SeperateTree.Value, LeftOrRightChild, DepthFound)
                    Set SeperateTree.ParentNode = ReinsertNode
                    If LeftOrRightChild = 0 Then
                        Set ReinsertNode.LeftChild = SeperateTree
                    Else
                        Set ReinsertNode.RightChild = SeperateTree
                    End If
                End If
            End If
        End If
    End If
    Count = Count - 1
End Function

Public Function Add(Obj As Variant, Value As Single, ItemKey As String) As HeapNode
Dim objHeapNode As New HeapNode
Dim objParentNode As HeapNode
Dim LeftOrRight As Long
Dim InsertionDepth As Long
    
    With objHeapNode
        .ItemKey = ItemKey
        If IsObject(Obj) Then
            Set .Item = Obj
        Else
            .Item = Obj
        End If
        .Value = Value
        Set .LeftChild = Nothing
        Set .RightChild = Nothing
        Set .ParentNode = Nothing
    End With

    If TopNode Is Nothing Then
        Set TopNode = objHeapNode
        TreeHeight = 1
    Else
        Set objParentNode = FindFreeSpace(Value, LeftOrRight, InsertionDepth)
        With objHeapNode
            Set .ParentNode = objParentNode
        End With
        With objParentNode
            If LeftOrRight = 0 Then
                Set objParentNode.LeftChild = objHeapNode
            Else
                Set objParentNode.RightChild = objHeapNode
            End If
        End With
        If InsertionDepth > TreeHeight Then
            TreeHeight = InsertionDepth
        End If
    End If
    Count = Count + 1
End Function

Public Function FindInHeap(Value As Single, ItemKey As String) As HeapNode
Dim CurrentNode As HeapNode
Dim FoundIt As Boolean

    Set CurrentNode = TopNode
    FoundIt = False
    Do While FoundIt = False
        If Value <= CurrentNode.Value Then
            If Value = CurrentNode.Value And ItemKey = CurrentNode.ItemKey Then
                Set FindInHeap = CurrentNode
                FoundIt = True
            Else
                If CurrentNode.LeftChild Is Nothing Then
                    Set FindInHeap = Nothing
                    FoundIt = True
                Else
                    Set CurrentNode = CurrentNode.LeftChild
                End If
            End If
        Else
            If CurrentNode.RightChild Is Nothing Then
                Set FindInHeap = Nothing
                FoundIt = True
            Else
                Set CurrentNode = CurrentNode.RightChild
            End If
        End If
    Loop
End Function

Public Function GetLeftMostElement() As HeapNode
Dim CurrentNode As HeapNode
    Set CurrentNode = TopNode
    If TopNode Is Nothing Then
        Set GetLeftMostElement = Nothing
        Exit Function
    End If
    Do While Not (CurrentNode.LeftChild Is Nothing)
        Set CurrentNode = CurrentNode.LeftChild
    Loop
    Set GetLeftMostElement = CurrentNode
End Function

Public Function FindFreeSpace(ByRef Value As Single, ByRef LeftOrRight As Long, ByRef DepthFound As Long) As HeapNode
Dim CurrentNode As HeapNode
Dim FoundFree As Boolean
Dim SearchDepth As Long
    
    SearchDepth = 1
    Set CurrentNode = TopNode
    FoundFree = False
    Do While FoundFree = False
        If Value <= CurrentNode.Value Then
            If CurrentNode.LeftChild Is Nothing Then
                FoundFree = True
                LeftOrRight = 0
            Else
                Set CurrentNode = CurrentNode.LeftChild
            End If
        Else
            If CurrentNode.RightChild Is Nothing Then
                FoundFree = True
                LeftOrRight = 1
            Else
                Set CurrentNode = CurrentNode.RightChild
            End If
        End If
        SearchDepth = SearchDepth + 1
    Loop
    Set FindFreeSpace = CurrentNode
    DepthFound = SearchDepth
End Function

Public Function DestroyHeap() As Boolean
Dim CurrentNode As HeapNode
Dim ParentNode As HeapNode
    Set CurrentNode = TopNode
    Do While Count > 0
        DoEvents
        If CurrentNode.LeftChild Is Nothing Then
            If CurrentNode.RightChild Is Nothing Then
                If CurrentNode.ParentNode Is Nothing Then
                    Set CurrentNode = Nothing
                    Set TopNode = Nothing
                Else
                    Set ParentNode = CurrentNode.ParentNode
                    If CurrentNode.Value <= ParentNode.Value Then
                        Set ParentNode.LeftChild = Nothing
                    Else
                        Set ParentNode.RightChild = Nothing
                    End If
                    Set CurrentNode = Nothing
                    Set CurrentNode = ParentNode
                End If
                Count = Count - 1
            Else
                Set CurrentNode = CurrentNode.RightChild
            End If
        Else
            Set CurrentNode = CurrentNode.LeftChild
        End If
    Loop
End Function
