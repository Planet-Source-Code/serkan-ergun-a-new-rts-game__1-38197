VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HeapNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarValue As Single
Private mvarLeftChild As HeapNode
Private mvarRightChild As HeapNode
Private mvarParentNode As HeapNode
Private mvarItem As Variant
Private mvarItemKey As String

Public Property Let ItemKey(ByVal vData As String)
    mvarItemKey = vData
End Property

Public Property Get ItemKey() As String
    ItemKey = mvarItemKey
End Property

Public Property Set Item(ByVal vData As Variant)
    Set mvarItem = vData
End Property

Public Property Let Item(ByVal vData As Variant)
    mvarItem = vData
End Property

Public Property Get Item() As Variant
    Set Item = mvarItem
End Property

Public Property Set ParentNode(ByVal vData As HeapNode)
    Set mvarParentNode = vData
End Property

Public Property Get ParentNode() As HeapNode
    Set ParentNode = mvarParentNode
End Property

Public Property Set RightChild(ByVal vData As HeapNode)
    Set mvarRightChild = vData
End Property

Public Property Get RightChild() As HeapNode
    Set RightChild = mvarRightChild
End Property

Public Property Set LeftChild(ByVal vData As HeapNode)
    Set mvarLeftChild = vData
End Property

Public Property Get LeftChild() As HeapNode
    Set LeftChild = mvarLeftChild
End Property

Public Property Let Value(ByVal vData As Single)
    mvarValue = vData
End Property

Public Property Get Value() As Single
    Value = mvarValue
End Property
