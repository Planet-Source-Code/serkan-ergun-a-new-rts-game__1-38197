VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AStarNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarY As Long
Private mvarX As Long
Private mvargval As Single
Private mvarhval As Single

Public Property Let hval(ByVal vData As Single)
    mvarhval = vData
End Property

Public Property Get hval() As Single
    hval = mvarhval
End Property

Public Property Let gval(ByVal vData As Single)
    mvargval = vData
End Property

Public Property Get gval() As Single
    gval = mvargval
End Property

Public Property Let x(ByVal vData As Long)
    mvarX = vData
End Property

Public Property Get x() As Long
    x = mvarX
End Property

Public Property Let y(ByVal vData As Long)
    mvarY = vData
End Property

Public Property Get y() As Long
    y = mvarY
End Property

