VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class modules
Private cIndex As Integer
Private cX As Integer
Private cY As Integer
Private cNextNode As Node

Public Property Get Index() As Integer
    Index = cIndex
End Property

Public Property Get X() As Integer
    X = cX
End Property

Public Property Get Y() As Integer
    Y = cY
End Property

Public Property Get NextNode() As Node
    Set NextNode = cNextNode
End Property

Public Property Let Index(ByVal Index As Integer)
    cIndex = Index
End Property

Public Property Let X(ByVal X As Integer)
    cX = X
End Property

Public Property Let Y(ByVal Y As Integer)
    cY = Y
End Property

Public Property Let NextNode(ByVal NextNode As Node)
    Set cNextNode = NextNode
End Property
