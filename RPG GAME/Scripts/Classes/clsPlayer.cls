VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Pos As Position
Private Direct As Position



Public Sub SetPosition(xValue, yValue)
    Pos.x = xValue
    Pos.Y = yValue
End Sub

Property Get Position() As Position
    Position.x = Pos.x
    Position.Y = Pos.Y
End Property

Public Sub SetDirection(xValue, yValue)
    Direct.x = xValue
    Direct.Y = yValue
End Sub

Property Get Direction() As Position
    Direction.x = Direct.x
    Direction.Y = Direct.Y
End Property
