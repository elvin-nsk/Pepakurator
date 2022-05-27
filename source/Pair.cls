VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Number As Long
  Shape1 As Shape
  Shape2 As Shape
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Shape1 As Shape, _
                       ByVal Shape2 As Shape _
                       ) As Pair
  With New Pair
    Set Create = .Init(Shape1, Shape2)
  End With
End Function

Friend Function Init(ByVal Shape1 As Shape, _
                     ByVal Shape2 As Shape _
                     ) As Pair
  With This
    Set .Shape1 = Shape1
    Set .Shape2 = Shape2
    If VBA.IsNumeric(Shape1.Text.Story.Text) Then
      .Number = VBA.CLng(Shape1.Text.Story.Text)
    End If
  End With
  Set Init = Me
End Function

'===============================================================================

Public Function IsSame(ByVal Pair As Pair) As Boolean
  IsSame = IsShapesSameAsThis(Pair.Shape1, Pair.Shape2)
End Function

Public Property Get Number() As Long
  Number = This.Number
End Property
Public Property Let Number(ByVal RHS As Long)
  With This
    .Number = RHS
    If .Number = 0 Then
      Common.SetTextShapeToPlaceholder .Shape1
      Common.SetTextShapeToPlaceholder .Shape2
    Else
      .Shape1.Text.Story.Text = VBA.CStr(.Number)
      .Shape2.Text.Story.Text = VBA.CStr(.Number)
    End If
  End With
End Property

Public Property Get Shape1() As Shape
  Set Shape1 = This.Shape1
End Property

Public Property Get Shape2() As Shape
  Set Shape2 = This.Shape2
End Property

Public Property Get Self() As Pair
  Set Self = Me
End Property

'===============================================================================

Private Function IsShapesSameAsThis(ByVal Shape1 As Shape, Shape2 As Shape) As Boolean
  With This
     IsShapesSameAsThis = ((lib_elvin.IsSameShapes(Shape1, .Shape1) And lib_elvin.IsSameShapes(Shape2, .Shape2)) Or _
                           (lib_elvin.IsSameShapes(Shape1, .Shape2) And lib_elvin.IsSameShapes(Shape2, .Shape1)))
  End With
End Function
