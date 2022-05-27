VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DashLinesParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Const DashCountForLongDash As Long = 1
Private Const DashCountForTwoDash As Long = 2

Private Type typeThis
  Page As Page
  LongDashLines As ShapeRange
  TwoDashLines As ShapeRange
  ShapesToAvoid As ShapeRange
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Page As Page) As DashLinesParser
  With New DashLinesParser
    Set Create = .Init(Page)
  End With
End Function

Friend Function Init(ByVal Page As Page) As DashLinesParser
  With This
    Set .Page = Page
    Set .LongDashLines = CreateShapeRange
    Set .TwoDashLines = CreateShapeRange
    Set .ShapesToAvoid = CreateShapeRange
  End With
  Parse
  Set Init = Me
End Function

'===============================================================================

Public Property Get LongDashLines() As ShapeRange
  Set LongDashLines = This.LongDashLines
End Property

Public Property Get TwoDashLines() As ShapeRange
  Set TwoDashLines = This.TwoDashLines
End Property

Public Property Get ShapesToAvoid() As ShapeRange
  Set ShapesToAvoid = This.ShapesToAvoid
End Property

Public Property Get Self() As DashLinesParser
  Set Self = Me
End Property

'===============================================================================

Private Sub Parse()
  With This
    Dim Shape As Shape
    For Each Shape In .Page.Shapes.FindShapes
      Select Case True
        Case IsLongDash(Shape)
          .LongDashLines.Add Shape
        Case IsTwoDash(Shape)
          .TwoDashLines.Add Shape
        Case ShapeToAvoid(Shape)
          .ShapesToAvoid.Add Shape
      End Select
    Next Shape
  End With
End Sub

Private Function IsLongDash(ByVal Shape As Shape) As Boolean
  If Not HaveTwoNodes(Shape) Then Exit Function
  IsLongDash = (Shape.Outline.Style.DashCount = DashCountForLongDash)
End Function

Private Function IsTwoDash(ByVal Shape As Shape) As Boolean
  If Not HaveTwoNodes(Shape) Then Exit Function
  IsTwoDash = (Shape.Outline.Style.DashCount = DashCountForTwoDash)
End Function

Private Function ShapeToAvoid(ByVal Shape As Shape) As Boolean
  ShapeToAvoid = (Shape.Type = cdrTextShape)
End Function

Private Function HaveTwoNodes(ByVal Shape As Shape) As Boolean
  On Error GoTo Finally
  If Shape.Curve Is Nothing Then Exit Function
  HaveTwoNodes = (Shape.Curve.Nodes.Count = 2)
Finally:
  On Error GoTo 0
End Function
